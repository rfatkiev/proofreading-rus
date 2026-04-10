import os
import re
import sys
import time

import win32com.client  # pywin32


WORD_TOKEN_RE = re.compile(r"[A-Za-zА-Яа-яЁё]+(?:-[A-Za-zА-Яа-яЁё]+)*")
SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?])\s+|\n+")
SPACED_THOUSANDS_RE = re.compile(r"\s(\d) (\d{3})(?!\d)")


WD_FIND_CONTINUE = 1
WD_REPLACE_ALL = 2
WD_COLLAPSE_END = 0
WD_HIGHLIGHT_YELLOW = 7


def is_capitalized(token):
    return len(token) >= 2 and token[:1].isupper() and not token.isupper()


def is_all_caps_line(text):
    letters = re.findall(r"[A-Za-zА-Яа-яЁё]", text)
    if not letters:
        return False
    return all(ch.isupper() for ch in letters)


def extract_proper_nouns_from_text(text, proper_phrases, proper_words):
    for sentence in SENTENCE_SPLIT_RE.split(text):
        tokens = WORD_TOKEN_RE.findall(sentence)
        i = 0
        while i < len(tokens):
            token = tokens[i]
            if is_capitalized(token):
                seq = [token]
                j = i + 1
                while j < len(tokens) and is_capitalized(tokens[j]):
                    seq.append(tokens[j])
                    j += 1

                if len(seq) >= 2:
                    phrase = " ".join(seq)
                    proper_phrases.add(phrase)
                    for w in seq:
                        proper_words.add(w)
                    i = j
                    continue

                if i != 0:
                    proper_phrases.add(token)
                    proper_words.add(token)
                i += 1
                continue

            i += 1


def search_replace_all(doc, find_text, replace_text, match_case=False, match_whole_word=False, match_wildcards=False):
    app = doc.Application
    doc.Content.Select()
    find = app.Selection.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    return find.Execute(
        find_text,
        match_case,
        match_whole_word,
        match_wildcards,
        False,
        False,
        True,
        WD_FIND_CONTINUE,
        False,
        replace_text,
        WD_REPLACE_ALL,
    )


def replace_ellipsis(doc):
    search_replace_all(doc, "...", "\u2026")


def replace_bracketed_ellipsis(doc):
    search_replace_all(doc, "[...]", "<\u2026>")
    search_replace_all(doc, "[\u2026]", "<\u2026>")


def replace_double_and_triple_spaces(doc):
    search_replace_all(doc, "   ", " ")
    search_replace_all(doc, "  ", " ")


def fix_quote_spacing(doc):
    search_replace_all(doc, "« ", "«")
    search_replace_all(doc, " »", "»")


def fix_dashes(doc):
    for para in doc.Paragraphs:
        rng = para.Range
        replacements = [
            (r"(?<=\d)[-—](?=\d)", "–"),
            (r"(?<=[A-Za-zА-Яа-яЁё])[—–](?=[A-Za-zА-Яа-яЁё])", "-"),
            (r"(?<=[A-Za-zА-Яа-яЁё])—", " —"),
            (r"—(?=[A-Za-zА-Яа-яЁё])", "— "),
        ]
        for pattern, repl in replacements:
            _replace_regex_in_range(rng, pattern, repl)


def replace_spaced_thousands(doc):
    text = doc.Content.Text
    matches = list(SPACED_THOUSANDS_RE.finditer(text))
    if not matches:
        return 0

    total = len(matches)
    unique_values = sorted({m.group(0) for m in matches})
    for value in unique_values:
        search_replace_all(doc, value, value.replace(" ", ""))
    return total


def fix_links_spacing(doc):
    for para in doc.Paragraphs:
        _replace_regex_in_range(para.Range, r"([A-Za-zА-Яа-яЁё0-9])(\[[0-9]{1,}\])", r"\1 \2")


def replace_percent_words(doc):
    search_replace_all(doc, "проценты", "%", match_whole_word=True)


def ensure_space_before_percent(doc):
    for para in doc.Paragraphs:
        _replace_regex_in_range(para.Range, r"([0-9])%", r"\1 %")


def remove_italic_from_brackets(doc):
    return


def fix_abbreviations(doc):
    search_replace_all(doc, "и т.д.", "и т. д.")
    search_replace_all(doc, "и т.п.", "и т. п.")
    search_replace_all(doc, "и т.д", "и т. д.")
    search_replace_all(doc, "и т.п", "и т. п.")

    search_replace_all(doc, "И т.д.", "И т. д.")
    search_replace_all(doc, "И т.п.", "И т. п.")
    search_replace_all(doc, "И т.д", "И т. д.")
    search_replace_all(doc, "И т.п", "И т. п.")

    search_replace_all(doc, "т.е.", "то есть")
    search_replace_all(doc, "Т.е.", "То есть")

    search_replace_all(doc, "и пр.", "и прочее")
    search_replace_all(doc, "И пр.", "И прочее")


def fix_quotes_language(doc):
    return


def _replace_regex_in_range(rng, pattern, repl):
    text = rng.Text
    matches = list(re.finditer(pattern, text))
    if not matches:
        return 0

    replaced = 0
    base = rng.Start
    for match in reversed(matches):
        start = base + match.start(0)
        end = base + match.end(0)
        sub = rng.Duplicate
        sub.SetRange(start, end)
        sub.Text = match.expand(repl)
        replaced += 1
    return replaced


def _set_char_at(rng, pos, ch):
    sub = rng.Duplicate
    sub.SetRange(pos, pos + 1)
    sub.Text = ch




def mark_image_captions(para_range):
    text = para_range.Text
    if re.search(r"\bИлл\.\s*\d+\b", text):
        para_range.HighlightColorIndex = WD_HIGHLIGHT_YELLOW
        return True
    return False


def remove_leading_glava(para_range):
    text = para_range.Text
    match = re.match(r"\s*Глава\b\s*", text, flags=re.IGNORECASE)
    if not match:
        return False
    start = para_range.Start + match.start()
    end = para_range.Start + match.end()
    sub = para_range.Duplicate
    sub.SetRange(start, end)
    sub.Text = ""
    return True


def replace_yo_in_range(word_range, proper_nouns_lower):
    replaced = 0
    for rng in word_range.Words:
        word = rng.Text
        if "\u0451" not in word and "\u0401" not in word:
            continue
        stripped = re.sub(r"[^A-Za-zА-Яа-яЁё-]", "", word)
        if not stripped:
            continue
        if stripped.lower() in proper_nouns_lower:
            continue
        updated = word.replace("\u0451", "\u0435").replace("\u0401", "\u0415")
        if updated != word:
            rng.Text = updated
            replaced += 1
    return replaced


def main():
    start_time = time.perf_counter()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if len(sys.argv) > 1:
        doc_path = os.path.join(script_dir, sys.argv[1])
    else:
        docx_files = [f for f in os.listdir(script_dir) if f.lower().endswith(".docx")]
        if not docx_files:
            print("No .docx files found in the script directory.")
            sys.exit(1)
        doc_path = os.path.join(script_dir, docx_files[0])

    if not os.path.isfile(doc_path):
        print(f"File not found: {doc_path}")
        sys.exit(1)

    print(f"Opening: {doc_path}")

    app = win32com.client.Dispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    app.UserName = "python proofreading"

    doc = app.Documents.Open(doc_path)
    try:
        doc.TrackRevisions = True

        print("Replacing ... -> …")
        replace_ellipsis(doc)
        print("Replacing [...] or […] -> <…>")
        replace_bracketed_ellipsis(doc)
        print("Replacing double/triple spaces -> single space")
        replace_double_and_triple_spaces(doc)
        print("Fixing spaces after « and before »")
        fix_quote_spacing(doc)
        print("Fixing dashes")
        fix_dashes(doc)
        print("Fixing spaced thousands (including 1 000)")
        spaced_count = replace_spaced_thousands(doc)
        print(f"Spaced thousands replaced: {spaced_count}")
        print("Fixing spacing before [1] links")
        fix_links_spacing(doc)
        print("Replacing word 'проценты' -> %")
        replace_percent_words(doc)
        print("Ensuring space before %")
        ensure_space_before_percent(doc)
        print("Skipping italic formatting removal for () and []")
        print("Fixing abbreviations")
        fix_abbreviations(doc)
        print("Fixing quotes for Latin/Cyrillic")
        fix_quotes_language(doc)

        print("Collecting proper nouns + replacing ё -> е (streaming paragraphs)...")
        proper_phrases = set()
        proper_words = set()
        total_yo = 0
        total_paras = doc.Paragraphs.Count
        captions_marked = 0
        glava_removed = 0

        for idx, para in enumerate(doc.Paragraphs, start=1):
            text = para.Range.Text
            if not is_all_caps_line(text):
                extract_proper_nouns_from_text(text, proper_phrases, proper_words)

            proper_nouns_lower = {w.lower() for w in proper_words}
            total_yo += replace_yo_in_range(para.Range, proper_nouns_lower)

            if mark_image_captions(para.Range):
                captions_marked += 1
            if remove_leading_glava(para.Range):
                glava_removed += 1

            if idx % 50 == 0 or idx == total_paras:
                print(f"Processed paragraphs: {idx}/{total_paras}")

        print(f"Proper nouns found: {len(proper_phrases)}")
        print(f"Replaced ё/Ё in {total_yo} word(s)")
        print(f"Captions marked: {captions_marked}")
        print(f"Removed 'Глава' at paragraph start: {glava_removed}")

        base, ext = os.path.splitext(doc_path)
        reviewed_path = f"{base}_reviewed{ext}"
        doc.SaveAs2(reviewed_path)

        nouns_path = os.path.join(script_dir, "proper_nouns.txt")
        with open(nouns_path, "w", encoding="utf-8") as f:
            for phrase in sorted(proper_phrases, key=lambda w: w.lower()):
                f.write(phrase + "\n")
    finally:
        doc.Close(SaveChanges=False)
        app.Quit()

    elapsed = time.perf_counter() - start_time
    print(f"Reviewed file saved to: {reviewed_path}")
    print(f"Proper nouns list saved to: {nouns_path}")
    print(f"Total time: {elapsed:.2f} s")


if __name__ == "__main__":
    main()
