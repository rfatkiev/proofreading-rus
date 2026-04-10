# Python for Proofreading

[ENG](#python-for-proofreading) | [RUS](#python-для-литературной-вычитки)

English:
[Functions](#what-the-script-does) | [Algorithm](#how-the-algorithm-works) | [Usage](#usage) | [Notes](#notes)

Русский:
[Функции](#что-делает-скрипт) | [Алгоритм](#как-работает-алгоритм) | [Использование](#использование) | [Примечания](#примечания)

This project contains a Word-based proofreading helper for Russian editorial work. The main script opens a `.docx` file in Microsoft Word through `pywin32`, enables Track Changes, applies a set of technical fixes, and saves the result as a new reviewed document.

The current main script is `review_word.py`. It is intended for repetitive, mechanical cleanup before deeper editing.

## What The Script Does

- Opens the first `.docx` file in the project folder, or a file passed as a command-line argument.
- Turns on Word review mode with the reviewer name `python proofreading`.
- Replaces `...` with the ellipsis character `…`.
- Replaces `[...]` and `[…]` with `<…>`.
- Collapses triple and double spaces into one space.
- Removes spaces after `«` and before `»`.
- Normalizes dash usage in several basic cases:
  - digit-dash-digit -> en dash `–`
  - letter-dash-letter -> hyphen `-`
  - stuck em dashes near letters -> spaced em dash ` — `
- Adds a space before reference markers like `[1]` when they are attached to a word.
- Replaces the exact word `проценты` with `%`.
- Adds a space before `%` when it is attached to a digit.
- Normalizes several abbreviations:
  - `и т.д.` -> `и т. д.`
  - `и т.п.` -> `и т. п.`
  - `т.е.` -> `то есть`
  - `и пр.` -> `и прочее`
- Extracts proper nouns into a separate file.
- Replaces `ё` with `е`, except for detected proper nouns.
- Highlights image-caption-like lines such as `Илл. 1. ...`.
- Removes the word `Глава` from the beginning of a paragraph.
- Prints progress and total runtime in the terminal.

## How The Algorithm Works

The script uses a two-stage workflow.

Stage 1: global technical cleanup
- Word `Find/Replace` is used for fast, simple fixes such as ellipses, spaces, and abbreviation cleanup.
- Some context-sensitive rules are applied paragraph by paragraph with Python regular expressions, because Word pattern matching was unreliable for several cases.

Stage 2: streaming paragraph pass
- The script iterates over all paragraphs in the document.
- It collects candidate proper nouns from capitalized word sequences.
- It skips all-caps lines when building the proper-noun list.
- It replaces `ё` only in words that are not recognized as proper nouns.
- It marks likely illustration captions.
- It removes leading `Глава` when found at the start of a paragraph.

At the end, the script saves:
- a reviewed `.docx` copy with tracked changes
- `proper_nouns.txt` with the extracted names

## Usage

Install dependency:

```powershell
pip install pywin32
```

Run on the first `.docx` in the folder:

```powershell
python review_word.py
```

Run on a specific file:

```powershell
python review_word.py "MyDocument.docx"
```

## Notes

- Microsoft Word must be installed on Windows.
- The script edits through Word COM automation, so performance depends partly on Word itself.
- Some planned rules are intentionally disabled right now because they were unstable or too slow in real documents.

---

# Python для литературной вычитки

Этот проект содержит вспомогательный парсер для технической вычитки Word-документов. Основной скрипт открывает `.docx` через Microsoft Word и `pywin32`, включает режим исправлений, применяет набор механических правок и сохраняет результат как новый файл.

Основной файл проекта: `review_word.py`. Скрипт рассчитан на черновую техническую чистку перед более глубокой редактурой.

## Что Делает Скрипт

- Открывает первый `.docx` в папке проекта или файл, переданный аргументом командной строки.
- Включает режим исправлений Word с именем редактора `python proofreading`.
- Заменяет `...` на знак многоточия `…`.
- Заменяет `[...]` и `[…]` на `<…>`.
- Убирает тройные и двойные пробелы.
- Убирает пробел после `«` и перед `»`.
- Нормализует тире в нескольких базовых случаях:
  - цифра-тире-цифра -> среднее тире `–`
  - буква-тире-буква -> дефис `-`
  - прилипшее длинное тире рядом с буквами -> длинное тире с пробелами
- Добавляет пробел перед ссылками вида `[1]`, если они слиты со словом.
- Заменяет точную словоформу `проценты` на `%`.
- Добавляет пробел перед `%`, если знак процента прилип к цифре.
- Нормализует несколько сокращений:
  - `и т.д.` -> `и т. д.`
  - `и т.п.` -> `и т. п.`
  - `т.е.` -> `то есть`
  - `и пр.` -> `и прочее`
- Сохраняет список имён собственных в отдельный файл.
- Меняет `ё` на `е`, кроме случаев, когда слово распознано как имя собственное.
- Подсвечивает строки, похожие на подписи к иллюстрациям, например `Илл. 1. ...`.
- Удаляет слово `Глава` в начале абзаца.
- Печатает прогресс и общее время работы в терминал.

## Как Работает Алгоритм

Скрипт работает в два этапа.

Этап 1: глобальная техническая чистка
- Для быстрых и простых замен используется Word `Find/Replace`: многоточия, пробелы, часть сокращений.
- Более чувствительные к контексту правила выполняются поабзацно через Python-регулярные выражения, потому что встроенный поиск Word оказался ненадёжным для ряда шаблонов.

Этап 2: потоковый проход по абзацам
- Скрипт проходит по всем абзацам документа.
- Из последовательностей слов с заглавной буквы собирает кандидатов в имена собственные.
- Строки, полностью набранные капсом, исключаются из этого списка.
- Замена `ё` выполняется только в словах, которые не попали в список имён собственных.
- Дополнительно подсвечиваются возможные подписи к иллюстрациям.
- Если в начале абзаца найдено слово `Глава`, оно удаляется.

В конце скрипт сохраняет:
- новую `.docx`-копию с исправлениями в режиме рецензирования
- файл `proper_nouns.txt` со списком найденных имён

## Использование

Установка зависимости:

```powershell
pip install pywin32
```

Запуск для первого `.docx` в папке:

```powershell
python review_word.py
```

Запуск для конкретного файла:

```powershell
python review_word.py "MyDocument.docx"
```

## Примечания

- На Windows должен быть установлен Microsoft Word.
- Скрипт работает через COM-автоматизацию Word, поэтому скорость зависит не только от Python, но и от самого Word.
- Часть задуманных правил сейчас намеренно отключена, потому что на реальных документах они работали нестабильно или слишком медленно.
