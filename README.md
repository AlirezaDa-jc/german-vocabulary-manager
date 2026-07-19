# German Vocabulary Manager

A Windows desktop app + Excel workbook for building a German vocabulary
database. You type German words into the workbook, click one button,
and the app enriches them with grammar, translations, examples, and
pronunciation audio pulled from free online sources.

## Project overview

German Vocabulary Manager is an Excel-based vocabulary system with a
small Windows GUI (`app.py`, packaged as `GermanVocabularyManager.exe`)
on top of it. The workbook itself is stored and edited locally, but the
**autofill** feature needs an internet connection to look up each word.

## Features

- One workbook (`vocabulary.xlsx`) holding your vocabulary, grammar
  reference, verb conjugations, adjective comparisons, statistics, and
  settings.
- Autofill: type a German word, get article, plural, declension, verb
  conjugation, adjective comparison, English/Persian translation,
  IPA, example sentence, synonyms/antonyms, and pronunciation audio —
  automatically.
- A simple GUI with four buttons: create/reset the workbook, run
  autofill, open the workbook, open the project folder.
- No paid services and no API key — only free, public data sources.

## What is online vs. local

This app is **not** a fully offline tool. Be precise about this:

| Action | Requires internet? |
|---|---|
| Creating / resetting the workbook | No |
| Opening and editing the workbook by hand | No |
| Autofill (looking up words) | **Yes** |

Workbook storage is local-first — your data lives in a plain `.xlsx`
file on your computer. Enrichment is online-first — autofill calls
public web APIs (Wiktionary, Wikimedia Commons, Tatoeba,
OpenThesaurus, and a translation fallback) to fetch data it doesn't
already have.

## Windows quick start (end users)

1. Extract `GermanVocabularyManager.zip`.
2. Run `GermanVocabularyManager.exe`.
3. Click **Create / Reset Workbook** if `vocabulary.xlsx` doesn't exist yet.
4. Click **Open Workbook**.
5. Go to the **Word** sheet and type in the German words you want to learn.
6. Save and close Excel.
7. Return to the app and click **Autofill Vocabulary**.
8. Wait for the log to show it's finished.
9. Click **Open Workbook** again to review the results on the
   **Vocabulary**, **Verbs**, and **Adjectives** sheets.

## Using the GUI buttons

**Create / Reset Workbook**
Runs `create_excel.py` to create `vocabulary.xlsx`, or rebuild it from
scratch if it already exists (this resets the workbook — back up first
if you want to keep existing data).

**Autofill Vocabulary**
Runs `autofill.py`. Looks up every pending word from the **Word**
sheet, fills in the **Vocabulary** sheet (and **Verbs**/**Adjectives**
when relevant), downloads pronunciation audio, updates statistics, and
writes progress to the log.

**Open Workbook**
Opens `vocabulary.xlsx` in Excel.

**Open Project Folder**
Opens the folder containing the workbook, `data/`, `audio/`,
`images/`, and the executable.

### Important: close Excel before running Create/Reset or Autofill

Excel locks the workbook file while it's open. The app checks for that
lock to avoid write errors or corrupting the file — if `vocabulary.xlsx`
is open in Excel, close it first, then click the button.

## Excel workbook workflow

1. Type words into the **Word** sheet — this is the only sheet you type
   new words into.
2. Save and close the workbook.
3. Run autofill from the app.
4. The app reads pending words from **Word**, marks them processed, and
   writes the enriched data into **Vocabulary** (plus **Verbs** /
   **Adjectives** for those word types).
5. Fields you maintain by hand — **Tags**, **Favorite**, **Learned**,
   **Review Date** — are never touched by autofill.

## Workbook sheets

| Sheet | Purpose |
|---|---|
| **Word** | Your input sheet — type new German words here. |
| **Vocabulary** | Main database: German, Article, Plural, Word Type, English, Persian, IPA, Pronunciation URL, Example Sentence, Example Translation, Synonyms, Antonyms, Gender, Genitive, Dative, Accusative, Notes, Tags, Favorite, Learned, Review Date. |
| **Verbs** | Infinitive, ich, du, er/sie/es, wir, ihr, sie/Sie, Präteritum, Perfekt, Partizip II, Hilfsverb, Trennbar, Reflexiv, Irregular, Meaning, Example. |
| **Adjectives** | Positive, Comparative, Superlative, Meaning, Example. |
| **Grammar** | Static German grammar reference with rules and examples. |
| **Statistics** | Formula-driven counts: total words, nouns, verbs, adjectives, favorites, learned, progress %. |
| **Settings** | App metadata and last-updated timestamp. |

## Data sources

All sources are free and require no API key:

- **German Wiktionary** — primary source: word type, article, plural,
  declension, conjugation data, adjective comparison, IPA, examples,
  translations, synonyms, antonyms.
- **Wikimedia Commons** — resolves and downloads pronunciation audio
  referenced by Wiktionary.
- **Tatoeba** — fallback example sentences.
- **OpenThesaurus** — fallback German synonyms.
- **deep-translator / GoogleTranslator** — last-resort machine
  translation for English/Persian and example-sentence translation,
  used only when the sources above don't provide enough. This is a
  best-effort, unofficial fallback — review its output.

## Installation from source (developers)

```bash
git clone <your-repo-url>
cd GermanVocabularyManager
pip install -r requirements.txt

python create_excel.py   # creates vocabulary.xlsx
python autofill.py       # fills pending words from the Word sheet
python app.py            # runs the GUI directly with Python
```

## Building the Windows executable

```bash
python -m PyInstaller GermanVocabularyAutofill.spec --clean --noconfirm
```

Output:

```
dist/GermanVocabularyManager/GermanVocabularyManager.exe
dist/GermanVocabularyManager.zip
```

Distribute the whole folder or the zip — not the `.exe` alone.
`vocabulary.xlsx`, `data/`, `audio/`, and `images/` are expected to sit
next to the executable (they're created there automatically if
missing).

## Project structure

```
GermanVocabularyManager/
├── app.py                          # Windows GUI entry point
├── autofill.py                      # autofill orchestration
├── create_excel.py                   # workbook generator
├── config.py                          # paths, endpoints, styling, PyInstaller-aware paths
├── GermanVocabularyAutofill.spec       # PyInstaller build spec
├── requirements.txt
├── README.md
├── vocabulary.xlsx
├── core/
│   ├── audio_downloader.py            # Commons audio download
│   ├── dictionary.py                   # Wiktionary fetch + parse
│   ├── excel_manager.py                 # workbook read/write
│   ├── http_client.py                    # retrying HTTP session
│   ├── models.py                          # data classes
│   ├── parsers.py                          # noun/verb/adjective parsing
│   ├── statistics_manager.py                # Statistics sheet
│   └── translators.py                        # Tatoeba/OpenThesaurus/MT fallback
├── data/                                # cache + logs
├── audio/                                # downloaded pronunciation files
├── images/                                # reserved for future use
├── dist/                                   # PyInstaller output (generated, not committed)
```

## Logs and generated files

Autofill logs are written to `data/`. Check them first if a run seems
to have failed or skipped words.

## Known limitations

- Autofill needs internet access; availability, rate limits, or source
  structure changes can affect results.
- Wiktionary entries vary in completeness — some words will come back
  with blank fields.
- Machine translation fallback is best-effort — review it.
- Verb conjugation and adjective comparison may need manual
  verification, especially for irregular words.
- Pronunciation audio depends on Wiktionary/Commons having a usable
  file for that word.
- Excel must be closed before Create/Reset or Autofill runs.

## Troubleshooting

**"Workbook is open" error** — Close Excel completely, then retry.

**No words get processed** — Add words to the **Word** sheet and save
the workbook before running Autofill.

**Missing translations/examples** — The online source didn't have
enough data for that word, or the network request failed; try again
later.

**The .exe won't open / Windows blocks it** — The app isn't
code-signed. Extract the zip fully first (don't run from inside the
zip), then allow it through Windows security if you trust the source.

**Missing audio** — Not every word has a pronunciation clip available
on Wikimedia Commons.

**Build fails** — Make sure dependencies from `requirements.txt` are
installed in the Python environment you're building from, then rerun
the PyInstaller command.

## License

See [LICENSE](LICENSE) for the license covering this project's code.
Data pulled by autofill comes from third parties with their own
licenses (e.g. Wiktionary content is CC BY-SA, Tatoeba sentences are
CC BY 2.0 FR) — that data is not covered by this project's license.
