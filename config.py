"""
config.py
=========

Central configuration for the German Vocabulary Manager.

Holds file paths, network settings, API endpoints (all free / public),
sheet and column layouts, and styling constants shared by
``create_excel.py`` and ``autofill.py``.

Nothing in this module performs I/O; it only defines constants so that
every other module has a single source of truth.
"""

from __future__ import annotations

from pathlib import Path

# ---------------------------------------------------------------------------
# Filesystem layout
# ---------------------------------------------------------------------------

BASE_DIR: Path = Path(__file__).resolve().parent
DATA_DIR: Path = BASE_DIR / "data"
AUDIO_DIR: Path = BASE_DIR / "audio"
IMAGES_DIR: Path = BASE_DIR / "images"
WORKBOOK_PATH: Path = BASE_DIR / "vocabulary.xlsx"
CACHE_FILE: Path = DATA_DIR / "lookup_cache.json"
LOG_FILE: Path = DATA_DIR / "autofill.log"

for _directory in (DATA_DIR, AUDIO_DIR, IMAGES_DIR):
    _directory.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Networking
# ---------------------------------------------------------------------------

USER_AGENT: str = (
    "GermanVocabularyManager/1.0 "
    "(offline personal study tool; contact: local-user@example.invalid)"
)
REQUEST_TIMEOUT: int = 15          # seconds per HTTP request
MAX_RETRIES: int = 3               # retry attempts for a failing request
RETRY_BACKOFF_SECONDS: float = 1.5  # exponential backoff base
RATE_LIMIT_DELAY_SECONDS: float = 0.5  # politeness delay between lookups

# ---------------------------------------------------------------------------
# Free data source endpoints
# ---------------------------------------------------------------------------

# German Wiktionary — MediaWiki action API, returns raw wikitext we parse
# ourselves. This is the primary source for grammar (gender, plural,
# declension, conjugation, IPA, examples, synonyms/antonyms, translations).
WIKTIONARY_DE_API: str = "https://de.wiktionary.org/w/api.php"

# Wikimedia Commons — used to resolve the real file URL for pronunciation
# audio files referenced inside Wiktionary entries (e.g. "De-Haus.ogg").
COMMONS_API: str = "https://commons.wikimedia.org/w/api.php"

# Tatoeba — free, community example-sentence database with translations.
# Used as a fallback / supplement when Wiktionary has no example sentence.
TATOEBA_API: str = "https://tatoeba.org/eng/api_v0/search"

# OpenThesaurus — free German synonym dictionary, used as a fallback when
# Wiktionary has no "Sinnverwandte Woerter" section.
OPENTHESAURUS_API: str = "https://www.openthesaurus.de/synonyme/search"

# Free, keyless machine-translation endpoint (via the `deep-translator`
# package's GoogleTranslator wrapper) used ONLY as a last-resort fallback
# for the English / Persian columns when Wiktionary supplies no
# translation at all. This talks to an unofficial public endpoint with no
# API key and no SLA — treat its output as "best effort", not authoritative.
ENABLE_MACHINE_TRANSLATION_FALLBACK: bool = True

# ---------------------------------------------------------------------------
# Workbook layout — sheet names
# ---------------------------------------------------------------------------

SHEET_WORD = "Word"
SHEET_VOCAB = "Vocabulary"
SHEET_VERBS = "Verbs"
SHEET_ADJECTIVES = "Adjectives"
SHEET_GRAMMAR = "Grammar"
SHEET_STATS = "Statistics"
SHEET_SETTINGS = "Settings"

ALL_SHEETS = (
    SHEET_WORD,
    SHEET_VOCAB,
    SHEET_VERBS,
    SHEET_ADJECTIVES,
    SHEET_GRAMMAR,
    SHEET_STATS,
    SHEET_SETTINGS,
)

# ---------------------------------------------------------------------------
# Word input sheet columns
# ---------------------------------------------------------------------------

WORD_COLUMNS = [
    "German",
    "Processed",
    "Detected Type",
    "Notes",
]

# ---------------------------------------------------------------------------
# Vocabulary sheet columns (ordered — index 0 == column A)
# ---------------------------------------------------------------------------

VOCAB_COLUMNS = [
    "German",
    "Article",
    "Plural",
    "Word Type",
    "English",
    "Persian",
    "IPA",
    "Pronunciation URL",
    "Example Sentence",
    "Example Translation",
    "Synonyms",
    "Antonyms",
    "Level (A1-C2)",
    "Gender",
    "Genitive",
    "Dative",
    "Accusative",
    "Notes",
    "Tags",
    "Favorite",
    "Learned",
    "Review Date",
]
LEARNED_UNCHECKED = "☐ Unchecked"
LEARNED_CHECKED = "☑ Checked"

# Columns that autofill.py is allowed to populate automatically. "German"
# is the user-entered seed column and is never overwritten.
# "Learned" is manually selected as ☐ Unchecked / ☑ Checked in the workbook.
VOCAB_AUTOFILL_COLUMNS = [c for c in VOCAB_COLUMNS if c != "German"]

# ---------------------------------------------------------------------------
# Verb sheet columns
# ---------------------------------------------------------------------------

VERB_COLUMNS = [
    "Infinitive",
    "ich",
    "du",
    "er/sie/es",
    "wir",
    "ihr",
    "sie/Sie",
    "Präteritum",
    "Perfekt",
    "Partizip II",
    "Hilfsverb",
    "Trennbar",
    "Reflexiv",
    "Irregular",
    "Meaning",
    "Example",
]

# ---------------------------------------------------------------------------
# Adjective sheet columns
# ---------------------------------------------------------------------------

ADJECTIVE_COLUMNS = [
    "Positive",
    "Comparative",
    "Superlative",
    "Meaning",
    "Example",
]

# ---------------------------------------------------------------------------
# Grammar sheet — static reference content (headers only; content is
# written once by create_excel.py)
# ---------------------------------------------------------------------------

GRAMMAR_COLUMNS = ["Topic", "Rule", "Example"]

# ---------------------------------------------------------------------------
# Statistics sheet — label / formula pairs written by create_excel.py
# ---------------------------------------------------------------------------

STATS_ROWS = [
    "Total Words",
    "Nouns",
    "Verbs",
    "Adjectives",
    "Favorites",
    "Learned",
    "Progress %",
]

# ---------------------------------------------------------------------------
# Settings sheet — key / value pairs
# ---------------------------------------------------------------------------

SETTINGS_ROWS = [
    ("App Name", "German Vocabulary Manager"),
    ("Version", "1.0.0"),
    ("Data Sources", "Wiktionary (de), Wikimedia Commons, Tatoeba, OpenThesaurus"),
    ("Autofill Script", "autofill.py"),
    ("Audio Folder", "audio/"),
    ("Image Folder", "images/"),
    ("Last Updated", ""),
]

# ---------------------------------------------------------------------------
# Styling constants
# ---------------------------------------------------------------------------

HEADER_FILL_COLOR = "1F2937"      # dark slate
HEADER_FONT_COLOR = "FFFFFF"      # white
FONT_NAME = "Calibri"

DER_COLOR = "1E3A8A"   # blue   (der)
DIE_COLOR = "991B1B"   # red    (die)
DAS_COLOR = "166534"   # green  (das)

LEARNED_FILL = "C6EFCE"   # light green row fill
FAVORITE_FILL = "FFF2CC"  # light yellow row fill

VALID_ARTICLES = {"der", "die", "das"}
VALID_LEVELS = ["A1", "A2", "B1", "B2", "C1", "C2"]
VALID_WORD_TYPES = [
    "Noun",
    "Verb",
    "Adjective",
    "Adverb",
    "Preposition",
    "Conjunction",
    "Pronoun",
    "Numeral",
    "Interjection",
    "Other",
]

# Common German separable verb prefixes, used as a heuristic for the
# "Trennbar" (separable) column when Wiktionary does not state it
# explicitly.
SEPARABLE_PREFIXES = [
    "ab", "an", "auf", "aus", "bei", "da", "dar", "ein", "empor", "fort",
    "her", "hin", "los", "mit", "nach", "nieder", "statt", "teil", "vor",
    "weg", "weiter", "zu", "zurueck", "zurück", "zusammen", "fest", "fern",
    "gegenueber", "gegenüber", "heim", "hoch", "wieder", "zurecht",
]

# Prefixes that are NEVER separable (inseparable verb prefixes) — checked
# first so we don't misclassify e.g. "verkaufen" or "besuchen".
INSEPARABLE_PREFIXES = [
    "be", "emp", "ent", "er", "ge", "miss", "ver", "voll", "wider", "zer",
]
