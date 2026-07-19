#!/usr/bin/env python3
"""
autofill.py
============

Type a single German word into the "German" column of the Vocabulary
sheet, leave every other cell in that row empty, then run:

    python autofill.py

Every empty column in that row is filled automatically from free
sources (German Wiktionary, Wikimedia Commons for audio, Tatoeba,
OpenThesaurus, and an optional keyless machine-translation fallback for
English/Persian). Nouns additionally populate declension into the
Vocabulary row; verbs and adjectives additionally get a full row added
to the Verbs / Adjectives sheet.

Rows that are already filled in (English column non-empty) are skipped,
so re-running the script is fast and safe — only new words are looked
up.
"""

from __future__ import annotations

import logging
import sys

import config
from core.audio_downloader import AudioDownloader
from core.dictionary import POS_MAP, GermanDictionary
from core.excel_manager import ExcelManager
from core.models import RawEntry
from core.parsers import AdjectiveParser, NounParser, VerbParser
from core.statistics_manager import StatisticsManager
from core.translators import MachineTranslator, OpenThesaurusClient, TatoebaClient

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
    handlers=[
        logging.FileHandler(config.LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("autofill")


class VocabularyAutofiller:
    """Coordinates every class to fill one Vocabulary row end-to-end."""

    def __init__(self) -> None:
        self.excel = ExcelManager()
        self.dictionary = GermanDictionary()
        self.noun_parser = NounParser()
        self.verb_parser = VerbParser()
        self.adjective_parser = AdjectiveParser()
        self.audio = AudioDownloader()
        self.tatoeba = TatoebaClient()
        self.thesaurus = OpenThesaurusClient()
        self.translator = MachineTranslator()

    # ------------------------------------------------------------------
    # Orchestration
    # ------------------------------------------------------------------

    def run(self) -> None:
        self.excel.load()
        pending = list(self.excel.iter_pending_word_rows())

        if not pending:
            logger.info(
                "No pending rows found. Type a German word into the Word "
                "sheet's 'German' column and run this script again."
            )
            return

        logger.info("Found %d word(s) to autofill.", len(pending))
        for word_row_idx, word in pending:
            vocab_row_idx = self.excel.append_vocab_seed_row(word)
            logger.info("Processing Word row %d -> Vocabulary row %d: %r", word_row_idx, vocab_row_idx, word)
            try:
                detected_type = self._process_word(vocab_row_idx, word)
                self.excel.update_word_row(
                    word_row_idx,
                    {
                        "Processed": "Yes",
                        "Detected Type": detected_type,
                    },
                )
            except Exception:
                logger.exception("Failed to process %r (Word row %d) — skipped.", word, word_row_idx)
                self.excel.update_word_row(
                    word_row_idx,
                    {
                        "Processed": "No",
                        "Notes": "Autofill failed. Check data/autofill.log for details.",
                    },
                )

        StatisticsManager().touch_last_updated(self.excel.workbook)
        self.excel.save()
        logger.info("Done. %d row(s) processed.", len(pending))

    # ------------------------------------------------------------------
    # Per-word processing
    # ------------------------------------------------------------------

    def _process_word(self, row_idx: int, word: str) -> str:
        entry = self.dictionary.lookup(word)

        if not entry.found:
            logger.warning("%r not found on German Wiktionary.", word)
            self.excel.update_vocab_row(
                row_idx,
                {
                    "Notes": (
                        "No Wiktionary entry found for this word — please "
                        "fill in manually or check spelling/capitalization."
                    )
                },
            )
            return "Not found"

        word_type = POS_MAP.get(entry.pos or "", "Other")
        print(f"Detected type: {word_type}")
        if word_type == "Noun":
            self._fill_noun(row_idx, entry)
        elif word_type == "Verb":
            self._fill_verb(row_idx, entry)
        elif word_type == "Adjective":
            self._fill_adjective(row_idx, entry)
        else:
            self._fill_generic(row_idx, entry, word_type)

        return word_type

    # ------------------------------------------------------------------
    # Noun
    # ------------------------------------------------------------------

    def _fill_noun(self, row_idx: int, entry: RawEntry) -> None:
        info = self.noun_parser.parse(entry)
        self._enrich_common(entry, info)

        pronunciation_url = self._download_audio(entry, info.word)
        english = info.english or self._machine_translate(info.word, "en")
        persian = self._persian_translation(entry, english or info.word)
        synonyms = info.synonyms[:3] if info.synonyms else []

        values = {
            "Article": info.article,
            "Plural": info.plural,
            "Word Type": "Noun",
            "English": english,
            "Persian": persian,
            "IPA": info.ipa,
            "Pronunciation URL": pronunciation_url,
            "Example Sentence": info.example,
            "Example Translation": self._translate_example(info.example),
            "Synonyms": ", ".join(synonyms) if synonyms else None,
            "Antonyms": ", ".join(info.antonyms) if info.antonyms else None,
            "Gender": info.gender,
            "Genitive": info.genitive,
            "Dative": info.dative,
            "Accusative": info.accusative,
            "Notes": "; ".join(info.notes) if info.notes else None,
        }
        self.excel.update_vocab_row(row_idx, values)

    # ------------------------------------------------------------------
    # Verb
    # ------------------------------------------------------------------

    def _fill_verb(self, row_idx: int, entry: RawEntry) -> None:
        info = self.verb_parser.parse(entry)

        pronunciation_url = self._download_audio(entry, info.infinitive)
        english = info.meaning or self._machine_translate(info.infinitive, "en")
        persian = self._persian_translation(entry, english or info.infinitive)
        example = info.example or self._fallback_example(info.infinitive)
        synonyms = (entry.synonyms or self.thesaurus.find_synonyms(entry.word))[:3]

        vocab_values = {
            "Word Type": "Verb",
            "English": english,
            "Persian": persian,
            "IPA": info.ipa,
            "Pronunciation URL": pronunciation_url,
            "Example Sentence": example,
            "Example Translation": self._translate_example(example),
            "Synonyms": ", ".join(synonyms) if synonyms else None,
            "Antonyms": ", ".join(entry.antonyms) if entry.antonyms else None,
            "Notes": "; ".join(info.notes) if info.notes else None,
        }
        self.excel.update_vocab_row(row_idx, vocab_values)

        verb_values = {
            "ich": info.ich,
            "du": info.du,
            "er/sie/es": info.er_sie_es,
            "wir": info.wir,
            "ihr": info.ihr,
            "sie/Sie": info.sie_Sie,
            "Präteritum": info.praeteritum,
            "Perfekt": info.perfekt,
            "Partizip II": info.partizip_ii,
            "Hilfsverb": info.hilfsverb,
            "Trennbar": self._bool_to_str(info.trennbar),
            "Reflexiv": self._bool_to_str(info.reflexiv),
            "Irregular": self._bool_to_str(info.irregular),
            "Meaning": english,
            "Example": example,
        }
        self.excel.upsert_row(
            config.SHEET_VERBS, "Infinitive", info.infinitive, verb_values
        )

    # ------------------------------------------------------------------
    # Adjective
    # ------------------------------------------------------------------

    def _fill_adjective(self, row_idx: int, entry: RawEntry) -> None:
        info = self.adjective_parser.parse(entry)

        pronunciation_url = self._download_audio(entry, info.positive)
        english = info.meaning or self._machine_translate(info.positive, "en")
        persian = self._persian_translation(entry, english or info.positive)
        example = info.example or self._fallback_example(info.positive)
        synonyms = (entry.synonyms or self.thesaurus.find_synonyms(entry.word))[:3]

        vocab_values = {
            "Word Type": "Adjective",
            "English": english,
            "Persian": persian,
            "IPA": info.ipa,
            "Pronunciation URL": pronunciation_url,
            "Example Sentence": example,
            "Example Translation": self._translate_example(example),
            "Synonyms": ", ".join(synonyms) if synonyms else None,
            "Antonyms": ", ".join(entry.antonyms) if entry.antonyms else None,
            "Notes": "; ".join(info.notes) if info.notes else None,
        }
        self.excel.update_vocab_row(row_idx, vocab_values)

        adjective_values = {
            "Comparative": info.comparative,
            "Superlative": info.superlative,
            "Meaning": english,
            "Example": example,
        }
        self.excel.upsert_row(
            config.SHEET_ADJECTIVES, "Positive", info.positive, adjective_values
        )

    # ------------------------------------------------------------------
    # Generic (adverb, preposition, pronoun, unknown, ...)
    # ------------------------------------------------------------------

    def _fill_generic(self, row_idx: int, entry: RawEntry, word_type: str) -> None:
        english = (
            ", ".join(dict.fromkeys(entry.english_translations))
            if entry.english_translations
            else self._machine_translate(entry.word, "en")
        )
        persian = self._persian_translation(entry, english or entry.word)
        example = entry.examples[0] if entry.examples else self._fallback_example(entry.word)
        synonyms = (entry.synonyms or self.thesaurus.find_synonyms(entry.word))[:3]

        values = {
            "Word Type": word_type,
            "English": english,
            "Persian": persian,
            "IPA": entry.ipa,
            "Pronunciation URL": self._download_audio(entry, entry.word),
            "Example Sentence": example,
            "Example Translation": self._translate_example(example),
            "Synonyms": ", ".join(synonyms) if synonyms else None,
            "Antonyms": ", ".join(entry.antonyms) if entry.antonyms else None,
        }
        self.excel.update_vocab_row(row_idx, values)

    # ------------------------------------------------------------------
    # Shared enrichment helpers
    # ------------------------------------------------------------------

    def _enrich_common(self, entry: RawEntry, info) -> None:
        """Fill in Tatoeba/OpenThesaurus fallbacks directly onto the info object."""
        if not info.example:
            fallback = self.tatoeba.find_example(getattr(info, "word", entry.word))
            if fallback:
                info.example = fallback[0]
        if hasattr(info, "synonyms") and not info.synonyms:
            info.synonyms = self.thesaurus.find_synonyms(getattr(info, "word", entry.word))[:3]
        elif hasattr(info, "synonyms"):
            info.synonyms = info.synonyms[:3]

    def _download_audio(self, entry: RawEntry, word: str):
        if not entry.audio_filename:
            return None
        return self.audio.download(entry.audio_filename, word)

    def _machine_translate(self, word: str, target: str):
        return self.translator.translate(word, target)

    def _persian_translation(self, entry: RawEntry, fallback_source: str):
        if entry.persian_translations:
            return ", ".join(dict.fromkeys(entry.persian_translations))
        return self._machine_translate(entry.word, "fa")

    def _translate_example(self, example):
        if not example:
            return None
        return self._machine_translate(example, "en")

    def _fallback_example(self, word: str):
        result = self.tatoeba.find_example(word)
        return result[0] if result else None

    @staticmethod
    def _bool_to_str(value):
        if value is None:
            return None
        return "TRUE" if value else "FALSE"


def main() -> None:
    autofiller = VocabularyAutofiller()
    autofiller.run()


if __name__ == "__main__":
    main()
