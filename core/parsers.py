"""
Part-of-speech-specific parsers. Each takes the generic
:class:`core.models.RawEntry` produced by ``GermanDictionary`` and turns
it into a structured ``*Info`` dataclass ready to be written into the
workbook.

Where Wiktionary's overview template is missing a form (this happens for
many entries, especially verbs), each parser falls back to regular
German grammar rules and records that the value was *derived, not
confirmed* in the ``notes`` field so the user knows to double-check it.
"""

from __future__ import annotations

import logging
from typing import Dict, Optional

import config
from core.dictionary import strip_wikitext
from core.models import AdjectiveInfo, NounInfo, RawEntry, VerbInfo

logger = logging.getLogger(__name__)

_GENUS_TO_ARTICLE = {"m": "der", "f": "die", "n": "das"}


class NounParser:
    """Builds a :class:`NounInfo` from a noun's ``RawEntry``."""

    def parse(self, entry: RawEntry) -> NounInfo:
        info = NounInfo(word=entry.word)
        params: Dict[str, str] = entry.templates.get(
            "Deutsch Substantiv Übersicht", {}
        )

        genus = params.get("Genus") or params.get("Genus 1")
        if genus:
            genus = genus.strip().lower()
            info.gender = {"m": "masculine", "f": "feminine", "n": "neuter"}.get(
                genus
            )
            info.article = _GENUS_TO_ARTICLE.get(genus)

        info.plural = self._first_present(
            params, ["Nominativ Plural", "Nominativ Plural 1"]
        )
        info.genitive = self._first_present(
            params, ["Genitiv Singular", "Genitiv Singular 1"]
        )
        info.dative = self._first_present(
            params, ["Dativ Singular", "Dativ Singular 1"]
        )
        info.accusative = self._first_present(
            params, ["Akkusativ Singular", "Akkusativ Singular 1"]
        )

        if not info.article:
            info.notes.append(
                "Article/gender not found in Wiktionary overview table — "
                "verify manually."
            )
        if not info.plural:
            info.notes.append("Plural form not found — verify manually.")

        info.ipa = entry.ipa
        info.audio_filename = entry.audio_filename
        if entry.examples:
            info.example = entry.examples[0]
        if entry.english_translations:
            info.english = ", ".join(dict.fromkeys(entry.english_translations))
        info.synonyms = entry.synonyms
        info.antonyms = entry.antonyms
        return info

    @staticmethod
    def _first_present(params: Dict[str, str], keys) -> Optional[str]:
        for key in keys:
            value = params.get(key)
            if value:
                return strip_wikitext(value)
        return None


class VerbParser:
    """Builds a :class:`VerbInfo` from a verb's ``RawEntry``.

    Present-tense persons not covered by the Wiktionary overview template
    (``wir``, ``ihr``, ``sie/Sie``) are derived with regular conjugation
    rules and flagged as derived, since German Wiktionary's verb overview
    template typically only lists ich/du/er/Präteritum-ich/Partizip II
    explicitly.
    """

    def parse(self, entry: RawEntry) -> VerbInfo:
        infinitive = entry.word
        info = VerbInfo(infinitive=infinitive)
        params: Dict[str, str] = entry.templates.get("Deutsch Verb Übersicht", {})

        info.ich = self._clean(params.get("Präsens_ich"))
        info.du = self._clean(params.get("Präsens_du"))
        info.er_sie_es = self._clean(
            params.get("Präsens_er, sie, es") or params.get("Präsens_er")
        )
        info.praeteritum = self._clean(
            params.get("Präteritum_ich") or params.get("Präteritum")
        )
        info.partizip_ii = self._clean(params.get("Partizip II"))
        info.hilfsverb = self._clean(params.get("Hilfsverb"))

        derived_any = False
        stem_present = self._present_stem(infinitive)

        if not info.ich:
            info.ich = f"{stem_present}e"
            derived_any = True
        if not info.du:
            info.du = f"{stem_present}st"
            derived_any = True
        if not info.er_sie_es:
            info.er_sie_es = f"{stem_present}t"
            derived_any = True

        # wir / sie(Sie) are regularly identical to the infinitive; ihr is
        # the infinitive stem + "t". These regular derivations are usually
        # correct even for irregular verbs (only ich/du/er change for
        # strong verbs), but we still flag them since exceptions exist
        # (e.g. modal verbs, "sein").
        info.wir = infinitive
        info.sie_Sie = infinitive
        info.ihr = f"{stem_present}t"
        derived_any = True  # wir/ihr/sie(Sie) are always heuristically derived

        if not info.praeteritum:
            info.praeteritum = f"{stem_present}te"
            derived_any = True
        if not info.perfekt and info.partizip_ii:
            aux = info.hilfsverb or "haben"
            info.perfekt = f"{aux} ... {info.partizip_ii}"
        if not info.partizip_ii:
            info.partizip_ii = f"ge{stem_present}t"
            derived_any = True
        if not info.hilfsverb:
            info.hilfsverb = "haben"
            derived_any = True

        info.irregular = not info.praeteritum.endswith("te")

        info.reflexiv = infinitive.strip().lower().startswith("sich ")

        info.trennbar = self._is_separable(infinitive)

        if derived_any:
            info.notes.append(
                "Some conjugation forms were derived using regular verb "
                "rules (Wiktionary overview table did not list them). "
                "Please verify, especially if this is an irregular "
                "(strong/mixed) verb."
            )

        info.ipa = entry.ipa
        info.audio_filename = entry.audio_filename
        if entry.examples:
            info.example = entry.examples[0]
        if entry.english_translations:
            info.meaning = ", ".join(dict.fromkeys(entry.english_translations))
        return info

    @staticmethod
    def _clean(value: Optional[str]) -> Optional[str]:
        return strip_wikitext(value) if value else None

    @staticmethod
    def _present_stem(infinitive: str) -> str:
        base = infinitive.strip()
        if base.lower().startswith("sich "):
            base = base[5:]
        if base.endswith("en"):
            return base[:-2]
        if base.endswith("n"):
            return base[:-1]
        return base

    @staticmethod
    def _is_separable(infinitive: str) -> Optional[bool]:
        word = infinitive.strip().lower()
        if word.startswith("sich "):
            word = word[5:]
        for prefix in config.INSEPARABLE_PREFIXES:
            if word.startswith(prefix):
                return False
        for prefix in sorted(config.SEPARABLE_PREFIXES, key=len, reverse=True):
            if word.startswith(prefix) and len(word) > len(prefix) + 2:
                return True
        return None  # inconclusive — leave blank rather than guess wrong


class AdjectiveParser:
    """Builds an :class:`AdjectiveInfo` from an adjective's ``RawEntry``."""

    def parse(self, entry: RawEntry) -> AdjectiveInfo:
        positive = entry.word
        info = AdjectiveInfo(positive=positive)
        params: Dict[str, str] = entry.templates.get(
            "Deutsch Adjektiv Übersicht", {}
        )

        comparative = params.get("Komparativ")
        superlative = params.get("Superlativ")
        derived = False

        if comparative:
            info.comparative = strip_wikitext(comparative)
        else:
            info.comparative = self._regular_comparative(positive)
            derived = True

        if superlative:
            info.superlative = strip_wikitext(superlative)
        else:
            info.superlative = self._regular_superlative(positive)
            derived = True

        if derived:
            info.notes.append(
                "Comparative/superlative derived using regular adjective "
                "rules — verify manually if this adjective is irregular "
                "(e.g. gut/besser/best-, viel/mehr/meist-)."
            )

        info.ipa = entry.ipa
        info.audio_filename = entry.audio_filename
        if entry.examples:
            info.example = entry.examples[0]
        if entry.english_translations:
            info.meaning = ", ".join(dict.fromkeys(entry.english_translations))
        return info

    @staticmethod
    def _regular_comparative(word: str) -> str:
        return f"{word}er"

    @staticmethod
    def _regular_superlative(word: str) -> str:
        return f"am {word}sten"
