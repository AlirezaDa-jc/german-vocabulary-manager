"""
Fallback data sources used only when Wiktionary itself did not supply a
value:

- :class:`TatoebaClient`        — example sentence + English translation
- :class:`OpenThesaurusClient`  — German synonyms
- :class:`MachineTranslator`    — best-effort English/Persian translation
  via the free, keyless ``deep-translator`` package (unofficial Google
  Translate web endpoint — no API key, no paid tier, no guaranteed
  uptime). This is a genuine best-effort fallback, not an authoritative
  source; results should be spot-checked.

All network calls go through ``core.http_client.client`` so retries,
timeouts and rate limiting are consistent across the whole project.
"""

from __future__ import annotations

import logging
from typing import List, Optional, Tuple

import config
from core.http_client import client

logger = logging.getLogger(__name__)


class TatoebaClient:
    """Free community sentence database — used as an example-sentence
    fallback when Wiktionary has none for the looked-up word."""

    def find_example(self, word: str) -> Optional[Tuple[str, str]]:
        """Return ``(german_sentence, english_translation)`` or ``None``."""
        examples = self.find_examples(word, limit=1)
        return examples[0] if examples else None

    def find_examples(self, word: str, limit: int = 4) -> List[Tuple[str, str]]:
        """Return up to ``limit`` German example sentences with English translations."""
        params = {
            "from": "deu",
            "to": "eng",
            "query": word,
            "orphans": "no",
            "unapproved": "no",
        }
        response = client.get(config.TATOEBA_API, params=params)
        if response is None:
            return []
        try:
            data = response.json()
            results = data.get("results", [])
        except ValueError:
            logger.warning("Tatoeba returned non-JSON response for %r", word)
            return []

        examples: List[Tuple[str, str]] = []
        for result in results:
            german_text = result.get("text")
            if not german_text:
                continue

            translations = result.get("translations", [])
            english_text = ""
            for group in translations:
                for translation in group:
                    if translation.get("lang") == "eng":
                        english_text = translation.get("text") or ""
                        break
                if english_text:
                    break

            examples.append((german_text, english_text))
            if len(examples) >= limit:
                break

        return examples


class OpenThesaurusClient:
    """Free German synonym dictionary — synonym fallback source."""

    def find_synonyms(self, word: str) -> List[str]:
        params = {"q": word, "format": "application/json"}
        response = client.get(config.OPENTHESAURUS_API, params=params)
        if response is None:
            return []
        try:
            data = response.json()
        except ValueError:
            logger.warning("OpenThesaurus returned non-JSON response for %r", word)
            return []
        synonyms: List[str] = []
        for synset in data.get("synsets", []):
            for term in synset.get("terms", []):
                value = term.get("term")
                if value and value.lower() != word.lower():
                    synonyms.append(value)
        # de-duplicate while preserving order, then keep at most 3
        return list(dict.fromkeys(synonyms))[:3]


class MachineTranslator:
    """Best-effort English/Persian translation fallback.

    Uses ``deep-translator``'s ``GoogleTranslator`` wrapper, which talks
    to an unofficial, keyless public endpoint. This is intentionally the
    LAST fallback in the pipeline (after Wiktionary's own translation
    table) because it is unofficial and can break or change without
    notice.
    """

    def __init__(self) -> None:
        self._available = False
        if not config.ENABLE_MACHINE_TRANSLATION_FALLBACK:
            return
        try:
            from deep_translator import GoogleTranslator  # type: ignore

            self._GoogleTranslator = GoogleTranslator
            self._available = True
        except ImportError:
            logger.warning(
                "deep-translator is not installed — English/Persian "
                "machine-translation fallback disabled. "
                "Install with: pip install deep-translator"
            )

    @property
    def available(self) -> bool:
        return self._available

    def translate(self, text: str, target: str) -> Optional[str]:
        """Translate ``text`` (German) to ``target`` ("en" or "fa")."""
        if not self._available or not text:
            return None
        try:
            translator = self._GoogleTranslator(source="de", target=target)
            result = translator.translate(text)
            return result.strip() if result else None
        except Exception as exc:  # pragma: no cover - defensive, network dependent
            logger.warning("Machine translation (%s) failed for %r: %s", target, text, exc)
            return None
