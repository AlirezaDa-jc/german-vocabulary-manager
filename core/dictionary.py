"""
``GermanDictionary`` fetches the raw wikitext of a German Wiktionary
entry and parses it into a :class:`core.models.RawEntry`. All later,
part-of-speech-specific interpretation (noun / verb / adjective) is left
to the classes in ``core/parsers.py`` — this module only understands
generic Wiktionary markup.

No paid or key-gated API is used anywhere in this file.
"""

from __future__ import annotations

import logging
import re
from typing import Dict, List, Optional

import config
from core.http_client import client
from core.models import RawEntry

logger = logging.getLogger(__name__)

# Maps the German "Wortart" template value (as used inside
# {{Wortart|X|Deutsch}}) to the value we store in the "Word Type" column.
POS_MAP: Dict[str, str] = {
    "Substantiv": "Noun",
    "Verb": "Verb",
    "Adjektiv": "Adjective",
    "Adverb": "Adverb",
    "Präposition": "Preposition",
    "Konjunktion": "Conjunction",
    "Pronomen": "Pronoun",
    "Numerale": "Numeral",
    "Interjektion": "Interjection",
}

_WIKILINK_RE = re.compile(r"\[\[(?:[^\]|]*\|)?([^\]]+)\]\]")
_TEMPLATE_RE = re.compile(r"\{\{[^{}]*\}\}")
_BOLD_ITALIC_RE = re.compile(r"'{2,5}")
_REF_RE = re.compile(r"<ref[^>]*>.*?</ref>|<ref[^/]*/>", re.DOTALL)


def strip_wikitext(text: str) -> str:
    """Reduce a wikitext fragment to plain, human-readable text."""
    if not text:
        return ""
    text = _REF_RE.sub("", text)
    text = _WIKILINK_RE.sub(r"\1", text)
    # Repeatedly strip innermost {{...}} templates (qualifiers like {{K|...}})
    previous = None
    while previous != text:
        previous = text
        text = _TEMPLATE_RE.sub("", text)
    text = _BOLD_ITALIC_RE.sub("", text)
    text = text.replace("[", "").replace("]", "")
    return " ".join(text.split()).strip(" :;,.")


def _find_balanced_template(text: str, start_idx: int) -> Optional[str]:
    """Return the full ``{{...}}`` block starting at ``start_idx``.

    ``start_idx`` must point at the first ``{`` of ``{{``. Handles nested
    templates (e.g. ``{{Ü-Tabelle| ... {{Ü|en|house}} ... }}``) by tracking
    brace depth, since a naive non-greedy regex breaks on nested braces.
    """
    if text[start_idx : start_idx + 2] != "{{":
        return None
    depth = 0
    i = start_idx
    n = len(text)
    while i < n - 1:
        two = text[i : i + 2]
        if two == "{{":
            depth += 1
            i += 2
            continue
        if two == "}}":
            depth -= 1
            i += 2
            if depth == 0:
                return text[start_idx:i]
            continue
        i += 1
    return None  # unbalanced / truncated — give up gracefully


def _extract_template(text: str, template_name: str) -> Optional[str]:
    """Find the first ``{{template_name ...}}`` block (case-sensitive)."""
    marker = "{{" + template_name
    idx = text.find(marker)
    if idx == -1:
        return None
    return _find_balanced_template(text, idx)


def _parse_template_params(block: str) -> Dict[str, str]:
    """Parse ``{{Name|k1=v1|k2=v2|...}}`` into a ``{k: v}`` dict.

    Positional (unnamed) parameters are stored under their 1-based index
    as a string key ("1", "2", ...).
    """
    if not block:
        return {}
    inner = block.strip()
    if inner.startswith("{{"):
        inner = inner[2:]
    if inner.endswith("}}"):
        inner = inner[:-2]
    # Split on top-level "|" only (depth-aware, in case of nested templates
    # inside a parameter value).
    parts: List[str] = []
    depth = 0
    current = []
    i = 0
    while i < len(inner):
        ch = inner[i]
        if inner[i : i + 2] == "{{":
            depth += 1
            current.append(ch)
            i += 1
        elif inner[i : i + 2] == "}}":
            depth -= 1
            current.append(ch)
            i += 1
        elif ch == "|" and depth == 0:
            parts.append("".join(current))
            current = []
        else:
            current.append(ch)
        i += 1
    parts.append("".join(current))

    params: Dict[str, str] = {}
    positional_index = 0
    for part in parts[1:]:  # parts[0] is the template name
        if "=" in part:
            key, _, value = part.partition("=")
            params[key.strip()] = value.strip()
        else:
            positional_index += 1
            params[str(positional_index)] = part.strip()
    return params


def _extract_section(wikitext: str, level: int, heading_pattern: str) -> Optional[str]:
    """Return the body of the first section whose heading matches.

    ``level`` is the number of ``=`` characters (2 for ``==``, 3 for
    ``===``, 4 for ``====``). The body runs until the next heading of the
    SAME OR LOWER level (fewer or equal ``=`` characters means broader
    scope and therefore ends this section).
    """
    marker = "=" * level
    heading_re = re.compile(
        rf"^{re.escape(marker)}\s*{heading_pattern}\s*{re.escape(marker)}\s*$",
        re.MULTILINE,
    )
    match = heading_re.search(wikitext)
    if not match:
        return None
    start = match.end()
    end_re = re.compile(rf"^={{2,{level}}}[^=].*$", re.MULTILINE)
    end_match = end_re.search(wikitext, pos=start)
    end = end_match.start() if end_match else len(wikitext)
    return wikitext[start:end]


class GermanDictionary:
    """Fetches and parses a single German Wiktionary entry."""

    def __init__(self) -> None:
        self._cache: Dict[str, RawEntry] = {}

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def lookup(self, word: str) -> RawEntry:
        """Return a :class:`RawEntry` for ``word`` (cached per run)."""
        key = word.strip()
        if key in self._cache:
            return self._cache[key]

        entry = RawEntry(word=key)
        wikitext = self._fetch_wikitext(key)
        if wikitext is None:
            logger.warning("No Wiktionary page found for %r", key)
            self._cache[key] = entry
            return entry

        entry.wikitext = wikitext
        german_section = _extract_section(
            wikitext, level=2, heading_pattern=rf"{re.escape(key)}\s*\(\{{\{{Sprache\|Deutsch\}}\}}\)"
        )
        if german_section is None:
            logger.info("No German-language section found for %r", key)
            self._cache[key] = entry
            return entry

        entry.found = True
        self._parse_pos_and_templates(entry, german_section)
        self._parse_pronunciation(entry, german_section)
        self._parse_examples(entry, german_section)
        self._parse_synonyms_antonyms(entry, german_section)
        self._parse_translations(entry, german_section)

        self._cache[key] = entry
        return entry

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _fetch_wikitext(self, word: str) -> Optional[str]:
        params = {
            "action": "query",
            "prop": "revisions",
            "rvprop": "content",
            "rvslots": "main",
            "format": "json",
            "formatversion": "2",
            "titles": word,
        }
        response = client.get(config.WIKTIONARY_DE_API, params=params)
        if response is None:
            return None
        try:
            data = response.json()
            pages = data["query"]["pages"]
            if not pages:
                return None
            page = pages[0]
            if page.get("missing"):
                return None
            return page["revisions"][0]["slots"]["main"]["content"]
        except (KeyError, IndexError, ValueError) as exc:
            logger.error("Unexpected Wiktionary response for %r: %s", word, exc)
            return None

    def _parse_pos_and_templates(self, entry: RawEntry, section: str) -> None:
        """Find the first recognised POS subsection and its overview template."""
        heading_re = re.compile(
            r"^===\s*\{\{Wortart\|([^|}]+)\|Deutsch\}\}.*?===\s*$", re.MULTILINE
        )
        for match in heading_re.finditer(section):
            pos_raw = match.group(1).strip()
            if pos_raw in POS_MAP:
                entry.pos = pos_raw
                # Grab this subsection's body to find its overview template.
                start = match.end()
                next_heading = re.search(r"^===[^=].*$", section[start:], re.MULTILINE)
                end = start + next_heading.start() if next_heading else len(section)
                body = section[start:end]
                for template_name in (
                    "Deutsch Substantiv Übersicht",
                    "Deutsch Verb Übersicht",
                    "Deutsch Adjektiv Übersicht",
                ):
                    block = _extract_template(body, template_name)
                    if block:
                        entry.templates[template_name] = _parse_template_params(block)
                        break
                return
        # Fallback: no recognised POS heading matched exactly — try a looser
        # search across the whole section for any overview template so we
        # still get grammar data even if the heading regex didn't match.
        for template_name, pos_guess in (
            ("Deutsch Substantiv Übersicht", "Substantiv"),
            ("Deutsch Verb Übersicht", "Verb"),
            ("Deutsch Adjektiv Übersicht", "Adjektiv"),
        ):
            block = _extract_template(section, template_name)
            if block:
                entry.pos = pos_guess
                entry.templates[template_name] = _parse_template_params(block)
                return

    def _parse_pronunciation(self, entry: RawEntry, section: str) -> None:
        ipa_match = re.search(r"\{\{Lautschrift\|([^}|]+)", section)
        if ipa_match:
            entry.ipa = ipa_match.group(1).strip()
        audio_match = re.search(r"\{\{Audio\|([^}|]+)", section)
        if audio_match:
            entry.audio_filename = audio_match.group(1).strip()

    def _parse_examples(self, entry: RawEntry, section: str) -> None:
        body = _extract_section(section, level=4, heading_pattern="Beispiele")
        if not body:
            return
        for line in body.splitlines():
            line = line.strip()
            match = re.match(r":\[\d+[a-z]?\]\s*(.+)", line)
            if match:
                cleaned = strip_wikitext(match.group(1))
                if cleaned:
                    entry.examples.append(cleaned)

    def _parse_synonyms_antonyms(self, entry: RawEntry, section: str) -> None:
        for headings, target in (
                (("Synonyme", "Sinnverwandte Wörter", "Sinnverwandte Woerter"), "synonyms"),
                (("Gegenwörter", "Gegenwoerter", "Antonyme"), "antonyms"),
        ):
            body = None
            for heading in headings:
                body = _extract_section(section, level=4, heading_pattern=heading)
                if body:
                    break
            if not body:
                continue

            words: List[str] = []
            for line in body.splitlines():
                line = line.strip()
                match = re.match(r":\[.*?\]\s*(.+)", line)
                if match:
                    cleaned = strip_wikitext(match.group(1))
                    for piece in cleaned.split(","):
                        piece = piece.strip()
                        if piece:
                            words.append(piece)

            unique_words = list(dict.fromkeys(words))
            if target == "synonyms":
                unique_words = unique_words[:3]
            else:
                unique_words = unique_words[:5]
            setattr(entry, target, unique_words)

    def _parse_translations(self, entry: RawEntry, section: str) -> None:
        body = _extract_section(section, level=4, heading_pattern="Übersetzungen")
        if not body:
            return
        for line in body.splitlines():
            if "{{en}}" in line:
                words = re.findall(r"\{\{Ü\|en\|([^}|]+)", line)
                if not words:
                    words = [
                        strip_wikitext(w)
                        for w in re.findall(r"\[\[([^\]|]+)", line)
                    ]
                entry.english_translations.extend(
                    w.strip() for w in words if w.strip()
                )
            elif "{{fa}}" in line:
                words = re.findall(r"\{\{Ü\|fa\|([^}|]+)", line)
                if not words:
                    words = [
                        strip_wikitext(w)
                        for w in re.findall(r"\[\[([^\]|]+)", line)
                    ]
                entry.persian_translations.extend(
                    w.strip() for w in words if w.strip()
                )
