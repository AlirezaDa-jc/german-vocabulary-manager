"""
Plain dataclasses used to move parsed dictionary data between
``GermanDictionary``, the parsers, and ``ExcelManager`` without passing
raw dicts around everywhere.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class RawEntry:
    """Unprocessed data pulled from Wiktionary for a single headword.

    ``templates`` maps a template name (e.g. "Deutsch Substantiv Übersicht")
    to its raw parameter dict as found in the wikitext of the FIRST matching
    German part-of-speech section.
    """

    word: str
    found: bool = False
    pos: Optional[str] = None                  # "Substantiv", "Verb", ...
    wikitext: str = ""
    templates: dict = field(default_factory=dict)
    examples: List[str] = field(default_factory=list)
    synonyms: List[str] = field(default_factory=list)
    antonyms: List[str] = field(default_factory=list)
    english_translations: List[str] = field(default_factory=list)
    persian_translations: List[str] = field(default_factory=list)
    ipa: Optional[str] = None
    audio_filename: Optional[str] = None


@dataclass
class NounInfo:
    word: str
    article: Optional[str] = None
    plural: Optional[str] = None
    gender: Optional[str] = None
    genitive: Optional[str] = None
    dative: Optional[str] = None
    accusative: Optional[str] = None
    english: Optional[str] = None
    ipa: Optional[str] = None
    example: Optional[str] = None
    example_translation: Optional[str] = None
    synonyms: List[str] = field(default_factory=list)
    antonyms: List[str] = field(default_factory=list)
    audio_filename: Optional[str] = None
    notes: List[str] = field(default_factory=list)


@dataclass
class VerbInfo:
    infinitive: str
    ich: Optional[str] = None
    du: Optional[str] = None
    er_sie_es: Optional[str] = None
    wir: Optional[str] = None
    ihr: Optional[str] = None
    sie_Sie: Optional[str] = None
    praeteritum: Optional[str] = None
    perfekt: Optional[str] = None
    partizip_ii: Optional[str] = None
    hilfsverb: Optional[str] = None
    trennbar: Optional[bool] = None
    reflexiv: Optional[bool] = None
    irregular: Optional[bool] = None
    meaning: Optional[str] = None
    example: Optional[str] = None
    ipa: Optional[str] = None
    audio_filename: Optional[str] = None
    notes: List[str] = field(default_factory=list)


@dataclass
class AdjectiveInfo:
    positive: str
    comparative: Optional[str] = None
    superlative: Optional[str] = None
    meaning: Optional[str] = None
    example: Optional[str] = None
    ipa: Optional[str] = None
    audio_filename: Optional[str] = None
    notes: List[str] = field(default_factory=list)
