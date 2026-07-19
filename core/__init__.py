"""Core package for the German Vocabulary Manager.

Contains the modular building blocks used by ``create_excel.py`` and
``autofill.py``:

- ``models``            : plain dataclasses describing a dictionary lookup
- ``http_client``        : a retrying, rate-limited HTTP session
- ``dictionary``         : ``GermanDictionary`` — fetches + parses Wiktionary
- ``parsers``            : ``NounParser`` / ``VerbParser`` / ``AdjectiveParser``
- ``translators``        : Tatoeba / OpenThesaurus / machine-translation fallbacks
- ``audio_downloader``   : resolves + downloads pronunciation audio
- ``excel_manager``      : ``ExcelManager`` — reads/writes the workbook
- ``statistics_manager``  : ``StatisticsManager`` — recomputes the Statistics sheet
"""
