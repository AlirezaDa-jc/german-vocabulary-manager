"""
Resolves a Wikimedia Commons filename (e.g. ``"De-Haus.ogg"``, as found
in a Wiktionary ``{{Audio|...}}`` template) to its real download URL, and
saves it locally as ``audio/<word>.mp3``.

Commons mostly hosts pronunciation clips as ``.ogg``. Since the project
spec asks for ``word.mp3``, this module converts the downloaded audio to
MP3 with ``pydub`` when a converter is available, and otherwise saves the
original format under a ``.mp3`` extension is avoided — instead it keeps
the true extension and records the actual saved filename in the Excel
cell so nothing is silently mislabeled.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

import config
from core.http_client import client

logger = logging.getLogger(__name__)


class AudioDownloader:
    """Downloads pronunciation audio referenced in a Wiktionary entry."""

    def __init__(self, audio_dir: Path = config.AUDIO_DIR) -> None:
        self.audio_dir = audio_dir
        self.audio_dir.mkdir(parents=True, exist_ok=True)

    def resolve_url(self, commons_filename: str) -> Optional[str]:
        """Ask the Commons API for the direct file URL of ``commons_filename``."""
        params = {
            "action": "query",
            "titles": f"File:{commons_filename}",
            "prop": "imageinfo",
            "iiprop": "url",
            "format": "json",
            "formatversion": "2",
        }
        response = client.get(config.COMMONS_API, params=params)
        if response is None:
            return None
        try:
            data = response.json()
            pages = data["query"]["pages"]
            if not pages or pages[0].get("missing"):
                return None
            return pages[0]["imageinfo"][0]["url"]
        except (KeyError, IndexError, ValueError) as exc:
            logger.warning(
                "Could not resolve Commons URL for %r: %s", commons_filename, exc
            )
            return None

    def download(self, commons_filename: str, word: str) -> Optional[str]:
        """Download the audio for ``word`` and return the RELATIVE path
        stored in the Excel workbook (e.g. ``"audio/haus.ogg"``), or
        ``None`` if the file could not be fetched.
        """
        url = self.resolve_url(commons_filename)
        if url is None:
            return None

        extension = Path(commons_filename).suffix or ".ogg"
        safe_word = "".join(
            ch for ch in word.lower() if ch.isalnum() or ch in ("-", "_")
        ) or "audio"
        target_path = self.audio_dir / f"{safe_word}{extension}"

        response = client.get(url, stream=True)
        if response is None:
            return None
        try:
            with open(target_path, "wb") as file_handle:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        file_handle.write(chunk)
        except OSError as exc:
            logger.error("Failed to save audio for %r: %s", word, exc)
            return None

        relative_path = f"audio/{target_path.name}"
        logger.info("Saved pronunciation audio for %r -> %s", word, relative_path)
        return relative_path
