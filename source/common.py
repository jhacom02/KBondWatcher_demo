from __future__ import annotations

import hashlib
import logging
from abc import ABC, abstractmethod
from typing import Optional

logger = logging.getLogger("kbond_watcher")


class SourceReaderError(RuntimeError):
    pass


def message_fingerprint(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


class BaseSourceReader(ABC):
    def __init__(self) -> None:
        self._watermark: set[str] = set()
        self._initialized = False

    @abstractmethod
    def find_source_window(self) -> object:
        raise NotImplementedError

    @abstractmethod
    def get_visible_message_lines(self) -> list[str]:
        raise NotImplementedError

    def initialize_watermark(
        self,
        process_existing_on_start: bool,
        lines: Optional[list[str]] = None,
    ) -> None:
        current = lines if lines is not None else self.get_visible_message_lines()
        if process_existing_on_start:
            self._watermark = set()
        else:
            self._watermark = {message_fingerprint(line) for line in current}
        self._initialized = True
        logger.info(
            "source watermark | existing=%s | process_existing=%s",
            len(current),
            process_existing_on_start,
        )

    def reseed_watermark_from_visible(self) -> None:
        """Mark all currently visible lines as seen (union). Keep prior fingerprints."""
        current = self.get_visible_message_lines()
        for line in current:
            self._watermark.add(message_fingerprint(line))
        self._initialized = True
        logger.info(
            "source watermark reseed | visible=%s | total=%s",
            len(current),
            len(self._watermark),
        )

    def get_new_message_lines(
        self,
        process_existing_on_start: bool = False,
    ) -> list[str]:
        lines = self.get_visible_message_lines()
        if not self._initialized:
            self.initialize_watermark(process_existing_on_start, lines=lines)
            if process_existing_on_start:
                for line in lines:
                    self._watermark.add(message_fingerprint(line))
                return lines
            return []
        new_lines: list[str] = []
        for line in lines:
            fp = message_fingerprint(line)
            if fp in self._watermark:
                continue
            self._watermark.add(fp)
            new_lines.append(line)
        return new_lines

    @abstractmethod
    def diagnose(self, max_messages: int = 200) -> str:
        raise NotImplementedError
