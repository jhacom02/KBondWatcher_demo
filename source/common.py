from __future__ import annotations

import hashlib
import logging
from abc import ABC, abstractmethod
from collections import deque
from dataclasses import dataclass
from typing import Optional, Sequence, TypeVar

logger = logging.getLogger("kbond_watcher")

WATERMARK_WINDOW = 2000

_T = TypeVar("_T")


class SourceReaderError(RuntimeError):
    pass


@dataclass(frozen=True)
class SourceLine:
    text: str
    watermark_key: str


def source_line(text: str, watermark_key: Optional[str] = None) -> SourceLine:
    return SourceLine(text=text, watermark_key=text if watermark_key is None else watermark_key)


def as_source_lines(lines: Sequence[str]) -> list[SourceLine]:
    return [source_line(item) for item in lines]


def message_fingerprint(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


def watermark_window(lines: list[_T], window: Optional[int] = None) -> list[_T]:
    if window is None:
        window = WATERMARK_WINDOW
    if window <= 0 or len(lines) <= window:
        return list(lines)
    return list(lines[-window:])


class BaseSourceReader(ABC):
    def __init__(self) -> None:
        self._watermark: set[str] = set()
        self._watermark_order: deque[str] = deque()
        self._initialized = False

    @abstractmethod
    def find_source_window(self) -> object:
        raise NotImplementedError

    @abstractmethod
    def get_visible_message_lines(self) -> list[SourceLine]:
        raise NotImplementedError

    def _remember(self, fp: str, window: Optional[int] = None) -> None:
        if window is None:
            window = WATERMARK_WINDOW
        if fp in self._watermark:
            return
        self._watermark.add(fp)
        self._watermark_order.append(fp)
        while window > 0 and len(self._watermark_order) > window:
            old = self._watermark_order.popleft()
            self._watermark.discard(old)

    def initialize_watermark(
        self,
        process_existing_on_start: bool,
        lines: Optional[list[SourceLine]] = None,
    ) -> None:
        current = lines if lines is not None else self.get_visible_message_lines()
        window = watermark_window(current)
        self._watermark = set()
        self._watermark_order.clear()
        if not process_existing_on_start:
            for line in window:
                self._remember(message_fingerprint(line.watermark_key))
        self._initialized = True
        logger.info(
            "source watermark | existing=%s | dump=%s | window=%s | process_existing=%s",
            len(window),
            len(current),
            WATERMARK_WINDOW,
            process_existing_on_start,
        )

    def reseed_watermark_from_visible(self) -> None:
        current = self.get_visible_message_lines()
        window = watermark_window(current)
        for line in window:
            self._remember(message_fingerprint(line.watermark_key))
        self._initialized = True
        logger.info(
            "source watermark reseed | visible=%s | window=%s | total=%s",
            len(current),
            len(window),
            len(self._watermark),
        )

    def get_new_message_lines(
        self,
        process_existing_on_start: bool = False,
    ) -> list[SourceLine]:
        lines = self.get_visible_message_lines()
        window = watermark_window(lines)
        if not self._initialized:
            self.initialize_watermark(process_existing_on_start, lines=lines)
            if process_existing_on_start:
                new_lines: list[SourceLine] = []
                for line in window:
                    fp = message_fingerprint(line.watermark_key)
                    self._remember(fp)
                    new_lines.append(line)
                return new_lines
            return []
        new_lines = []
        for line in window:
            fp = message_fingerprint(line.watermark_key)
            if fp in self._watermark:
                continue
            self._remember(fp)
            new_lines.append(line)
        return new_lines

    @abstractmethod
    def diagnose(self, max_messages: int = 200) -> str:
        raise NotImplementedError
