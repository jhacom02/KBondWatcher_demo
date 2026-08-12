from __future__ import annotations

from typing import TYPE_CHECKING

from source_common import BaseSourceReader, SourceReaderError, message_fingerprint
from source_reader_kbond import KbondSourceReader
from source_reader_uia import UiaSourceReader

if TYPE_CHECKING:
    from config import Config

__all__ = [
    "BaseSourceReader",
    "SourceReaderError",
    "message_fingerprint",
    "create_source_reader",
]


def create_source_reader(cfg: "Config") -> BaseSourceReader:
    if cfg.mode in (1, 2):
        return KbondSourceReader(
            source_window_title=cfg.source_window_title,
            source_process_name=cfg.source_process_name,
        )
    if cfg.mode == 3:
        return UiaSourceReader(source_window_title=cfg.source_window_title)
    raise SourceReaderError(f"unsupported MODE: {cfg.mode}")
