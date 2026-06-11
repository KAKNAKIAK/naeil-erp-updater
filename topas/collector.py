"""Small TOPAS collection helpers shared by the V5 GUI.

The Selenium collection loop stays in gui.py for the MVP because v4 already has
the live-tested browser handling. This module owns the data shaping and backup
policy around collected raw blocks.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re


@dataclass(frozen=True)
class RawBackup:
    path: Path
    block_count: int
    saved_at: str


def join_raw_blocks(blocks) -> str:
    return "\n\n".join(str(block).strip() for block in blocks if str(block).strip())


def save_raw_backup(
    blocks,
    raw_dir: str | Path,
    route: str = "UNKNOWN",
    direction: str = "raw",
) -> RawBackup:
    raw_text = join_raw_blocks(blocks)
    target_dir = Path(raw_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    saved_at = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_route = _safe_filename(route or "UNKNOWN")
    safe_direction = _safe_filename(direction or "raw")
    path = target_dir / f"{saved_at}_{safe_route}_{safe_direction}.txt"
    path.write_text(raw_text, encoding="utf-8")
    return RawBackup(path=path, block_count=len([b for b in blocks if str(b).strip()]), saved_at=saved_at)


def _safe_filename(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("_") or "raw"
