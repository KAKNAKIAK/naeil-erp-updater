"""Pacing rules for SellConnect terminal macro execution."""

from __future__ import annotations

from dataclasses import dataclass
import re


PROMPT_RE = re.compile(r"(?m)^\s*>\s*$")
LOADING_RE = re.compile(r"(\|\||\.\.\.|loading|processing|조회중)", re.IGNORECASE)


@dataclass(frozen=True)
class TopasPacingPolicy:
    """Decide when the next TOPAS command may be sent."""

    min_delay_seconds: float = 0.35
    max_wait_seconds: float = 25.0
    stable_prompt_seconds: float = 0.4

    def is_ready_for_next_command(self, screen_text: str) -> bool:
        tail = "\n".join(screen_text.replace("\r\n", "\n").splitlines()[-8:])
        return bool(PROMPT_RE.search(tail)) and not LOADING_RE.search(tail)


def split_completed_blocks(screen_text: str) -> list[str]:
    """Return terminal chunks that appear to contain completed AN responses."""

    lines = screen_text.replace("\r\n", "\n").splitlines()
    starts = [
        idx
        for idx, line in enumerate(lines)
        if re.search(r"^\s*AN\d{1,2}[A-Z]{3}[A-Z]{6}/A[A-Z0-9]{2}\d{1,4}", line)
    ]
    if not starts:
        return []
    starts.append(len(lines))
    blocks: list[str] = []
    for idx in range(len(starts) - 1):
        chunk = "\n".join(lines[starts[idx] : starts[idx + 1]]).strip()
        if "AMADEUS AVAILABILITY" in chunk.upper():
            blocks.append(chunk)
    return blocks
