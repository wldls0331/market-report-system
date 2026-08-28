"""Pass-through comment composer.

v1 writes user facts as-is. Keep this module as the single place to later
plug in an LLM rewrite without changing Excel cell mapping.
"""

from __future__ import annotations

TBN = "TBN"


def compose_comment_lines(facts: str, cell_count: int) -> list[str]:
    text = (facts or "").strip()
    if not text:
        lines = [TBN]
    else:
        lines = [line.rstrip() for line in text.splitlines() if line.strip()]
        if not lines:
            lines = [TBN]
    if len(lines) < cell_count:
        lines = lines + [""] * (cell_count - len(lines))
    elif len(lines) > cell_count:
        head = lines[: cell_count - 1]
        tail = " ".join(lines[cell_count - 1 :])
        lines = head + [tail]
    return lines


def compose_strategy_text(prefix: str, facts: str) -> str:
    text = (facts or "").strip()
    body = text if text else TBN
    if body.startswith(prefix.strip()):
        return body
    return f"{prefix}{body}"
