"""Helpers for composing text paragraph JSX snippets."""
from __future__ import annotations


def append_text_block(add_lines: list[str], style_name: str, escaped_text: str) -> None:
    """Append the default paragraph handling snippet for a plain text block."""
    add_lines.append("__ensureLayoutDefault();")
    add_lines.append(f'addParaWithNotes(story, "{style_name}", "{escaped_text}");')
