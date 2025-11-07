"""Helpers for TABLE -> JSX snippets."""
from __future__ import annotations
import json


def append_table_block(add_lines: list[str], table_obj: dict) -> None:
    """Append the table-rendering snippet for a normalized table object."""
    payload = json.dumps(table_obj, ensure_ascii=False)
    add_lines.append(f"addTableHiFi({payload});")
