"""Validation and conversion helpers for values passed to Word COM."""

import re


def rgb_hex_to_word(value: str, *, field_name: str = "color") -> int:
    """Convert ``#RRGGBB`` text to the integer representation used by Word COM."""
    normalized = value.removeprefix("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", normalized):
        raise ValueError(f"{field_name} must be a six-digit RGB hex value")
    red, green, blue = (int(normalized[index : index + 2], 16) for index in (0, 2, 4))
    return red + (green << 8) + (blue << 16)
