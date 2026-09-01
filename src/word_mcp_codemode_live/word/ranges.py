"""Validated targeting of ranges in live Microsoft Word documents."""

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True, slots=True)
class ResolvedRange:
    """A Word range and a stable description of how it was selected."""

    com_range: Any
    label: str


def character_range(document: Any, start: int, end: int) -> ResolvedRange:
    """Resolve a non-empty zero-based character range within a document."""
    content_end = int(document.Content.End)
    if start < 0 or end <= start or end > content_end:
        raise ValueError(f"Range {start}-{end} is outside document bounds 0-{content_end}")
    return ResolvedRange(document.Range(start, end), f"{start}-{end}")


def paragraph_range(
    document: Any, start_paragraph: int, end_paragraph: int | None = None
) -> ResolvedRange:
    """Resolve an inclusive one-based paragraph range."""
    final_paragraph = end_paragraph or start_paragraph
    total_paragraphs = int(document.Paragraphs.Count)
    if (
        start_paragraph < 1
        or final_paragraph < start_paragraph
        or final_paragraph > total_paragraphs
    ):
        raise ValueError(
            f"Paragraph range {start_paragraph}-{final_paragraph} is outside document bounds "
            f"1-{total_paragraphs}"
        )
    start = int(document.Paragraphs(start_paragraph).Range.Start)
    end = int(document.Paragraphs(final_paragraph).Range.End)
    return ResolvedRange(
        document.Range(start, end), f"paragraphs {start_paragraph}-{final_paragraph}"
    )


def character_or_paragraph_range(
    document: Any,
    *,
    start: int | None,
    end: int | None,
    start_paragraph: int | None,
    end_paragraph: int | None,
) -> ResolvedRange:
    """Resolve exactly one of the supported public range-addressing modes."""
    has_characters = start is not None or end is not None
    has_paragraphs = start_paragraph is not None or end_paragraph is not None
    if has_characters and has_paragraphs:
        raise ValueError("Use character positions or paragraph indexes, not both")
    if has_characters:
        if start is None or end is None:
            raise ValueError("Both start and end character positions are required")
        return character_range(document, start, end)
    if start_paragraph is None:
        raise ValueError("Provide start/end character positions or a start_paragraph")
    return paragraph_range(document, start_paragraph, end_paragraph)
