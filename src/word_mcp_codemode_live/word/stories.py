"""Traversal of Microsoft Word story ranges."""

import logging
from dataclasses import dataclass
from typing import Any, Final

LOGGER = logging.getLogger(__name__)

STORY_NAMES: Final[dict[int, str]] = {
    1: "main_text",
    2: "footnotes",
    3: "endnotes",
    4: "comments",
    5: "text_frames",
    6: "even_pages_header",
    7: "primary_header",
    8: "even_pages_footer",
    9: "primary_footer",
    10: "first_page_header",
    11: "first_page_footer",
    12: "footnote_separator",
    13: "footnote_continuation_separator",
    14: "footnote_continuation_notice",
    15: "endnote_separator",
    16: "endnote_continuation_separator",
    17: "endnote_continuation_notice",
}


@dataclass(frozen=True, slots=True)
class StoryRange:
    """One populated range in a Word story chain."""

    story_type: int
    name: str
    instance_index: int
    com_range: Any


def collect_story_ranges(document: Any) -> tuple[list[StoryRange], list[str]]:
    """Collect populated story ranges and names unavailable in this document.

    Word exposes section-specific headers, footers, and text frames through linked
    ``NextStoryRange`` chains. The cycle guard prevents a malformed COM proxy from
    hanging an MCP request indefinitely.
    """
    populated: list[StoryRange] = []
    unavailable: list[str] = []

    for story_type, name in STORY_NAMES.items():
        try:
            current = document.StoryRanges(story_type)
        except Exception as exc:
            LOGGER.debug("Word story %s (%s) is unavailable: %s", story_type, name, exc)
            unavailable.append(name)
            continue

        instance_index = 1
        while current is not None:
            populated.append(StoryRange(story_type, name, instance_index, current))
            instance_index += 1
            if instance_index > 10_000:
                raise RuntimeError(f"Word returned a cyclic story chain for story {story_type}")
            try:
                current = current.NextStoryRange
            except Exception as exc:
                LOGGER.debug(
                    "Could not follow Word story %s (%s) after instance %s: %s",
                    story_type,
                    name,
                    instance_index - 1,
                    exc,
                )
                current = None

    return populated, unavailable
