from pathlib import Path

import pytest

from word_mcp_codemode_live.word import images


def _options(path: Path, **overrides):
    values = {
        "image_path": str(path),
        "width_inches": None,
        "height_inches": None,
        "width_pt": None,
        "height_pt": None,
        "alignment": None,
        "wrapping": None,
        "border_style": None,
        "border_width_pt": None,
        "border_color": None,
    }
    values.update(overrides)
    return values


def test_image_options_validate_paths_and_enums(tmp_path: Path) -> None:
    image_path = tmp_path / "image.png"
    image_path.write_bytes(b"not decoded by validation")

    resolved, color = images.validate_options(_options(image_path, wrapping="square"))

    assert resolved == str(image_path)
    assert color == 0
    with pytest.raises(ValueError, match="Unknown wrapping"):
        images.validate_options(_options(image_path, wrapping="around"))


def test_image_dimensions_convert_inches_to_points(tmp_path: Path) -> None:
    width, height = images.dimensions(
        _options(tmp_path / "unused.png", width_inches=2.0, height_inches=1.5)
    )

    assert width == 144.0
    assert height == 108.0
