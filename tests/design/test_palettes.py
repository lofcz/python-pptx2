"""Unit tests for :mod:`pptx2.design.palettes`."""

from __future__ import annotations

import re

import pytest

from pptx2 import PALETTES, Palette, palette
from pptx2.lint import _contrast_ratio  # type: ignore[attr-defined]

HEX = re.compile(r"^#[0-9A-F]{6}$")


def _rgb(value: str) -> tuple[int, int, int]:
    v = value.lstrip("#")
    return int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)


class DescribePalettes:
    def it_ships_only_well_formed_hex_values(self):
        for name, p in PALETTES.items():
            assert p.name == name
            for value in p:
                assert HEX.match(value), (name, value)

    @pytest.mark.parametrize("name", sorted(PALETTES))
    def it_keeps_body_text_readable_on_paper_and_surfaces(self, name):
        p = PALETTES[name]
        assert _contrast_ratio(_rgb(p.ink), _rgb(p.paper)) >= 7.0
        assert _contrast_ratio(_rgb(p.ink), _rgb(p.surface)) >= 7.0
        assert _contrast_ratio(_rgb(p.ink), _rgb(p.accent_soft)) >= 4.5
        assert _contrast_ratio(_rgb(p.muted), _rgb(p.paper)) >= 4.5
        assert _contrast_ratio(_rgb(p.accent_ink), _rgb(p.accent)) >= 4.5

    def it_looks_up_by_name_and_lists_options_on_miss(self):
        assert palette("slate") is PALETTES["slate"]
        with pytest.raises(KeyError, match="slate"):
            palette("nope")

    def it_derives_a_readable_dark_variant(self):
        light = PALETTES["slate"]
        dark = light.dark()
        assert isinstance(dark, Palette)
        assert dark.paper == light.ink
        assert dark.ink == light.paper
        assert dark.accent == light.accent
        assert _contrast_ratio(_rgb(dark.ink), _rgb(dark.paper)) >= 7.0
        assert _contrast_ratio(_rgb(dark.muted), _rgb(dark.paper)) >= 4.5
        assert _contrast_ratio(_rgb(dark.ink), _rgb(dark.surface)) >= 7.0
