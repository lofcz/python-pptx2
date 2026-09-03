"""Curated slide palettes.

A deck reads as designed when every colour on it comes from one small,
deliberate set: a paper colour for the background, an ink for text, one
accent that carries emphasis, a soft tint of that accent for surfaces,
and a muted tone for captions and rules.  :class:`Palette` names those
roles so a script can ask for ``P.accent`` instead of remembering a hex
string, and :data:`PALETTES` ships a handful of combinations that are
known to sit well together and pass contrast checks for projected text.

Every value is a ``"#RRGGBB"`` string, so it can be passed straight to
``add_text(color=...)``, ``shape.fill_hex(...)``, ``cell.format(fill=...)``
and the rest of the hex-accepting surface.

Typical use::

    from pptx2.design.palettes import PALETTES

    P = PALETTES["slate"]
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = P.paper
    slide.shapes.add_text(bb, text="Title", size_pt=36, bold=True, color=P.ink)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, *cell).fill_hex(P.surface)

``P.dark()`` flips the same hues onto a dark background (``paper`` becomes
the ink colour and vice-versa) for title and section slides, so a deck can
alternate light content slides with dark feature slides without picking
new colours.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Dict, Iterator

__all__ = ["Palette", "PALETTES", "palette"]


@dataclass(frozen=True)
class Palette:
    """One coherent colour set for a deck.

    Roles:

    * ``paper`` — slide background.
    * ``surface`` — card / panel fill that sits *on* the paper; a step
      away from ``paper`` so a card reads as a surface without a border.
    * ``line`` — hairline rules, table grid, card outline when one is
      wanted.
    * ``ink`` — headings and body text.
    * ``muted`` — captions, sources, secondary labels.
    * ``accent`` — the one emphasis colour: key numbers, the active step
      of a process, a highlighted word.
    * ``accent_soft`` — a light tint of ``accent`` for a callout fill or a
      highlighted row; text on it stays ``ink``.
    * ``accent_ink`` — text colour that is readable *on* ``accent``
      (white on a dark accent, ink on a light one).
    """

    name: str
    paper: str
    surface: str
    line: str
    ink: str
    muted: str
    accent: str
    accent_soft: str
    accent_ink: str = "#FFFFFF"

    def dark(self) -> "Palette":
        """The same hues arranged for a dark slide (title, section, closing).

        ``paper`` and ``ink`` swap; ``surface`` becomes a slightly lifted
        dark panel; ``muted`` lightens so captions stay legible; the
        accent is kept because a saturated accent reads well on both.
        """
        return replace(
            self,
            name=f"{self.name}-dark",
            paper=self.ink,
            surface=_mix(self.ink, self.paper, 0.10),
            line=_mix(self.ink, self.paper, 0.25),
            ink=self.paper,
            muted=_mix(self.paper, self.ink, 0.35),
            accent_soft=_mix(self.ink, self.accent, 0.35),
        )

    def __iter__(self) -> Iterator[str]:
        yield from (
            self.paper,
            self.surface,
            self.line,
            self.ink,
            self.muted,
            self.accent,
            self.accent_soft,
            self.accent_ink,
        )


def _hex_to_rgb(value: str) -> tuple[int, int, int]:
    v = value.lstrip("#")
    if len(v) != 6:
        raise ValueError(f"expected #RRGGBB, got {value!r}")
    return int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)


def _rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    return "#%02X%02X%02X" % rgb


def _mix(a: str, b: str, t: float) -> str:
    """Linear blend of two hex colours; ``t=0`` → *a*, ``t=1`` → *b*."""
    ra, ga, ba = _hex_to_rgb(a)
    rb, gb, bb = _hex_to_rgb(b)
    t = max(0.0, min(1.0, float(t)))
    return _rgb_to_hex(
        (
            int(round(ra + (rb - ra) * t)),
            int(round(ga + (gb - ga) * t)),
            int(round(ba + (bb - ba) * t)),
        )
    )


PALETTES: Dict[str, Palette] = {
    # Calm blue-grey. The safe default for any subject.
    "slate": Palette(
        name="slate",
        paper="#FFFFFF",
        surface="#F1F5F9",
        line="#CBD5E1",
        ink="#0F172A",
        muted="#64748B",
        accent="#2563EB",
        accent_soft="#DBEAFE",
    ),
    # Warm paper with a deep teal accent. Humanities, languages, civics.
    "linen": Palette(
        name="linen",
        paper="#FBF8F3",
        surface="#F2ECE2",
        line="#D9CFC0",
        ink="#1F2933",
        muted="#7B6F63",
        accent="#0F766E",
        accent_soft="#CCEDE8",
    ),
    # Green on near-white. Biology, geography, sustainability.
    "forest": Palette(
        name="forest",
        paper="#FFFFFF",
        surface="#F0F5F1",
        line="#C8D6CB",
        ink="#14261B",
        muted="#5F7365",
        accent="#15803D",
        accent_soft="#DCFCE7",
    ),
    # Plum accent on cool paper. Literature, arts, history.
    "plum": Palette(
        name="plum",
        paper="#FFFFFF",
        surface="#F5F1F8",
        line="#D9CFE3",
        ink="#231A2B",
        muted="#6F6478",
        accent="#7E22CE",
        accent_soft="#EDE0F7",
    ),
    # Warm orange accent. Physics, technology, energy — anything that should feel active.
    "ember": Palette(
        name="ember",
        paper="#FFFFFF",
        surface="#FBF3EC",
        line="#E7D4C4",
        ink="#1C1917",
        muted="#78716C",
        accent="#C2410C",
        accent_soft="#FFEDD5",
    ),
    # Deep navy accent on white. Mathematics, economics, formal topics.
    "navy": Palette(
        name="navy",
        paper="#FFFFFF",
        surface="#EEF2F7",
        line="#C9D3E0",
        ink="#0B1B33",
        muted="#5B6B82",
        accent="#1E3A8A",
        accent_soft="#DCE4F5",
    ),
    # Dark deck: charcoal paper with a mint accent. Use as the base palette
    # when the whole deck should be dark, or take PALETTES["slate"].dark()
    # for a single dark slide inside a light deck.
    "graphite": Palette(
        name="graphite",
        paper="#111827",
        surface="#1F2937",
        line="#374151",
        ink="#F9FAFB",
        muted="#9CA3AF",
        accent="#34D399",
        accent_soft="#1F3D34",
        accent_ink="#062017",
    ),
}


def palette(name: str) -> Palette:
    """Return the curated palette called *name* (``KeyError`` lists the options)."""
    try:
        return PALETTES[name]
    except KeyError:
        raise KeyError(
            f"unknown palette {name!r}; choose one of {', '.join(sorted(PALETTES))}"
        ) from None
