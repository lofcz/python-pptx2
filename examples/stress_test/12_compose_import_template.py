"""Composition torture: build separate decks, import slides across decks with
both merge_master strategies, then re-skin everything with apply_template.
"""

from __future__ import annotations

from pathlib import Path

from _util import blank, deck, save

from pptx2 import Presentation
from pptx2.compose import from_spec
from pptx2.dml.color import RGBColor
from pptx2.enum.dml import MSO_THEME_COLOR
from pptx2.enum.shapes import MSO_SHAPE
from pptx2.util import Inches


def _low_level_import(prs, src_slide, merge_master="dedupe"):
    """Import via the part-level engine (keeps the old dedupe/clone knobs)."""
    from pptx2._slide_importer import import_slide

    return import_slide(src_slide.part, prs.part, merge_master=merge_master)

ASSETS = Path(__file__).parent / "_assets"


def _source_deck():
    """A small deck with rich slides to clone from."""
    src = deck()
    s = blank(src)
    c = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1),
                           Inches(5), Inches(3))
    c.fill.linear_gradient("#0F2D6B", "#4F9DFF", angle=90)
    c.line.fill.background()
    c.text_frame.text = "Imported slide A"
    s2 = blank(src)
    if (ASSETS / "logo.png").exists():
        s2.shapes.add_picture(str(ASSETS / "logo.png"), Inches(2), Inches(2),
                              Inches(3), Inches(3))
    tb = s2.shapes.add_textbox(Inches(1), Inches(0.5), Inches(11), Inches(1))
    tb.text_frame.text = "Imported slide B"
    return src


def _template(path: Path):
    """Author a brand template and save it for apply_template."""
    t = Presentation()
    t.theme.colors[MSO_THEME_COLOR.ACCENT_1] = RGBColor(0xFF, 0x66, 0x00)
    t.theme.colors[MSO_THEME_COLOR.ACCENT_2] = RGBColor(0x12, 0x1E, 0x4D)
    t.theme.fonts.major = "Inter"
    t.theme.fonts.minor = "Inter"
    t.save(path)
    return path


def build():
    # Ensure logo asset exists (10_* may not have run in this process)
    ASSETS.mkdir(exist_ok=True)
    if not (ASSETS / "logo.png").exists():
        from PIL import Image
        Image.new("RGB", (200, 200), (79, 157, 255)).save(ASSETS / "logo.png")

    src = _source_deck()

    # Destination starts from a declarative from_spec deck
    dst = from_spec({
        "slide_size": "16:9",
        "tokens": {"preset": "modern_light"},
        "slides": [
            {"layout": "title", "title": "Composition Test",
             "subtitle": "import_slide + apply_template"},
            {"layout": "bullets", "title": "Agenda",
             "bullets": ["Import slides", "Dedupe masters", "Re-skin"]},
        ],
    })

    # Import with dedupe (default-ish) and clone strategies
    _low_level_import(dst, src.slides[0], merge_master="dedupe")
    _low_level_import(dst, src.slides[1], merge_master="clone")

    # Re-skin against a brand template
    tpl = _template(ASSETS / "_brand_template.pptx")
    dst.apply_template(str(tpl))

    return dst


if __name__ == "__main__":
    save(build(), "12_compose_import_template.pptx")
