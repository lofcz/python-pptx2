"""Unit tests for the new shape-level helpers — fill_hex, line_hex, add_text,
add_arrow, set_text_preserving_format, replace_with, enclosing_container,
content_bbox, find_empty_region, tidy."""

from __future__ import annotations

import pytest

from pptx2 import BBox, Presentation
from pptx2.enum.dml import MSO_LINE_END_TYPE
from pptx2.enum.shapes import MSO_CONNECTOR_TYPE, MSO_SHAPE
from pptx2.util import Inches, Pt


@pytest.fixture
def slide():
    prs = Presentation()
    return prs.slides.add_slide(prs.slide_layouts[6])


# ---------------------------------------------------------------------------- fill/line


class DescribeFillHex:
    def it_sets_a_solid_fill_from_hex(self, slide):
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(2), Inches(1))
        rect.fill_hex("#0B5CFF")
        assert str(rect.fill.fore_color.rgb) == "0B5CFF"

    def it_returns_self_for_chaining(self, slide):
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(2), Inches(1))
        result = rect.fill_hex("#0B5CFF")
        assert result is rect

    def it_accepts_hex_without_hash(self, slide):
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(2), Inches(1))
        rect.fill_hex("0B5CFF")
        assert str(rect.fill.fore_color.rgb) == "0B5CFF"


class DescribeLineHex:
    def it_sets_line_color_and_width(self, slide):
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(2), Inches(1))
        rect.line_hex("#0D0D0D", weight_pt=2.0)
        assert str(rect.line.color.rgb) == "0D0D0D"
        assert int(rect.line.width) == int(Pt(2.0))


# ---------------------------------------------------------------------------- text


class DescribeAddText:
    def it_creates_a_textbox_with_text(self, slide):
        tx = slide.shapes.add_text(BBox.from_inches(1, 1, 4, 1), text="Hello")
        assert tx.text_frame.text == "Hello"

    def it_applies_font_styling(self, slide):
        tx = slide.shapes.add_text(
            BBox.from_inches(1, 1, 4, 1),
            text="Hi",
            font="Inter",
            size_pt=24,
            bold=True,
            color="#FF0000",
        )
        run = tx.text_frame.paragraphs[0].runs[0]
        assert run.font.name == "Inter"
        assert run.font.size == Pt(24)
        assert run.font.bold is True
        assert str(run.font.color.rgb) == "FF0000"

    def it_accepts_positional_lengths(self, slide):
        tx = slide.shapes.add_text(
            Inches(1), Inches(2), Inches(3), Inches(1),
            text="positional",
        )
        assert tx.text_frame.text == "positional"

    def it_applies_alignment_short_names(self, slide):
        from pptx2.enum.text import PP_PARAGRAPH_ALIGNMENT

        tx = slide.shapes.add_text(BBox.from_inches(0, 0, 4, 1), text="x", align="center")
        assert tx.text_frame.paragraphs[0].alignment == PP_PARAGRAPH_ALIGNMENT.CENTER

    def it_rejects_unknown_align(self, slide):
        with pytest.raises(ValueError):
            slide.shapes.add_text(BBox.from_inches(0, 0, 4, 1), text="x", align="diagonal")

    def it_accepts_font_family_as_an_alias_for_font(self, slide):
        # -- matplotlib-trained generators habitually write font_family= --
        tx = slide.shapes.add_text(BBox.from_inches(1, 1, 4, 1), text="Hi", font_family="Inter")
        assert tx.text_frame.paragraphs[0].runs[0].font.name == "Inter"

    def it_accepts_identical_font_and_font_family(self, slide):
        tx = slide.shapes.add_text(
            BBox.from_inches(1, 1, 4, 1), text="Hi", font="Inter", font_family="Inter"
        )
        assert tx.text_frame.paragraphs[0].runs[0].font.name == "Inter"

    def it_rejects_conflicting_font_and_font_family(self, slide):
        with pytest.raises(TypeError, match="font"):
            slide.shapes.add_text(
                BBox.from_inches(1, 1, 4, 1), text="Hi", font="Inter", font_family="Roboto"
            )

    def it_accepts_valign_mid_as_an_alias_for_anchor(self, slide):
        from pptx2.enum.text import MSO_VERTICAL_ANCHOR

        tx = slide.shapes.add_text(BBox.from_inches(0, 0, 4, 1), text="x", valign="mid")
        assert tx.text_frame.vertical_anchor == MSO_VERTICAL_ANCHOR.MIDDLE

    def it_rejects_conflicting_anchor_and_valign(self, slide):
        with pytest.raises(TypeError, match="anchor"):
            slide.shapes.add_text(
                BBox.from_inches(0, 0, 4, 1), text="x", anchor="top", valign="bottom"
            )

    def it_absorbs_halign_font_size_and_colour_synonyms(self, slide):
        from pptx2.enum.text import PP_PARAGRAPH_ALIGNMENT

        tx = slide.shapes.add_text(
            BBox.from_inches(1, 1, 4, 1),
            text="Hi",
            halign="center",
            font_size=20,
            colour="#FF0000",
        )
        run = tx.text_frame.paragraphs[0].runs[0]
        assert tx.text_frame.paragraphs[0].alignment == PP_PARAGRAPH_ALIGNMENT.CENTER
        assert run.font.size == Pt(20)
        assert str(run.font.color.rgb) == "FF0000"

    def it_absorbs_matplotlib_ha_va_and_label(self, slide):
        from pptx2.enum.text import MSO_VERTICAL_ANCHOR

        tx = slide.shapes.add_text(
            BBox.from_inches(1, 1, 4, 1), label="Hi", ha="center", va="middle"
        )
        assert tx.text_frame.text == "Hi"
        assert tx.text_frame.vertical_anchor == MSO_VERTICAL_ANCHOR.MIDDLE

    def it_fuzzy_matches_near_miss_kwargs(self, slide):
        from pptx2.enum.text import PP_PARAGRAPH_ALIGNMENT

        tx = slide.shapes.add_text(BBox.from_inches(1, 1, 4, 1), text="x", algn="right")
        assert tx.text_frame.paragraphs[0].alignment == PP_PARAGRAPH_ALIGNMENT.RIGHT

    def it_accepts_geometry_keywords_and_xywh_shorts(self, slide):
        tx = slide.shapes.add_text(
            x=Inches(1), y=Inches(2), w=Inches(3), h=Inches(1), text="xywh"
        )
        assert (tx.left, tx.top, tx.width, tx.height) == (
            Inches(1), Inches(2), Inches(3), Inches(1)
        )

    def it_rejects_unknown_kwargs_with_a_didactic_error(self, slide):
        with pytest.raises(TypeError, match="totally_bogus.*Accepted"):
            slide.shapes.add_text(
                BBox.from_inches(1, 1, 4, 1), text="x", totally_bogus=1
            )

    def it_rejects_partial_geometry_keywords(self, slide):
        with pytest.raises(TypeError, match="left/top/width/height"):
            slide.shapes.add_text(x=Inches(1), y=Inches(2), text="partial")


class DescribeAddArrowSynonyms:
    def it_absorbs_begin_to_colour_and_weight(self, slide):
        arrow = slide.shapes.add_arrow(
            begin=(Inches(1), Inches(1)),
            to=(Inches(3), Inches(1)),
            colour="#00FF00",
            weight=2.5,
        )
        assert arrow.begin_x == Inches(1)
        assert arrow.end_x == Inches(3)

    def it_still_accepts_positional_endpoints(self, slide):
        arrow = slide.shapes.add_arrow((Inches(0), Inches(3)), (Inches(2), Inches(3)))
        assert arrow.begin_x == Inches(0)


class DescribeAgentFriendlyShapeApi:
    """Every shapes.add_* absorbs cross-library kwargs, not just add_text."""

    def it_add_shape_takes_shape_type_and_xywh(self, slide):
        rect = slide.shapes.add_shape(
            shape_type=MSO_SHAPE.RECTANGLE,
            x=Inches(1), y=Inches(1), w=Inches(2), h=Inches(1),
        )
        assert (rect.left, rect.top) == (Inches(1), Inches(1))

    def it_add_picture_takes_image_and_xywh(self, slide, tmp_path):
        import struct, zlib

        def _png(path):
            def chunk(tag, data):
                c = struct.pack(">I", len(data)) + tag + data
                return c + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
            ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
            raw = b"\x00\xff\x00\x00"
            blob = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
                    + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))
            path.write_bytes(blob)

        pic_path = tmp_path / "dot.png"
        _png(pic_path)
        pic = slide.shapes.add_picture(
            image=str(pic_path), x=Inches(1), y=Inches(1), w=Inches(1), h=Inches(1)
        )
        assert pic.left == Inches(1)

    def it_add_table_takes_columns_and_xywh(self, slide):
        gf = slide.shapes.add_table(
            rows=2, columns=3, x=Inches(1), y=Inches(1), w=Inches(6), h=Inches(2)
        )
        assert len(gf.table.columns) == 3

    def it_add_textbox_takes_xywh(self, slide):
        tb = slide.shapes.add_textbox(x=Inches(2), y=Inches(2), w=Inches(3), h=Inches(1))
        assert tb.top == Inches(2)

    def it_add_connector_takes_x1y1x2y2(self, slide):
        cxn = slide.shapes.add_connector(
            MSO_CONNECTOR_TYPE.STRAIGHT,
            x1=Inches(0), y1=Inches(0), x2=Inches(2), y2=Inches(2),
        )
        assert cxn.begin_x == Inches(0)

    def it_unknown_kwarg_error_lists_accepted_names(self, slide):
        with pytest.raises(TypeError, match="Accepted"):
            slide.shapes.add_textbox(Inches(0), Inches(0), Inches(1), Inches(1), bogus=1)

    def it_diagram_recipes_absorb_item_and_colour_synonyms(self, slide):
        from pptx2 import BBox
        from pptx2.diagrams import horizontal_pipeline

        horizontal_pipeline(
            slide, BBox.from_inches(0.5, 1, 12, 1.5), items=["a", "b"], colour="#0D0D0D"
        )


class DescribeSetTextPreservingFormat:
    def it_preserves_font_attributes_across_replacement(self, slide):
        tx = slide.shapes.add_text(
            BBox.from_inches(0, 0, 4, 1),
            text="<TITLE>",
            font="Inter",
            size_pt=18,
            bold=True,
            color="#0B5CFF",
        )
        tx.set_text_preserving_format("New title")
        run = tx.text_frame.paragraphs[0].runs[0]
        assert run.font.name == "Inter"
        assert run.font.size == Pt(18)
        assert run.font.bold is True
        assert str(run.font.color.rgb) == "0B5CFF"
        assert tx.text_frame.text == "New title"

    def it_preserves_format_across_multiple_lines(self, slide):
        tx = slide.shapes.add_text(
            BBox.from_inches(0, 0, 4, 2), text="<x>", font="Inter", size_pt=14, bold=True,
        )
        tx.set_text_preserving_format("a\nb\nc")
        assert len(tx.text_frame.paragraphs) == 3
        for para in tx.text_frame.paragraphs:
            for run in para.runs:
                assert run.font.name == "Inter"
                assert run.font.bold is True

    def it_raises_when_shape_has_no_text_frame(self, slide):
        conn = slide.shapes.add_connector(
            MSO_CONNECTOR_TYPE.STRAIGHT, Inches(0), Inches(0), Inches(1), Inches(1)
        )
        with pytest.raises(ValueError):
            conn.set_text_preserving_format("nope")


# ---------------------------------------------------------------------------- arrow


class DescribeAddArrow:
    def it_creates_a_connector_with_a_tail_arrowhead(self, slide):
        a = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
        b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5), Inches(1), Inches(2), Inches(1))
        arrow = slide.shapes.add_arrow(a, b)
        assert arrow.line.tail_end.type == MSO_LINE_END_TYPE.TRIANGLE

    def it_accepts_bbox_endpoints(self, slide):
        arrow = slide.shapes.add_arrow(
            BBox.from_inches(0, 0, 1, 1),
            BBox.from_inches(5, 0, 1, 1),
        )
        assert arrow is not None

    def it_accepts_xy_tuple_endpoints(self, slide):
        arrow = slide.shapes.add_arrow(
            (Inches(0), Inches(0)),
            (Inches(5), Inches(0)),
        )
        # Begin point is at the start tuple
        assert int(arrow.begin_x) == int(Inches(0))

    def it_rejects_unknown_route(self, slide):
        with pytest.raises(ValueError):
            slide.shapes.add_arrow((Inches(0), Inches(0)), (Inches(5), Inches(0)), route="zigzag")

    def it_rejects_unknown_head(self, slide):
        with pytest.raises(ValueError):
            slide.shapes.add_arrow(
                (Inches(0), Inches(0)), (Inches(5), Inches(0)),
                head="curly",
            )

    def it_rejects_unknown_tail(self, slide):
        with pytest.raises(ValueError):
            slide.shapes.add_arrow(
                (Inches(0), Inches(0)), (Inches(5), Inches(0)),
                tail="splash",
            )

    def it_rejects_unknown_head_size(self, slide):
        with pytest.raises(ValueError):
            slide.shapes.add_arrow(
                (Inches(0), Inches(0)), (Inches(5), Inches(0)),
                head_size="huge",
            )

    def it_normalises_case_for_head(self, slide):
        # "TRIANGLE" should be accepted equivalently to "triangle".
        arrow = slide.shapes.add_arrow(
            (Inches(0), Inches(0)), (Inches(5), Inches(0)),
            head="TRIANGLE",
        )
        assert arrow.line.tail_end.type == MSO_LINE_END_TYPE.TRIANGLE


# ---------------------------------------------------------------------------- picture helpers


class DescribePictureReplaceWith:
    def it_calls_builder_with_the_pictures_bbox(self, slide):
        # add_picture needs an image; use a tiny 1x1 PNG placeholder.
        import io

        png = bytes.fromhex(
            "89504e470d0a1a0a0000000d49484452000000010000000108020000009077"
            "53de0000000c4944415478da6300010000000500010d0a2db40000000049"
            "454e44ae426082"
        )
        pic = slide.shapes.add_picture(io.BytesIO(png), Inches(1), Inches(1), Inches(2), Inches(2))
        called = {}

        def build(slide_arg, bbox_arg):
            called["slide"] = slide_arg
            called["bbox"] = bbox_arg

        pic.replace_with(build)
        assert called["slide"] is slide
        assert isinstance(called["bbox"], BBox)


class DescribePictureEnclosingContainer:
    def _make_picture(self, slide, left, top, width, height):
        import io

        png = bytes.fromhex(
            "89504e470d0a1a0a0000000d49484452000000010000000108020000009077"
            "53de0000000c4944415478da6300010000000500010d0a2db40000000049"
            "454e44ae426082"
        )
        return slide.shapes.add_picture(io.BytesIO(png), left, top, width, height)

    def it_does_not_collapse_onto_a_title_strip(self, slide):
        # Regression for the Codex finding: a "card" rectangle holds a
        # title strip at the top plus the picture below.  ``shrink_around``
        # must trim the title region away while still containing the
        # picture, not push the bottom edge up onto the title strip.
        slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(1), Inches(1), Inches(6), Inches(5),   # outer card
        )
        slide.shapes.add_text(
            BBox.from_inches(1.2, 1.2, 5.6, 0.6),         # title strip
            text="Architecture",
        )
        picture = self._make_picture(
            slide, Inches(1.5), Inches(2), Inches(5), Inches(3.5),
        )
        container = picture.enclosing_container()
        assert container is not None
        # The returned box must enclose the picture, not the title strip.
        assert container.contains(BBox.from_shape(picture))


# ---------------------------------------------------------------------------- slide-level


class DescribeSlideHelpers:
    def it_returns_a_slide_bbox(self, slide):
        bb = slide.slide_bbox()
        assert bb.width == Inches(10)
        assert bb.height == Inches(7.5)

    def it_returns_a_content_bbox(self, slide):
        slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
        slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5), Inches(2), Inches(2), Inches(1))
        bb = slide.content_bbox()
        assert int(bb.left) == int(Inches(1))
        assert int(bb.right) == int(Inches(7))

    def it_returns_none_for_empty_slide(self, slide):
        bb = slide.content_bbox()
        assert bb is None

    def it_finds_empty_region(self, slide):
        slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(3), Inches(1))
        region = slide.find_empty_region(min_width=Inches(0.5), min_height=Inches(0.5))
        assert region is not None
        assert isinstance(region, BBox)

    def it_tidies_off_slide_shapes(self, slide):
        slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(15), Inches(15), Inches(2), Inches(1))
        fixes = slide.tidy()
        assert any("Clamped" in f for f in fixes)

    def it_rejects_unknown_near_type_in_find_empty_region(self, slide):
        with pytest.raises(TypeError):
            slide.find_empty_region(near="not-a-shape")


class DescribeCssHexShorthand:
    def it_expands_three_digit_hex(self, slide):
        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(2), Inches(1)
        )
        rect.fill_hex("#888")
        assert str(rect.fill.fore_color.rgb) == "888888"
