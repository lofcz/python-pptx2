"""Unit tests for :mod:`pptx2.design.blocks`."""

from __future__ import annotations

import io

import pytest
from PIL import Image

from pptx2 import BBox, Presentation, add_bullets, add_card, add_picture_fit
from pptx2.design.blocks import Card, FittedPicture
from pptx2.enum.text import PP_BULLET_TYPE
from pptx2.util import Inches, Pt


@pytest.fixture
def slide():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    return prs.slides.add_slide(prs.slide_layouts[6])


def _png(width: int, height: int) -> io.BytesIO:
    buf = io.BytesIO()
    Image.new("RGB", (width, height), (40, 120, 90)).save(buf, format="PNG")
    buf.seek(0)
    return buf


class DescribeAddBullets:
    def it_writes_one_paragraph_per_item_with_real_bullets(self, slide):
        box = add_bullets(
            slide, BBox.from_inches(1, 1, 6, 4), items=["alpha", "beta", "gamma"], size_pt=20
        )
        paras = box.text_frame.paragraphs
        assert [p.text for p in paras] == ["alpha", "beta", "gamma"]
        assert all(p.bullet.type is PP_BULLET_TYPE.CHARACTER for p in paras)
        assert all(p.bullet.char == "•" for p in paras)
        # Hanging indent: marker at the edge, text indented.
        pPr = paras[0]._p.pPr
        assert pPr.marL > 0
        assert pPr.indent == -pPr.marL

    def it_keeps_the_requested_size_when_the_text_fits(self, slide):
        box = add_bullets(slide, BBox.from_inches(1, 1, 8, 5), items=["short"], size_pt=24)
        assert box.text_frame.paragraphs[0].runs[0].font.size == Pt(24)

    def it_numbers_continuously_when_numbered(self, slide):
        box = add_bullets(
            slide, BBox.from_inches(1, 1, 6, 4), items=["one", "two", "three"], numbered=True
        )
        paras = box.text_frame.paragraphs
        assert all(p.bullet.type is PP_BULLET_TYPE.NUMBERED for p in paras)
        assert paras[0].bullet.start_at == 1
        # Later items must not restart the sequence.
        assert paras[1].bullet.start_at is None
        assert paras[2].bullet.start_at is None

    def it_adds_space_between_items_but_not_after_the_last(self, slide):
        box = add_bullets(slide, BBox.from_inches(1, 1, 6, 4), items=["a", "b"], gap_pt=10)
        paras = box.text_frame.paragraphs
        assert paras[0].space_after == Pt(10)
        assert paras[1].space_after is None

    def it_accepts_positional_geometry(self, slide):
        box = add_bullets(slide, Inches(1), Inches(1), Inches(4), Inches(2), items=["x"])
        assert box.left == Inches(1)
        assert box.width == Inches(4)

    def it_rejects_an_empty_list(self, slide):
        with pytest.raises(ValueError):
            add_bullets(slide, BBox.from_inches(1, 1, 4, 2), items=[])


class DescribeAddCard:
    def it_returns_a_surface_plus_title_and_body(self, slide):
        card = add_card(
            slide,
            BBox.from_inches(1, 1, 4, 3),
            title="Heading",
            body="Some explanatory body text.",
            fill="#F1F5F9",
        )
        assert isinstance(card, Card)
        assert card.title_box.text_frame.text == "Heading"
        assert card.body_box.text_frame.text == "Some explanatory body text."
        assert len(card.shapes) == 3

    def it_is_a_flat_surface_with_no_outline_and_no_shadow_by_default(self, slide):
        card = add_card(slide, BBox.from_inches(1, 1, 4, 3), title="T", body="B")
        surface = card.card
        assert str(surface.fill.fore_color.rgb) == "F1F5F9"
        # No outline unless asked for one.
        assert surface.line.fill.type is not None  # background() fill was applied
        # No theme drop-shadow left behind.
        assert surface.shadow.blur_radius is None
        assert surface.shadow.distance is None

    def it_draws_an_outline_when_a_line_colour_is_given(self, slide):
        card = add_card(slide, BBox.from_inches(1, 1, 4, 3), title="T", line="#CBD5E1")
        assert str(card.card.line.color.rgb) == "CBD5E1"

    def it_keeps_text_inside_the_padding(self, slide):
        bb = BBox.from_inches(1, 1, 4, 3)
        card = add_card(slide, bb, title="T", body="B", pad_pt=24)
        pad = Pt(24)
        assert card.inner == bb.inset(all=pad)
        assert card.title_box.left == bb.left + pad
        assert card.title_box.width == bb.width - 2 * pad
        assert card.body_box.top > card.title_box.top
        assert card.body_box.top + card.body_box.height <= bb.top + bb.height - pad

    def it_renders_a_list_body_as_bullets(self, slide):
        card = add_card(
            slide, BBox.from_inches(1, 1, 5, 3), title="T", body=["one", "two"], numbered=True
        )
        paras = card.body_box.text_frame.paragraphs
        assert [p.text for p in paras] == ["one", "two"]
        assert paras[0].bullet.type is PP_BULLET_TYPE.NUMBERED

    def it_tags_all_its_shapes_as_one_lint_group(self, slide):
        card = add_card(slide, BBox.from_inches(1, 1, 4, 3), title="T", body="B")
        groups = {s.lint_group for s in card.shapes}
        assert len(groups) == 1
        assert next(iter(groups)).startswith("card@")

    def it_does_not_trip_the_collision_linter(self, slide):
        add_card(slide, BBox.from_inches(1, 1, 4, 3), title="T", body="B")
        report = slide.lint()
        assert not [i for i in report.issues if i.code == "ShapeCollision"]

    def it_works_with_only_a_body(self, slide):
        card = add_card(slide, BBox.from_inches(1, 1, 4, 3), body="Just a body")
        assert card.title_box is None
        assert card.body_box.text_frame.text == "Just a body"

    def it_caps_the_corner_radius_on_thin_cards(self, slide):
        card = add_card(slide, BBox.from_inches(1, 1, 6, 0.3), title="T", radius_pt=40)
        assert int(card.card.corner_radius) <= int(Inches(0.3)) // 2


class DescribeAddPictureFit:
    def it_letterboxes_a_tall_image_inside_a_wide_box(self, slide):
        bb = BBox.from_inches(1, 1, 6, 3)
        fitted = add_picture_fit(slide, _png(300, 600), bb, mode="contain")
        assert isinstance(fitted, FittedPicture)
        pic = fitted.picture
        assert pic.height == bb.height
        assert pic.width == Inches(1.5)
        # Centred horizontally.
        assert abs((pic.left + pic.width // 2) - (bb.left + bb.width // 2)) <= 1
        assert pic.crop_left == 0
        assert pic.crop_top == 0

    def it_honours_align_in_contain_mode(self, slide):
        bb = BBox.from_inches(1, 1, 6, 3)
        fitted = add_picture_fit(slide, _png(300, 600), bb, mode="contain", align="right")
        assert fitted.picture.left + fitted.picture.width == bb.left + bb.width

    def it_crops_symmetrically_in_cover_mode(self, slide):
        bb = BBox.from_inches(1, 1, 6, 3)  # 2:1 box
        fitted = add_picture_fit(slide, _png(400, 400), bb, mode="cover")  # square image
        pic = fitted.picture
        assert (pic.left, pic.top, pic.width, pic.height) == tuple(bb)
        assert pic.crop_left == 0
        assert pic.crop_right == 0
        assert pic.crop_top == pytest.approx(0.25)
        assert pic.crop_bottom == pytest.approx(0.25)

    def it_reserves_room_for_a_caption(self, slide):
        bb = BBox.from_inches(1, 1, 6, 4)
        fitted = add_picture_fit(slide, _png(600, 400), bb, caption="Figure 1")
        assert fitted.caption_box is not None
        assert fitted.caption_box.text_frame.text == "Figure 1"
        assert fitted.caption_box.top >= fitted.picture.top + fitted.picture.height
        assert fitted.caption_box.top + fitted.caption_box.height <= bb.top + bb.height

    def it_rejects_unknown_modes(self, slide):
        with pytest.raises(ValueError):
            add_picture_fit(slide, _png(10, 10), BBox.from_inches(1, 1, 2, 2), mode="stretch")
