"""Integration tests for adding and cloning slide layouts.

Covers the public API ported from scanny/python-pptx#1091
(``SlideLayouts.add_layout()``) plus the fork's same-package
``SlideLayouts.clone()``.
"""

from __future__ import annotations

import io

import pytest

from pptx2 import Presentation
from pptx2.enum.shapes import PP_PLACEHOLDER
from pptx2.media import SPEAKER_IMAGE_BYTES
from pptx2.util import Inches

from .round_trip import assert_round_trip, round_trip_diff


def _round_trip(prs: Presentation) -> Presentation:
    """Save *prs* to a BytesIO buffer and reopen it."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


def _add_layout_with_picture(prs: Presentation):
    """Return a layout from *prs* carrying a picture shape (shared image part)."""
    layout = prs.slide_layouts[6]
    image_part, rId = layout.part.get_or_add_image_part(io.BytesIO(SPEAKER_IMAGE_BYTES))
    layout.shapes._spTree.add_pic(
        layout.shapes._next_shape_id,
        "Layout Picture",
        "",
        rId,
        Inches(1),
        Inches(1),
        Inches(2),
        Inches(2),
    )
    return layout, image_part


# ---------------------------------------------------------------------------
# SlideLayouts.add_layout()
# ---------------------------------------------------------------------------


class Describe_add_layout:
    def it_appends_a_blank_layout_to_the_master(self):
        prs = Presentation()
        layouts = prs.slide_layouts
        original_count = len(layouts)

        layout = layouts.add_layout()

        assert len(layouts) == original_count + 1
        assert layouts[-1] is layout
        assert layout.slide_master is prs.slide_masters[0]
        assert len(layout.shapes) == 0
        assert layout.name == "Layout 13"

    def it_assigns_a_distinct_name_and_layout_id(self):
        prs = Presentation()
        layouts = prs.slide_layouts

        first = layouts.add_layout()
        second = layouts.add_layout()

        assert first.name != second.name
        sldLayoutIdLst = prs.slide_masters[0]._element.get_or_add_sldLayoutIdLst()
        ids = [entry.id for entry in sldLayoutIdLst.sldLayoutId_lst]
        assert len(ids) == len(set(ids))
        assert all(id_ >= 2147483648 for id_ in ids)

    def it_accepts_placeholders_and_slides_can_use_the_layout(self):
        prs = Presentation()
        layout = prs.slide_layouts.add_layout("My Layout")

        title = layout.shapes.add_placeholder(PP_PLACEHOLDER.TITLE)
        body = layout.shapes.add_placeholder(PP_PLACEHOLDER.BODY)
        assert title in layout.placeholders
        assert body in layout.placeholders

        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = "New layout in action"
        body_ph = slide.placeholders[1]
        body_ph.text_frame.text = "Body text"

        reopened = _round_trip(prs)
        assert len(reopened.slide_layouts) == 12
        my_layout = reopened.slide_layouts.get_by_name("My Layout")
        assert my_layout is not None
        assert reopened.slides[0].shapes.title.text == "New layout in action"
        assert reopened.slides[0].slide_layout.name == "My Layout"


# ---------------------------------------------------------------------------
# SlideLayouts.clone()
# ---------------------------------------------------------------------------


class Describe_clone:
    def it_appends_a_deep_copy_of_the_source_layout(self):
        prs = Presentation()
        layouts = prs.slide_layouts
        source = layouts[5]
        original_count = len(layouts)

        clone = layouts.clone(source)

        assert len(layouts) == original_count + 1
        assert layouts[-1] is clone
        assert clone is not source
        assert clone.name == "%s 2" % source.name
        assert [(ph.element.ph_type, ph.element.ph_idx) for ph in clone.placeholders] == [
            (ph.element.ph_type, ph.element.ph_idx) for ph in source.placeholders
        ]
        assert len(clone.shapes) == len(source.shapes)

    def it_gives_the_clone_a_distinct_layout_id(self):
        prs = Presentation()
        layouts = prs.slide_layouts
        layouts.clone(layouts[5])
        layouts.clone(layouts[5])

        sldLayoutIdLst = prs.slide_masters[0]._element.get_or_add_sldLayoutIdLst()
        ids = [entry.id for entry in sldLayoutIdLst.sldLayoutId_lst]
        assert len(ids) == len(set(ids))

    def it_does_not_reflect_later_edits_back_into_the_source(self):
        prs = Presentation()
        layouts = prs.slide_layouts
        source = layouts[6]
        placeholders_before = len(source.placeholders)
        clone = layouts.clone(source)

        clone.shapes.add_placeholder(PP_PLACEHOLDER.TITLE)

        assert len(clone.placeholders) == placeholders_before + 1
        assert len(source.placeholders) == placeholders_before

    def it_applies_an_explicit_name_verbatim(self):
        prs = Presentation()
        clone = prs.slide_layouts.clone(prs.slide_layouts[5], name="Explicit")
        assert clone.name == "Explicit"

    def it_supports_add_slide_with_the_clone(self):
        prs = Presentation()
        clone = prs.slide_layouts.clone(prs.slide_layouts[0])

        slide = prs.slides.add_slide(clone)
        slide.shapes.title.text = "On a clone"

        assert slide.slide_layout is clone
        reopened = _round_trip(prs)
        assert reopened.slides[0].shapes.title.text == "On a clone"

    def it_keeps_an_image_referenced_by_the_layout(self):
        prs = Presentation()
        source, image_part = _add_layout_with_picture(prs)
        layouts = prs.slide_layouts

        clone = layouts.clone(source)

        # ---the clone has its own picture shape with a remapped rId---
        clone_pic = next(s for s in clone.shapes if s.shape_type == 13)  # MSO_SHAPE_TYPE.PICTURE
        source_pic = next(s for s in source.shapes if s.shape_type == 13)
        clone_rId = clone_pic._element.blip_rId
        assert clone_rId != source_pic._element.blip_rId
        # ---and the remapped rId resolves to the (shared) image part---
        assert clone.part.related_part(clone_rId) is image_part

        # ---the saved file contains the image exactly once and the reopened
        # ---clone still resolves it (it renders in PowerPoint)---
        reopened = _round_trip(prs)
        reopened_clone = reopened.slide_layouts.get_by_name("%s 2" % source.name)
        assert reopened_clone is not None
        reopened_pic = next(s for s in reopened_clone.shapes if s.shape_type == 13)
        reopened_part = reopened_clone.part.related_part(reopened_pic._element.blip_rId)
        assert reopened_part.blob == SPEAKER_IMAGE_BYTES
        image_parts = [
            p
            for p in reopened.part.package.iter_parts()
            if getattr(p, "blob", b"") == SPEAKER_IMAGE_BYTES
        ]
        assert len(image_parts) == 1

    def but_it_raises_for_a_layout_from_another_presentation(self):
        prs, other = Presentation(), Presentation()

        with pytest.raises(ValueError):
            prs.slide_layouts.clone(other.slide_layouts[5])


# ---------------------------------------------------------------------------
# Round-trip stability
# ---------------------------------------------------------------------------


class Describe_round_trip:
    def it_round_trips_a_deck_with_added_and_cloned_layouts(self):
        prs = Presentation()
        layout = prs.slide_layouts.add_layout()
        layout.shapes.add_placeholder(PP_PLACEHOLDER.TITLE)
        layout.shapes.add_placeholder(PP_PLACEHOLDER.BODY)
        prs.slides.add_slide(layout)
        prs.slide_layouts.clone(prs.slide_layouts[5])
        _add_layout_with_picture(prs)
        prs.slide_layouts.clone(prs.slide_layouts[-1])

        assert round_trip_diff(prs) == {}
        assert_round_trip(prs)
