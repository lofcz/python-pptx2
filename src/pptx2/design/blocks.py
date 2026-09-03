"""Slide building blocks: the handful of pieces most content slides are made of.

Where :mod:`pptx2.design.recipes` builds whole slides and
:mod:`pptx2.design.components` builds token-driven dashboard widgets, this
module covers the everyday vocabulary of a teaching or explanatory deck,
driven by plain hex colours and points so a script needs no token setup:

* :func:`add_card` — one calm surface with a padded title and body. The
  card *is* the emphasis: a tinted fill or a hairline outline, generous
  padding, text that fits. No decorative stripes, badges or icons.
* :func:`add_bullets` — real PowerPoint bullets (``a:buChar`` /
  ``a:buAutoNum``) with a hanging indent and breathing room between
  items, shrunk to fit the box when the list runs long.
* :func:`add_picture_fit` — a picture placed *inside* a box, either
  letter-boxed (``mode="contain"``) or cropped to fill (``mode="cover"``),
  centred, with an optional caption underneath.

Every block tags the shapes it stacks with ``lint_group`` so the linter
treats a card and the text on it as one deliberate cluster, and every
block returns the shapes it made so callers can keep styling.
"""

from __future__ import annotations

import contextlib
import os
from dataclasses import dataclass
from typing import IO, TYPE_CHECKING, Any, Optional, Sequence, Union

from pptx2.enum.shapes import MSO_SHAPE
from pptx2.geometry import BBox
from pptx2.util import Emu, Pt

if TYPE_CHECKING:
    from pptx2.shapes.autoshape import Shape
    from pptx2.slide import Slide

__all__ = (
    "Card",
    "FittedPicture",
    "add_card",
    "add_bullets",
    "add_picture_fit",
)


def _as_bbox(bbox_or_positional: Sequence[Any]) -> BBox:
    if len(bbox_or_positional) == 1 and isinstance(bbox_or_positional[0], BBox):
        return bbox_or_positional[0]
    if len(bbox_or_positional) == 4:
        left, top, width, height = bbox_or_positional
        return BBox(Emu(int(left)), Emu(int(top)), Emu(int(width)), Emu(int(height)))
    raise TypeError(
        "pass either a BBox or (left, top, width, height); got %d positional arg(s)"
        % len(bbox_or_positional)
    )


def _tag(shapes: Sequence[Any], group: str) -> None:
    for shape in shapes:
        if shape is None:
            continue
        with contextlib.suppress(AttributeError, NotImplementedError):
            shape.lint_group = group


def _fit(tf: Any, *, font: Optional[str], max_pt: float, min_pt: float, bold: bool) -> None:
    """Shrink *tf* to fit its shape, never below *min_pt*; fall back to autofit."""
    from pptx2.enum.text import MSO_AUTO_SIZE

    try:
        applied = tf.fit_text(font_family=font, max_size=max(1, int(round(max_pt))), bold=bold)
    except (ValueError, OSError):
        applied = None
    if applied is not None and applied < min_pt:
        for para in tf.paragraphs:
            for run in para.runs:
                run.font.size = Pt(min_pt)
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    elif applied is None:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


# ----------------------------------------------------------------------------- bullets


def add_bullets(
    slide: "Slide",
    *bbox_or_positional,
    items: Sequence[str],
    size_pt: float = 18.0,
    color: Any = "#0F172A",
    font: Optional[str] = None,
    bold: bool = False,
    numbered: bool = False,
    bullet: str = "•",
    gap_pt: float = 8.0,
    line_spacing: float = 1.1,
    align: str = "left",
    anchor: str = "top",
    margin_pt: float = 0.0,
    min_size_pt: float = 12.0,
) -> "Shape":
    """Add a bulleted (or numbered) list that fits its box.

    ``items`` become one paragraph each, carrying a real PowerPoint bullet
    with a hanging indent — so wrapped lines align under the first word,
    not under the bullet — and ``gap_pt`` of space after every item.
    Text is measured and shrunk to fit the box, never below
    ``min_size_pt``; if even that overflows, PowerPoint's shrink-on-render
    is switched on as a safety net.

    Returns the textbox :class:`Shape`.
    """
    from pptx2._color import coerce_color
    from pptx2._textstyle import coerce_align, coerce_anchor

    bb = _as_bbox(bbox_or_positional)
    items = [str(item) for item in items]
    if not items:
        raise ValueError("items must be non-empty")

    box = slide.shapes.add_textbox(*bb)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = coerce_anchor(anchor)
    if margin_pt:
        m = Pt(margin_pt)
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = m
    rgb = coerce_color(color)
    align_value = coerce_align(align)

    for i, item in enumerate(items):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align_value
        para.line_spacing = float(line_spacing)
        if i < len(items) - 1:
            para.space_after = Pt(gap_pt)
        run = para.add_run()
        run.text = item
        if font is not None:
            run.font.name = font
        run.font.size = Pt(size_pt)
        run.font.bold = bool(bold)
        run.font.color.rgb = rgb
        # Hanging list: marker at the left edge, every wrapped line aligned
        # under the first word. The indent scales with the type size so a
        # two-digit number still clears the text.
        hang = Pt(size_pt * 1.4)
        if numbered:
            para.bullet.set_numbered(
                "arabicPeriod", start_at=1, left_margin=hang, hanging_indent=hang
            )
            if i > 0:
                # Only the first item pins the start number; the rest
                # continue the sequence (an explicit startAt on every
                # paragraph would restart each one at 1).
                para._p.pPr.buAutoNum.startAt = None
        else:
            para.bullet.set_character(bullet, left_margin=hang, hanging_indent=hang)

    _fit(tf, font=font, max_pt=size_pt, min_pt=min_size_pt, bold=bold)
    return box


# ----------------------------------------------------------------------------- card


@dataclass
class Card:
    """Shapes produced by :func:`add_card`."""

    card: Any
    title_box: Optional[Any]
    body_box: Optional[Any]
    inner: BBox
    """The padded content area, for placing extra shapes inside the card."""

    @property
    def shapes(self) -> list:
        return [s for s in (self.card, self.title_box, self.body_box) if s is not None]


def add_card(
    slide: "Slide",
    *bbox_or_positional,
    title: Optional[str] = None,
    body: Union[str, Sequence[str], None] = None,
    fill: Any = "#F1F5F9",
    line: Any = None,
    line_pt: float = 1.0,
    radius_pt: float = 12.0,
    pad_pt: float = 20.0,
    title_size_pt: float = 20.0,
    body_size_pt: float = 16.0,
    title_color: Any = "#0F172A",
    body_color: Any = "#334155",
    font: Optional[str] = None,
    align: str = "left",
    anchor: str = "top",
    title_gap_pt: float = 6.0,
    body_min_size_pt: float = 12.0,
    numbered: bool = False,
) -> Card:
    """Add a card: one surface, padded title and body, nothing else.

    The surface is a rounded rectangle with a flat ``fill`` (and no theme
    shadow). Give it *either* a tinted fill *or* a ``line`` outline — a
    tint reads as a surface on its own, an outline reads as a frame; both
    at once compete. Text sits inside ``pad_pt`` of padding on every side.

    ``body`` may be a string (one paragraph, wrapped) or a sequence of
    strings (rendered through :func:`add_bullets`, numbered when
    ``numbered=True``). Body text is fitted to the remaining height and
    never shrinks below ``body_min_size_pt``.

    Returns a :class:`Card` exposing ``card``, ``title_box``, ``body_box``
    and ``inner`` (the padded content box) so a picture, equation or
    diagram can be dropped inside the same card.
    """
    from pptx2._color import coerce_color
    from pptx2._textstyle import coerce_align, coerce_anchor

    bb = _as_bbox(bbox_or_positional)
    pad = Pt(pad_pt)
    inner = bb.inset(all=pad)

    surface = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, *bb)
    surface.fill.solid()
    surface.fill.fore_color.rgb = coerce_color(fill)
    if line is not None:
        surface.line.color.rgb = coerce_color(line)
        surface.line.width = Pt(line_pt)
    else:
        surface.line.fill.background()
    surface.shadow.clear()
    short_edge = min(int(bb.width), int(bb.height))
    surface.corner_radius = Emu(min(int(Pt(radius_pt)), short_edge // 2))
    surface.text_frame.text = ""

    title_box = None
    body_box = None
    cursor_top = int(inner.top)
    remaining = int(inner.height)

    if title:
        title_h = int(Pt(title_size_pt * 1.45))
        title_h = min(title_h, remaining)
        title_box = slide.shapes.add_textbox(inner.left, Emu(cursor_top), inner.width, Emu(title_h))
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
        tf.vertical_anchor = coerce_anchor("top")
        para = tf.paragraphs[0]
        para.alignment = coerce_align(align)
        run = para.add_run()
        run.text = title
        if font is not None:
            run.font.name = font
        run.font.size = Pt(title_size_pt)
        run.font.bold = True
        run.font.color.rgb = coerce_color(title_color)
        _fit(tf, font=font, max_pt=title_size_pt, min_pt=max(12.0, title_size_pt * 0.7), bold=True)
        cursor_top += title_h + int(Pt(title_gap_pt))
        remaining = int(inner.bottom) - cursor_top

    if body is not None and remaining > int(Pt(body_size_pt)):
        body_bb = BBox(inner.left, Emu(cursor_top), inner.width, Emu(remaining))
        if isinstance(body, str):
            body_box = slide.shapes.add_textbox(*body_bb)
            tf = body_box.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
            tf.vertical_anchor = coerce_anchor(anchor if title is None else "top")
            para = tf.paragraphs[0]
            para.alignment = coerce_align(align)
            para.line_spacing = 1.1
            run = para.add_run()
            run.text = body
            if font is not None:
                run.font.name = font
            run.font.size = Pt(body_size_pt)
            run.font.color.rgb = coerce_color(body_color)
            _fit(tf, font=font, max_pt=body_size_pt, min_pt=body_min_size_pt, bold=False)
        else:
            body_box = add_bullets(
                slide,
                body_bb,
                items=list(body),
                size_pt=body_size_pt,
                color=body_color,
                font=font,
                numbered=numbered,
                align=align,
                anchor=anchor if title is None else "top",
                min_size_pt=body_min_size_pt,
            )

    card = Card(card=surface, title_box=title_box, body_box=body_box, inner=inner)
    _tag(card.shapes, f"card@{int(bb.left)},{int(bb.top)}")
    return card


# ----------------------------------------------------------------------------- picture


@dataclass
class FittedPicture:
    """Shapes produced by :func:`add_picture_fit`."""

    picture: Any
    caption_box: Optional[Any]
    frame: BBox
    """The box the picture actually occupies (after contain/cover)."""


def _image_size(image: Union[str, "os.PathLike[str]", IO[bytes]]) -> tuple[int, int]:
    from PIL import Image as _PIL

    if hasattr(image, "read"):
        pos = image.tell()  # type: ignore[union-attr]
        with _PIL.open(image) as im:  # type: ignore[arg-type]
            size = im.size
        image.seek(pos)  # type: ignore[union-attr]
        return size
    with _PIL.open(image) as im:  # type: ignore[arg-type]
        return im.size


def add_picture_fit(
    slide: "Slide",
    image: Union[str, "os.PathLike[str]", IO[bytes]],
    *bbox_or_positional,
    mode: str = "contain",
    align: str = "center",
    caption: Optional[str] = None,
    caption_size_pt: float = 12.0,
    caption_color: Any = "#64748B",
    caption_gap_pt: float = 6.0,
    caption_height_pt: float = 22.0,
    font: Optional[str] = None,
) -> FittedPicture:
    """Place *image* inside a box without distorting it.

    ``mode="contain"`` scales the picture to fit entirely inside the box
    (letter-boxed, positioned by ``align``: ``"center"``, ``"left"``,
    ``"right"``, ``"top"``, ``"bottom"`` or corner pairs like
    ``"top-left"``). ``mode="cover"`` fills the whole box and crops the
    overflow symmetrically — the right choice for photographic
    backgrounds and edge-to-edge hero images.

    When ``caption`` is given, ``caption_height_pt`` is reserved at the
    bottom of the box and the picture is fitted above it.

    Returns a :class:`FittedPicture` with the picture, the caption box
    (or ``None``) and the box the picture finally occupies.
    """
    bb = _as_bbox(bbox_or_positional)
    if mode not in ("contain", "cover"):
        raise ValueError("mode must be 'contain' or 'cover', got %r" % (mode,))

    caption_box = None
    pic_area = bb
    if caption:
        cap_h = int(Pt(caption_height_pt))
        gap = int(Pt(caption_gap_pt))
        pic_area = BBox(bb.left, bb.top, bb.width, Emu(max(1, int(bb.height) - cap_h - gap)))

    iw, ih = _image_size(image)
    box_w, box_h = int(pic_area.width), int(pic_area.height)
    img_ratio = iw / float(ih)
    box_ratio = box_w / float(box_h)

    if mode == "contain":
        if img_ratio >= box_ratio:
            w = box_w
            h = int(round(box_w / img_ratio))
        else:
            h = box_h
            w = int(round(box_h * img_ratio))
        a = align.lower()
        if "left" in a:
            left = int(pic_area.left)
        elif "right" in a:
            left = int(pic_area.right) - w
        else:
            left = int(pic_area.left) + (box_w - w) // 2
        if "top" in a:
            top = int(pic_area.top)
        elif "bottom" in a:
            top = int(pic_area.bottom) - h
        else:
            top = int(pic_area.top) + (box_h - h) // 2
        frame = BBox(Emu(left), Emu(top), Emu(w), Emu(h))
        picture = slide.shapes.add_picture(image, frame.left, frame.top, frame.width, frame.height)
    else:
        frame = pic_area
        picture = slide.shapes.add_picture(image, frame.left, frame.top, frame.width, frame.height)
        if img_ratio > box_ratio:
            # Image is wider than the box: trim left/right.
            keep = box_ratio / img_ratio
            trim = (1.0 - keep) / 2.0
            picture.crop_left = trim
            picture.crop_right = trim
        elif img_ratio < box_ratio:
            keep = img_ratio / box_ratio
            trim = (1.0 - keep) / 2.0
            picture.crop_top = trim
            picture.crop_bottom = trim

    if caption:
        cap_top = int(frame.bottom) + int(Pt(caption_gap_pt))
        cap_h = int(bb.bottom) - cap_top
        if cap_h > 0:
            a = align.lower()
            caption_align = "left" if "left" in a else "right" if "right" in a else "center"
            caption_box = slide.shapes.add_text(
                BBox(bb.left, Emu(cap_top), bb.width, Emu(cap_h)),
                text=caption,
                font=font,
                size_pt=caption_size_pt,
                italic=True,
                color=caption_color,
                align=caption_align,
                anchor="top",
                margin_pt=0,
            )
            from pptx2.enum.text import MSO_AUTO_SIZE

            caption_box.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    _tag([picture, caption_box], f"picture@{int(bb.left)},{int(bb.top)}")
    return FittedPicture(picture=picture, caption_box=caption_box, frame=frame)
