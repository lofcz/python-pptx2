"""Slide-related objects, including masters, layouts, and notes."""

from __future__ import annotations

import copy
from contextlib import contextmanager
from dataclasses import dataclass
from typing import TYPE_CHECKING, Iterator, Sequence, cast

from pptx2.dml.fill import FillFormat
from pptx2.enum.presentation import (
    MSO_TRANSITION_TYPE,
    P14_TRANSITION_NAMES,
    P159_TRANSITION_NAMES,
)
from pptx2.enum.shapes import PP_PLACEHOLDER
from pptx2.errors import TargetNotFoundError, UnsupportedStructureError
from pptx2.opc.constants import RELATIONSHIP_TYPE as RT
from pptx2.oxml.ns import qn
from pptx2.shapes.shapetree import (
    LayoutPlaceholders,
    LayoutShapes,
    MasterPlaceholders,
    MasterShapes,
    NotesSlidePlaceholders,
    NotesSlideShapes,
    SlidePlaceholders,
    SlideShapes,
)
from pptx2.shared import ElementProxy, ParentedElementProxy, PartElementProxy
from pptx2.util import lazyproperty


def _relationship_references(root, rId: str) -> bool:
    """Return whether any relationship-qualified XML attribute contains `rId`."""
    return any(
        value == rId
        and name.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}")
        for element in root.iter()
        for name, value in element.attrib.items()
    )


def _require_slide_enrolled(slide: "Slide", *, argument: str = "slide") -> None:
    """Refuse a detached slide proxy before it can mutate unreachable XML."""
    package = slide.part.package
    presentation_part = package.presentation_part
    matches = []
    for sldId in presentation_part._element.sldIdLst.sldId_lst:
        try:
            rel = presentation_part.rels[sldId.rId]
            if not rel.is_external and rel.reltype == RT.SLIDE and rel.target_part is slide.part:
                matches.append(sldId)
        except (AssertionError, KeyError, ValueError):
            continue
    if len(matches) != 1 or slide._element is not slide.part._element:
        raise TargetNotFoundError("%s is stale or no longer enrolled" % argument)


def _require_layout_enrolled(layout: "SlideLayout", *, argument: str = "slide_layout") -> None:
    """Refuse a detached layout proxy before it can be selected as a target."""
    from pptx2.parts.slide import SlideMasterPart

    package = layout.part.package
    matches = []
    for part in package.iter_parts():
        if not isinstance(part, SlideMasterPart):
            continue
        id_list = part._element.sldLayoutIdLst
        if id_list is None:
            continue
        for entry in id_list.sldLayoutId_lst:
            try:
                rel = part.rels[entry.rId]
                if (
                    not rel.is_external
                    and rel.reltype == RT.SLIDE_LAYOUT
                    and rel.target_part is layout.part
                ):
                    matches.append((part, entry))
            except (AssertionError, KeyError, ValueError):
                continue
    if len(matches) != 1 or layout._element is not layout.part._element:
        raise TargetNotFoundError("%s is stale or no longer enrolled" % argument)


def _inbound_relationships(package, target_part):
    """Return every reachable internal relationship targeting `target_part`."""
    inbound = []
    owners = [package] + list(package.iter_parts())
    for owner in owners:
        relationships = package._rels if owner is package else owner.rels
        for rId, rel in relationships.items():
            if rel.is_external:
                continue
            try:
                if rel.target_part is target_part:
                    inbound.append((owner, rId, rel))
            except (AssertionError, ValueError):
                continue
    return tuple(inbound)


if TYPE_CHECKING:
    from datetime import datetime

    from pptx2.animation import SlideAnimations
    from pptx2.lint import SlideLintReport
    from pptx2.oxml.presentation import CT_SlideIdList, CT_SlideMasterIdList
    from pptx2.oxml.slide import (
        CT_CommonSlideData,
        CT_NotesSlide,
        CT_Slide,
        CT_SlideLayoutIdList,
        CT_SlideMaster,
    )
    from pptx2.parts.presentation import PresentationPart
    from pptx2.parts.slide import SlideLayoutPart, SlideMasterPart, SlidePart
    from pptx2.presentation import Presentation
    from pptx2.rebind import RebindReport
    from pptx2.shapes.placeholder import LayoutPlaceholder, MasterPlaceholder
    from pptx2.shapes.shapetree import NotesSlidePlaceholder
    from pptx2.smart_art import SmartArtCollection
    from pptx2.text.text import TextFrame


class _BaseSlide(PartElementProxy):
    """Base class for slide objects, including masters, layouts and notes."""

    _element: CT_Slide

    @lazyproperty
    def background(self) -> _Background:
        """|_Background| object providing slide background properties.

        This property returns a |_Background| object whether or not the
        slide, master, or layout has an explicitly defined background.

        The same |_Background| object is returned on every call for the same
        slide object.
        """
        return _Background(self._element.cSld)

    @property
    def name(self) -> str:
        """String representing the internal name of this slide.

        Returns an empty string (`''`) if no name is assigned. Assigning an empty string or |None|
        to this property causes any name to be removed.
        """
        return self._element.cSld.name

    @name.setter
    def name(self, value: str | None):
        new_value = "" if value is None else value
        self._element.cSld.name = new_value


class _BaseMaster(_BaseSlide):
    """Base class for master objects such as |SlideMaster| and |NotesMaster|.

    Provides access to placeholders and regular shapes.
    """

    @lazyproperty
    def placeholders(self) -> MasterPlaceholders:
        """|MasterPlaceholders| collection of placeholder shapes in this master.

        Sequence sorted in `idx` order.
        """
        return MasterPlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self):
        """
        Instance of |MasterShapes| containing sequence of shape objects
        appearing on this slide.
        """
        return MasterShapes(self._element.spTree, self)


class NotesMaster(_BaseMaster):
    """Proxy for the notes master XML document.

    Provides access to shapes, the most commonly used of which are placeholders.
    """


class NotesSlide(_BaseSlide):
    """Notes slide object.

    Provides access to slide notes placeholder and other shapes on the notes handout
    page.
    """

    element: CT_NotesSlide  # pyright: ignore[reportIncompatibleMethodOverride]

    def clone_master_placeholders(self, notes_master: NotesMaster) -> None:
        """Selectively add placeholder shape elements from `notes_master`.

        Selected placeholder shape elements from `notes_master` are added to the shapes
        collection of this notes slide. Z-order of placeholders is preserved. Certain
        placeholders (header, date, footer) are not cloned.
        """

        def iter_cloneable_placeholders() -> Iterator[MasterPlaceholder]:
            """Generate a reference to each cloneable placeholder in `notes_master`.

            These are the placeholders that should be cloned to a notes slide when the a new notes
            slide is created.
            """
            cloneable = (
                PP_PLACEHOLDER.SLIDE_IMAGE,
                PP_PLACEHOLDER.BODY,
                PP_PLACEHOLDER.SLIDE_NUMBER,
            )
            for placeholder in notes_master.placeholders:
                if placeholder.element.ph_type in cloneable:
                    yield placeholder

        shapes = self.shapes
        for placeholder in iter_cloneable_placeholders():
            shapes.clone_placeholder(cast("LayoutPlaceholder", placeholder))

    @property
    def notes_placeholder(self) -> NotesSlidePlaceholder | None:
        """the notes placeholder on this notes slide, the shape that contains the actual notes text.

        Return |None| if no notes placeholder is present; while this is probably uncommon, it can
        happen if the notes master does not have a body placeholder, or if the notes placeholder
        has been deleted from the notes slide.
        """
        for placeholder in self.placeholders:
            if placeholder.placeholder_format.type == PP_PLACEHOLDER.BODY:
                return placeholder
        return None

    @property
    def notes_text_frame(self) -> TextFrame | None:
        """The text frame of the notes placeholder on this notes slide.

        |None| if there is no notes placeholder. This is a shortcut to accommodate the common case
        of simply adding "notes" text to the notes "page".
        """
        notes_placeholder = self.notes_placeholder
        if notes_placeholder is None:
            return None
        return notes_placeholder.text_frame

    @lazyproperty
    def placeholders(self) -> NotesSlidePlaceholders:
        """Instance of |NotesSlidePlaceholders| for this notes-slide.

        Contains the sequence of placeholder shapes in this notes slide.
        """
        return NotesSlidePlaceholders(self.element.spTree, self)

    @lazyproperty
    def shapes(self) -> NotesSlideShapes:
        """Sequence of shape objects appearing on this notes slide."""
        return NotesSlideShapes(self._element.spTree, self)


class SlideTransition(object):
    """Provides access to the transition into a slide.

    A |SlideTransition| object is returned by :attr:`Slide.transition`
    whether or not an explicit ``<p:transition>`` element is present on the
    slide; reads on properties of an absent transition return |None| and
    never mutate the underlying XML, so theme inheritance is preserved.

    Setting any property creates the ``<p:transition>`` element on demand;
    use :meth:`clear` to remove the element entirely (restoring the default
    "no explicit transition" state).
    """

    def __init__(self, sld_elm):
        self._sld = sld_elm

    @property
    def kind(self) -> MSO_TRANSITION_TYPE | None:
        """Transition kind as :ref:`MsoTransitionType`, or |None| if not set."""
        transition = self._sld.transition
        if transition is None:
            return None
        kind_elm = transition.kind_element
        if kind_elm is None:
            # explicit `<p:transition/>` with no child means "cut" / no animation
            return MSO_TRANSITION_TYPE.NONE
        local = kind_elm.tag.rsplit("}", 1)[-1]
        try:
            return MSO_TRANSITION_TYPE.from_xml(local)
        except ValueError:
            return None

    @kind.setter
    def kind(self, value: MSO_TRANSITION_TYPE | None) -> None:
        if value is None:
            self.clear()
            return
        if not isinstance(value, MSO_TRANSITION_TYPE):
            raise TypeError("kind must be a MSO_TRANSITION_TYPE member or None, got %r" % (value,))
        transition = self._sld.get_or_add_transition()
        # remove any pre-existing kind child
        existing = transition.kind_element
        if existing is not None:
            transition.remove(existing)
        if value is MSO_TRANSITION_TYPE.NONE:
            return
        local = value.xml_value
        if local in P159_TRANSITION_NAMES:
            prefix = "p159"
        elif local in P14_TRANSITION_NAMES:
            prefix = "p14"
        else:
            prefix = "p"
        kind_elm = etree.Element(
            qn("%s:%s" % (prefix, local)),
            nsmap={prefix: _PREFIX_TO_URI[prefix]},
        )
        # insert at position 0 (before any sndAc/extLst)
        transition.insert(0, kind_elm)

    @property
    def duration(self) -> int | None:
        """Transition duration in milliseconds, or |None| if not explicitly set.

        Resolves the ``p14:dur`` attribute (PowerPoint 2010+ extension) if
        present; falls back to mapping the legacy ``spd`` bucket
        (``slow``/``med``/``fast`` ↔ 1000/750/500 ms) otherwise.
        """
        transition = self._sld.transition
        if transition is None:
            return None
        dur_attr = transition.get(qn("p14:dur"))
        if dur_attr is not None:
            try:
                return int(dur_attr)
            except ValueError:
                return None
        spd = transition.spd
        if spd is None:
            return None
        return _SPD_TO_MS.get(spd)

    @duration.setter
    def duration(self, ms: int | None) -> None:
        if ms is None:
            # clearing on a slide that inherits should be a no-op, not a
            # mutation that introduces an empty `<p:transition>` element
            transition = self._sld.transition
            if transition is None:
                return
            transition.attrib.pop(qn("p14:dur"), None)
            # also drop the legacy `spd` bucket; otherwise the getter falls
            # back to it and reads as still-explicitly-set
            transition.spd = None
            return
        if ms < 0:
            raise ValueError("duration must be a non-negative integer (milliseconds)")
        transition = self._sld.get_or_add_transition()
        transition.set(qn("p14:dur"), str(int(ms)))
        # writing an explicit ms duration supersedes any legacy bucket
        transition.spd = None

    @property
    def advance_on_click(self) -> bool | None:
        """Whether the slide advances on mouse-click; |None| if unset."""
        transition = self._sld.transition
        if transition is None:
            return None
        return transition.advClick

    @advance_on_click.setter
    def advance_on_click(self, value: bool | None) -> None:
        if value is None:
            transition = self._sld.transition
            if transition is None:
                return
            transition.advClick = None
            return
        transition = self._sld.get_or_add_transition()
        transition.advClick = bool(value)

    @property
    def advance_after(self) -> int | None:
        """Auto-advance time (milliseconds), or |None| if not auto-advancing."""
        transition = self._sld.transition
        if transition is None:
            return None
        return transition.advTm

    @advance_after.setter
    def advance_after(self, ms: int | None) -> None:
        if ms is None:
            transition = self._sld.transition
            if transition is None:
                return
            transition.advTm = None
            return
        if ms < 0:
            raise ValueError("advance_after must be a non-negative integer (milliseconds)")
        transition = self._sld.get_or_add_transition()
        transition.advTm = int(ms)

    def clear(self) -> None:
        """Remove the ``<p:transition>`` element entirely.

        After this call, the slide has no explicit transition; reads return
        |None| again. Idempotent: safe to call when no transition is set.
        """
        self._sld._remove_transition()


_SPD_TO_MS = {"slow": 1000, "med": 750, "fast": 500}


_PREFIX_TO_URI = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
}


# -- imported here to avoid a circular import at module load time --
from lxml import etree  # noqa: E402


class Slide(_BaseSlide):
    """Slide object. Provides access to shapes and slide-level properties."""

    part: SlidePart  # pyright: ignore[reportIncompatibleMethodOverride]

    @lazyproperty
    def animations(self) -> SlideAnimations:
        """Return a |SlideAnimations| object for adding animation effects to this slide.

        Animations are appended to the slide's timing tree in the order they
        are added.  Existing animations (e.g. authored in PowerPoint) are
        preserved and new effects are appended after them.

        Example::

            from pptx2.animation import Entrance, Trigger

            Entrance.fade(slide, shape)
            Entrance.fly_in(slide, shape2, direction="left",
                            trigger=Trigger.WITH_PREVIOUS)
        """
        from pptx2.animation import SlideAnimations

        return SlideAnimations(self)

    def render_thumbnail(self, **kwargs):
        """Render this slide to a PNG via headless LibreOffice.

        Thin wrapper around :func:`pptx2.render.render_slide_thumbnail`.
        Forwards `out_path`, `soffice_bin`, `timeout`, and `return_bytes`
        keyword arguments.  Requires ``soffice`` on PATH.
        """
        from pptx2.render import render_slide_thumbnail

        return render_slide_thumbnail(self, **kwargs)

    def lint_group(self, name: str | None, *shapes) -> None:
        """Tag every shape in *shapes* with ``lint_group = name``.

        Convenience batch form of ``shape.lint_group = name``. Shapes that
        share a non-empty ``lint_group`` are allowed to overlap without
        producing :class:`~pptx2.lint.ShapeCollision` warnings.

        Example::

            slide.lint_group("kpi-card-1", card, accent_bar, label_box, value_box)

        Passing ``name=None`` clears the tag on each supplied shape.
        """
        for shape in shapes:
            shape.lint_group = name

    def lint_group_overlaps(self, *shapes, name: str | None = None) -> str:
        """Tag *shapes* as a co-overlapping design group, returning the group name.

        Convenience over :meth:`lint_group` that also auto-generates a
        unique-on-the-slide group name when one isn't supplied, so
        callers don't have to invent ``"kpi-card-1"`` / ``"kpi-card-2"``
        labels by hand.  Inspired by IMPROVEMENT_PLAN.md item 12 — the
        single-line equivalent of:

            slide.lint_group(f"design-group-{n}", *shapes)

        Example::

            slide.lint_group_overlaps(card, accent_bar, label, value)

        When *name* is supplied it is used verbatim (matching
        :meth:`lint_group`).  Otherwise a name of the form
        ``"design-group-N"`` is chosen, where ``N`` starts at 1 and
        increments to the smallest positive integer that doesn't
        already appear as a ``lint_group`` tag on this slide.
        """
        if not shapes:
            raise ValueError("lint_group_overlaps requires at least one shape; got 0")
        if name is None:
            existing = {getattr(s, "lint_group", None) for s in self.shapes} - {None, ""}
            n = 1
            while f"design-group-{n}" in existing:
                n += 1
            name = f"design-group-{n}"
        for shape in shapes:
            shape.lint_group = name
        return name

    @contextmanager
    def design_group(self, name: str):
        """Context manager that auto-tags shapes added inside the block.

        Any shape appended to this slide's shape tree while the block is
        active receives ``lint_group = name`` (provided it doesn't already
        have a non-empty group, so nested ``design_group`` calls behave
        intuitively — the innermost label wins).

        Example::

            with slide.design_group("kpi-card-1"):
                slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, ...)   # card
                slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, ...)   # accent
                slide.shapes.add_textbox(...)                       # label
                slide.shapes.add_textbox(...)                       # value
            # All four now share ``lint_group = "kpi-card-1"``.
        """
        if not isinstance(name, str) or not name:
            raise ValueError("design_group name must be a non-empty string")

        sp_tree = self._element.spTree
        before = {id(elm) for elm in sp_tree.iter_shape_elms()}
        try:
            yield
        finally:
            from pptx2.lint import _read_lint_group, _write_lint_group

            for elm in sp_tree.iter_shape_elms():
                if id(elm) in before:
                    continue
                try:
                    cNvPr = elm._nvXxPr.cNvPr
                except AttributeError:
                    continue
                # Don't overwrite an explicit group set by the caller or by
                # an inner ``design_group`` block that already tagged this
                # shape.
                if _read_lint_group(cNvPr):
                    continue
                _write_lint_group(cNvPr, name)

    def lint(
        self,
        *,
        include_effect_bleed: bool = False,
        disable=(),
        min_severity="info",
    ) -> SlideLintReport:
        """Inspect this slide for geometric and typographic issues.

        Returns a |SlideLintReport| with a list of detected issues (text
        overflow, shapes off-slide, shape collisions).  The report is
        generated fresh on each call.

        *include_effect_bleed* (opt-in, default |False|) widens each
        shape's bbox by its shadow blur radius before the OffSlide and
        ShapeCollision checks run.  Bleed-only issues are emitted as
        ``OffSlideShadow`` / ``ShapeCollisionShadow`` so callers can
        suppress them via ``shape.lint_skip`` without losing real
        geometry warnings.

        *disable* is an iterable of issue ``code`` values to skip
        entirely — e.g. ``disable=["ShapeCollision"]``.

        *min_severity* drops issues below the named threshold from the
        report (``"info"`` / ``"warning"`` / ``"error"``).

        Example::

            report = slide.lint()
            if report.has_errors:
                print(report.summary())
        """
        from pptx2.lint import lint_slide

        return lint_slide(
            self,
            include_effect_bleed=include_effect_bleed,
            disable=disable,
            min_severity=min_severity,
        )

    def apply_footers(
        self,
        *,
        footer: str | None = None,
        slide_number: bool = False,
        date_format: str | None = None,
        fixed_date: str | None = None,
        now: "datetime | None" = None,
    ) -> None:
        """Apply the complete footer state to this slide only (the dialog's "Apply").

        paper-pptx addition. Same parameters, mechanism, and refusals as
        :meth:`.Presentation.apply_footers`, restricted to this slide — the per-slide
        override path (e.g. removing just this slide's footer while the rest of the deck
        keeps it). Each call sets this slide's full three-element state.
        """
        from pptx2.hf import apply_slide_footers

        apply_slide_footers(
            self,
            footer=footer,
            slide_number=slide_number,
            date_format=date_format,
            fixed_date=fixed_date,
            now=now,
        )

    def rebind_layout(
        self,
        target_layout: "SlideLayout",
        *,
        placeholder_map="auto",
        orphan_policy: str = "refuse",
    ) -> "RebindReport":
        """Move this slide to `target_layout`; return the required |RebindReport|.

        paper-pptx addition — the template-migration *primitive* (bulk-migration
        workflows are left to the caller). Placeholders reconcile against the
        target layout: auto-matching binds by exact type+idx, then same type, then
        interchangeable type family (title/ctrTitle; body/object/subTitle); pass
        `placeholder_map={source_idx: target_idx | None}` to override any of it (None
        force-orphans a source). Source placeholders with no destination follow
        `orphan_policy`: "refuse" (default; typed, atomic) or "bake" — convert to a free
        shape with inherited geometry materialized and each run's *resolved* effective
        formatting written locally, so the text keeps its look.

        The report is not optional: the effective-value resolver runs before and after,
        and every run whose resolved values changed appears with its before/after payloads
        — a rebind never shifts appearance silently. Same-package only (cross-package
        composition is `import_slide`'s job); slides carrying `mc:AlternateContent`
        refuse (shapes inside are invisible to reconciliation).
        """
        from pptx2.rebind import rebind_layout

        return rebind_layout(
            self,
            target_layout,
            placeholder_map=placeholder_map,
            orphan_policy=orphan_policy,
        )

    @property
    def follow_master_background(self):
        """|True| if this slide inherits the slide master background.

        Read-only.  Inheritance is broken as a side-effect of giving the
        slide its own background rather than by assigning here: touching
        :attr:`background` adds a ``<p:bg>`` element, after which this
        reports |False|.

        ::

            slide.follow_master_background          # True
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = "#102030"
            slide.follow_master_background          # now False

        To restore inheritance, remove the slide's own background.
        """
        return self._element.bg is None

    @property
    def has_notes_slide(self) -> bool:
        """`True` if this slide has a notes slide, `False` otherwise.

        A notes slide is created by :attr:`.notes_slide` when one doesn't exist; use this property
        to test for a notes slide without the possible side effect of creating one.
        """
        return self.part.has_notes_slide

    @property
    def notes_slide(self) -> NotesSlide:
        """The |NotesSlide| instance for this slide.

        If the slide does not have a notes slide, one is created. The same single instance is
        returned on each call.
        """
        return self.part.notes_slide

    @property
    def notes(self) -> str:
        """The speaker-notes text for this slide as a |str|.

        Returns an empty string when this slide has no notes slide or its notes
        placeholder is empty. This is a first-class, LLM-friendly shortcut over
        :attr:`.notes_slide` / :attr:`~.NotesSlide.notes_text_frame`; reading it
        never creates a notes slide.

        Assigning a string writes the notes text, creating the notes slide (and
        therefore its notes placeholder) on demand::

            slide.notes = "Remember to thank the sponsors."
        """
        if not self.has_notes_slide:
            return ""
        text_frame = self.notes_slide.notes_text_frame
        if text_frame is None:
            return ""
        return text_frame.text

    @notes.setter
    def notes(self, text: str) -> None:
        text_frame = self.notes_slide.notes_text_frame
        if text_frame is None:
            raise ValueError("notes slide has no notes placeholder; cannot set notes text")
        text_frame.text = text

    @lazyproperty
    def placeholders(self) -> SlidePlaceholders:
        """Sequence of placeholder shapes in this slide."""
        return SlidePlaceholders(self._element.spTree, self)

    def read_notes_text(self) -> str:
        """Return the text of this slide's existing speaker notes.

        paper-pptx addition. Unlike :attr:`notes_slide`, this NEVER creates a notes slide:
        a slide with no notes part raises |UnsupportedStructureError| (as does a notes slide
        with no body placeholder). Returns "" for an empty existing notes body.
        """
        return self._existing_notes_text_frame().text

    def replace_notes_text(self, text: str) -> None:
        """Replace the text of this slide's existing speaker notes with `text`.

        paper-pptx addition. Only the notes *body* placeholder is touched — slide-number and
        other notes placeholders are preserved untouched. The first paragraph's properties
        and its first run's character formatting are kept and applied to the replacement
        text; `"\\n"` in `text` starts a new paragraph. Never creates a notes slide: a slide
        with no notes part raises |UnsupportedStructureError| before anything changes
        (creating the notes part graph is intentionally not supported).
        """
        if not isinstance(text, str):
            raise ValueError("text must be a str, got %r" % type(text).__name__)
        try:
            text.encode("utf-8")  # -- lone surrogates would explode mid-mutation otherwise
        except UnicodeEncodeError:
            raise ValueError("text contains characters not encodable in XML: %r" % (text,))
        text_frame = self._existing_notes_text_frame()  # -- full validation before mutation

        txBody = text_frame._txBody
        paragraphs = txBody.p_lst
        first_p = paragraphs[0]
        first_r = first_p.find(qn("a:r"))
        rPr_template = None
        if first_r is not None:
            rPr = first_r.find(qn("a:rPr"))
            if rPr is not None:
                rPr_template = copy.deepcopy(rPr)

        # -- keep the first a:p element (preserving its a:pPr); drop the rest --
        for surplus_p in paragraphs[1:]:
            txBody.remove(surplus_p)
        for content in first_p.content_children:
            first_p.remove(content)
        pPr_template = first_p.find(qn("a:pPr"))

        lines = text.split("\n")
        for index, line in enumerate(lines):
            if index == 0:
                p = first_p
            else:
                p = txBody.add_p()
                if pPr_template is not None:
                    p.insert(0, copy.deepcopy(pPr_template))
            if line == "":
                continue  # -- an empty line is an empty paragraph
            r = p.add_r()
            if rPr_template is not None:
                r.insert(0, copy.deepcopy(rPr_template))
            r.text = line

    def _existing_notes_text_frame(self) -> TextFrame:
        """Return the body-placeholder text frame of this slide's EXISTING notes slide.

        Raises |UnsupportedStructureError| (never creates anything) when the slide has no
        notes part or its notes slide has no body placeholder.
        """
        if not self.has_notes_slide:
            raise UnsupportedStructureError(
                "slide %d has no notes slide; creating one is out of scope for this API"
                " (use notes_slide if you explicitly want creation)" % self.slide_id
            )
        notes_slide = self.part.part_related_by(RT.NOTES_SLIDE).notes_slide
        text_frame = notes_slide.notes_text_frame
        if text_frame is None:
            raise UnsupportedStructureError(
                "notes slide of slide %d has no body placeholder to hold notes text" % self.slide_id
            )
        return text_frame

    @lazyproperty
    def shapes(self) -> SlideShapes:
        """Sequence of shape objects appearing on this slide."""
        return SlideShapes(self._element.spTree, self)

    def slide_bbox(self):
        """Return the slide's full area as a :class:`~pptx2.geometry.BBox`."""
        from pptx2.geometry import BBox

        return BBox.from_slide(self)

    def content_bbox(self, *, include_decorative: bool = False):
        """Return the bounding box covering all non-decorative shapes.

        ``include_decorative=False`` (the default) skips slide-spanning
        backgrounds (any shape whose width and height each exceed 95%
        of the slide area), so the returned box reflects "where the
        real content is" rather than the full slide.

        Returns ``None`` when the slide has no qualifying shapes.
        """
        from pptx2.geometry import BBox

        slide_box = BBox.from_slide(self)
        union_box: BBox | None = None
        threshold_w = int(slide_box.width) * 0.95
        threshold_h = int(slide_box.height) * 0.95
        for shape in self.shapes:
            try:
                box = BBox.from_shape(shape)
            except Exception:
                continue
            if not include_decorative:
                if int(box.width) >= threshold_w and int(box.height) >= threshold_h:
                    continue
            union_box = box if union_box is None else union_box.union(box)
        return union_box

    def find_empty_region(
        self,
        *,
        near=None,
        min_width=0,
        min_height=0,
    ):
        """Return a :class:`BBox` of an unused region on the slide.

        Walks a coarse grid over the slide and returns the largest cell
        (or cluster of cells) that doesn't overlap any existing
        shape.  ``near`` is an optional BBox / Shape; when given, the
        cell whose centre is nearest its centre is preferred over
        strictly the largest free area.

        ``min_width`` / ``min_height`` filter out tiny free pockets in
        EMU.  Returns ``None`` when no region meets the criteria.

        Approximate by design — for one-off LLM placement decisions,
        not pixel-perfect packing.
        """
        from pptx2.geometry import BBox

        slide_box = BBox.from_slide(self)
        # 12×8 sample grid is fine-grained enough for typical decks.
        cells = slide_box.grid(12, 8)
        existing = []
        for shape in self.shapes:
            try:
                existing.append(BBox.from_shape(shape))
            except Exception:
                pass

        free = [c for c in cells if not any(c.intersects(s) for s in existing)]
        if not free:
            return None

        # Merge horizontally-adjacent free cells in the same row.
        merged: list[BBox] = []
        for c in free:
            if merged:
                last = merged[-1]
                if int(last.top) == int(c.top) and int(last.right) == int(c.left):
                    merged[-1] = last.union(c)
                    continue
            merged.append(c)

        candidates = [
            m for m in merged if int(m.width) >= int(min_width) and int(m.height) >= int(min_height)
        ]
        if not candidates:
            return None

        if near is not None:
            from pptx2.shapes.base import BaseShape

            if isinstance(near, BaseShape):
                target = BBox.from_shape(near)
            elif isinstance(near, BBox):
                target = near
            else:
                raise TypeError(
                    "find_empty_region(near=...) must be a BaseShape or BBox; "
                    f"got {type(near).__name__}"
                )
            tx, ty = int(target.cx), int(target.cy)
            candidates.sort(key=lambda b: (int(b.cx) - tx) ** 2 + (int(b.cy) - ty) ** 2)
            return candidates[0]
        candidates.sort(key=lambda b: -b.area)
        return candidates[0]

    def tidy(
        self,
        *,
        fix_offslide: bool = True,
        fix_overflow: bool = True,
        fix_grid_drift: bool = False,
        fix_layer_order: bool = True,
    ) -> list[str]:
        """One-call cleanup: lint then auto-fix the safe subset.

        Wraps :meth:`lint` + :meth:`SlideLintReport.auto_fix` with the
        flags most decks want by default.  Returns the list of fixes
        applied (the same shape as ``auto_fix()``).

        * ``fix_offslide`` (default ``True``) clamps shapes back
          on-slide.
        * ``fix_overflow`` (default ``True``) flips overflowing text
          frames to ``MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE``.
        * ``fix_grid_drift`` (default ``False``) snaps minor grid drift
          — off by default because the snap can move a shape by several
          EMU when the inferred grid is wrong.
        * ``fix_layer_order`` (default ``True``) restacks shapes whose
          ``layer_above`` declaration is contradicted by the drawing
          order.  On by default because it only enforces an ordering
          the author already declared, and it never moves geometry.
        """
        disable: list[str] = []
        if not fix_offslide:
            disable.append("OffSlide")
        if not fix_overflow:
            disable.append("TextOverflow")
        if not fix_grid_drift:
            disable.append("OffGridDrift")
        if not fix_layer_order:
            disable.append("LayerOrderViolation")
        report = self.lint(disable=disable)
        return report.auto_fix()

    @property
    def smart_art(self) -> SmartArtCollection:
        """Return a |SmartArtCollection| for this slide.

        Provides indexed access to SmartArt graphics on the slide.  Each item
        is a |SmartArtShape| whose :attr:`~SmartArtShape.texts` property gives
        the current text list and :meth:`~SmartArtShape.set_text` replaces it.

        Example::

            org_chart = slide.smart_art[0]
            print(org_chart.texts)                # ['CEO', 'CTO', 'CFO']
            org_chart.set_text(['Alice', 'Bob', 'Carol'])

        Returns an empty collection (length 0) when there are no SmartArt
        shapes on the slide.
        """
        from pptx2.smart_art import SmartArtCollection

        return SmartArtCollection(self)

    @property
    def slide_id(self) -> int:
        """Integer value that uniquely identifies this slide within this presentation.

        The slide id does not change if the position of this slide in the slide sequence is changed
        by adding, rearranging, or deleting slides.
        """
        return self.part.slide_id

    @property
    def slide_layout(self) -> SlideLayout:
        """|SlideLayout| object this slide inherits appearance from."""
        return self.part.slide_layout

    @property
    def color_variant(self) -> str | None:
        """Per-slide color-mapping variant: ``"light"`` / ``"dark"`` / |None|.

        Reads / writes the ``<p:clrMapOvr>`` element to apply a built-in
        light or dark variant of the deck's master color map. ``"light"``
        is the master's default mapping (``bg1=lt1``, ``tx1=dk1``, …);
        ``"dark"`` swaps backgrounds and text (``bg1=dk1``, ``tx1=lt1``,
        …) for a dark-on-light slide without changing the deck theme.

        Reading returns:

        * ``"dark"``  — slide has an explicit override that swaps bg/tx.
        * ``"light"`` — slide inherits from the master (or has an
          explicit ``<a:masterClrMapping/>`` element).
        * ``None``    — slide has a custom override that doesn't match
          either of the two named variants.

        Assigning ``None`` removes the override entirely (returning to
        master inheritance).

        For more flexible control, use :meth:`set_clr_map_override`.
        """
        clr_ovr = self._element.clrMapOvr
        if clr_ovr is None:
            return "light"
        # If <a:masterClrMapping/> is present we're inheriting.
        for child in clr_ovr:
            local = child.tag.rsplit("}", 1)[-1]
            if local == "masterClrMapping":
                return "light"
            if local == "overrideClrMapping":
                if (
                    child.get("bg1") == "dk1"
                    and child.get("tx1") == "lt1"
                    and child.get("bg2") == "dk2"
                    and child.get("tx2") == "lt2"
                ):
                    return "dark"
                return None
        return None

    @color_variant.setter
    def color_variant(self, value: str | None) -> None:
        if value is None:
            self._element._remove_clrMapOvr()
            return
        if value == "light":
            self.set_clr_map_override(masterClrMapping=True)
            return
        if value == "dark":
            self.set_clr_map_override(
                bg1="dk1",
                tx1="lt1",
                bg2="dk2",
                tx2="lt2",
                accent1="accent1",
                accent2="accent2",
                accent3="accent3",
                accent4="accent4",
                accent5="accent5",
                accent6="accent6",
                hlink="hlink",
                folHlink="folHlink",
            )
            return
        raise ValueError(f"color_variant must be 'light', 'dark', or None; got {value!r}")

    def set_clr_map_override(self, *, masterClrMapping: bool = False, **mapping: str) -> None:
        """Set this slide's ``<p:clrMapOvr>`` element directly.

        With ``masterClrMapping=True`` (and no other args), removes any
        existing override and writes ``<a:masterClrMapping/>`` so the
        slide inherits the master's color map.

        Otherwise, writes an ``<a:overrideClrMapping>`` with the supplied
        attributes.  Standard mapping attributes are ``bg1``, ``tx1``,
        ``bg2``, ``tx2``, ``accent1``..``accent6``, ``hlink``,
        ``folHlink``; each value is the slot it should resolve to
        (e.g. ``bg1="dk1"`` redirects "background 1" lookups to the
        ``dk1`` palette slot).

        Use :attr:`color_variant` for the common light/dark presets.
        """
        clr_ovr = self._element.get_or_add_clrMapOvr()
        # Clear existing children.
        for child in list(clr_ovr):
            clr_ovr.remove(child)

        a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        if masterClrMapping and not mapping:
            child = etree.SubElement(clr_ovr, "{%s}masterClrMapping" % a_ns)
            return
        child = etree.SubElement(clr_ovr, "{%s}overrideClrMapping" % a_ns)
        for k, v in mapping.items():
            child.set(k, v)

    @lazyproperty
    def transition(self) -> SlideTransition:
        """|SlideTransition| object describing the transition into this slide.

        The same instance is returned on each call. Reads on individual
        properties of the returned object are non-mutating; the underlying
        ``<p:transition>`` element is created only when a property is
        assigned.
        """
        return SlideTransition(self._element)


@dataclass(frozen=True)
class SlideClonePolicy:
    """Relationship policy for `Slides.clone` (paper-pptx addition).

    Defaults encode the production-proven policy: charts (with their embedded workbooks and
    style parts) and speaker notes are deep-copied so clone and original can never
    cross-contaminate; image/media parts are shared deliberately.

    - `deep_copy_charts`: must be True to clone a slide bearing charts; False refuses
      (`RelationshipPolicyError`) rather than share an editable chart part between slides.
    - `deep_copy_notes`: False drops the notes slide from the clone (original unaffected).
    - `share_media`: False deep-copies image/media parts instead of sharing them.
    """

    deep_copy_charts: bool = True
    deep_copy_notes: bool = True
    share_media: bool = True


class Slides(ParentedElementProxy):
    """Sequence of slides belonging to an instance of |Presentation|.

    Has list semantics for access to individual slides. Supports indexed access, len(), and
    iteration.
    """

    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldIdLst: CT_SlideIdList, prs: Presentation):
        super(Slides, self).__init__(sldIdLst, prs)
        self._sldIdLst = sldIdLst

    def __getitem__(self, idx: int) -> Slide:
        """Provide indexed access, (e.g. 'slides[0]')."""
        try:
            sldId = self._sldIdLst.sldId_lst[idx]
        except IndexError:
            raise IndexError("slide index out of range")
        return self.part.related_slide(sldId.rId)

    def __iter__(self) -> Iterator[Slide]:
        """Support iteration, e.g. `for slide in slides:`."""
        for sldId in self._sldIdLst.sldId_lst:
            yield self.part.related_slide(sldId.rId)

    def __len__(self) -> int:
        """Support len() built-in function, e.g. `len(slides) == 4`."""
        return len(self._sldIdLst)

    def add_slide(self, slide_layout: SlideLayout) -> Slide:
        """Return a newly added slide that inherits layout from `slide_layout`."""
        rId, slide = self.part.add_slide(slide_layout)
        slide.shapes.clone_layout_placeholders(slide_layout)
        sldId = self._sldIdLst.add_sldId(rId)
        self._add_to_final_section(sldId.id)
        return slide

    def _add_to_final_section(self, slide_id: int) -> None:
        """Register a newly appended slide with the deck's final section.

        PowerPoint keeps the 2010 section extension a complete partition of
        the deck — every slide belongs to exactly one section.  A slide
        appended at the end of a sectioned deck therefore falls into the
        last section, so mirror that here to keep the `p14:sectionLst` in
        sync.  A no-op when the deck has no sections.
        """
        prs_elm = self._sldIdLst.getparent()
        sectionLst = getattr(prs_elm, "sectionLst", None)
        if sectionLst is None or not sectionLst.section_lst:
            return
        sectionLst.section_lst[-1].add_sldId(slide_id)

    def _remove_from_sections(self, slide_id: int) -> None:
        """Purge `slide_id` from every section's membership list.

        A `p14:section` reference to a slide that no longer exists is a
        dangling pointer PowerPoint treats as a repair trigger, so each
        section claiming the slide gives it up.  Sections left empty by
        the removal are kept — an empty section is schema-valid and
        matches what `Sections.add()` can already produce.  A no-op when
        the deck has no sections.
        """
        prs_elm = self._sldIdLst.getparent()
        sectionLst = getattr(prs_elm, "sectionLst", None)
        if sectionLst is None:
            return
        for section_elm in sectionLst.section_lst:
            section_elm.remove_sldId(slide_id)

    def clone(
        self,
        source: Slide | int,
        *,
        after: Slide | int | None = None,
        policy: SlideClonePolicy | None = None,
    ) -> Slide:
        """Return a new slide that is a policy-governed deep copy of `source`.

        paper-pptx addition. The clone's relationship graph follows `policy` (default
        |SlideClonePolicy|): layout shared; charts deep-copied WITH their embedded workbooks
        and style parts; notes deep-copied and re-linked to the clone; image/media shared;
        external (hyperlink) relationships copied. A slide bearing any other relationship
        type (OLE objects, controls, SmartArt, comments, …) refuses with
        |RelationshipPolicyError| before anything changes.

        The clone is inserted directly after `source`, or after the slide given by `after`.
        `source`/`after` accept a |Slide| or a 0-based index; a |Slide| from another
        presentation raises |TargetNotFoundError|.
        """
        from pptx2._transaction import PackageTransaction
        from pptx2.slideops import clone_slide_part, enroll_clone_in_section

        if policy is None:
            policy = SlideClonePolicy()
        if not isinstance(policy, SlideClonePolicy):
            raise ValueError("policy must be a SlideClonePolicy, got %r" % (policy,))
        source_slide = self._resolve_slide(source)
        anchor_slide = source_slide if after is None else self._resolve_slide(after)
        anchor_index = self.index(anchor_slide)

        source_slide_id = source_slide.slide_id
        with PackageTransaction(self.part.package, self, source_slide, anchor_slide):
            new_part = clone_slide_part(source_slide.part, policy)
            rId = self.part.relate_to(new_part, RT.SLIDE)
            self._sldIdLst.add_sldId(rId)
            sldId = self._sldIdLst[-1]
            self._sldIdLst.remove(sldId)
            self._sldIdLst.insert(anchor_index + 1, sldId)
            # -- enroll the copy in the source's section, right after it (custom shows are
            # -- deliberately not extended: a copy is not part of a curated show)
            enroll_clone_in_section(self._sldIdLst.getparent(), source_slide_id, sldId.id)
            cloned_slide = new_part.slide
        return cloned_slide

    def delete(self, slide: Slide | int) -> None:
        """Remove `slide` from this presentation.

        paper-pptx addition. Removes the slide's `p:sldId` entry and the presentation's
        relationship to the slide part; parts then unreachable through the relationship
        graph (the slide, and e.g. its charts and notes if unshared) are never serialized
        again — orphans structurally cannot reach disk. Deleting the last slide is allowed.
        """
        from pptx2._transaction import PackageTransaction
        from pptx2.slideops import remove_slide_from_id_lists

        target = self._resolve_slide(slide)
        for sldId in self._sldIdLst.sldId_lst:
            if sldId.id == target.slide_id:
                slide_id, rId = sldId.id, sldId.rId
                notes_owners = {
                    rel.target_part
                    for rel in target.part.rels.values()
                    if not rel.is_external and rel.reltype == RT.NOTES_SLIDE
                }
                for notes_part in notes_owners:
                    shared_notes = [
                        (owner, notes_rId)
                        for owner, notes_rId, _ in _inbound_relationships(
                            self.part.package, notes_part
                        )
                        if owner is not target.part
                    ]
                    if shared_notes:
                        raise UnsupportedStructureError(
                            "slide deletion refused: its notes part is shared by another "
                            "reachable package part"
                        )
                aliases = [
                    (owner, alias_rId)
                    for owner, alias_rId, _ in _inbound_relationships(
                        self.part.package, target.part
                    )
                    if not ((owner is self.part and alias_rId == rId) or owner in notes_owners)
                ]
                if aliases:
                    raise UnsupportedStructureError(
                        "slide deletion refused: slide part has additional inbound "
                        "relationship aliases"
                    )
                with PackageTransaction(self.part.package, self, target):
                    self._sldIdLst.remove(sldId)
                    # -- sections (by slide id) and custom shows (by rId) reference slides
                    # -- outside the rels graph; purge those entries too
                    remove_slide_from_id_lists(self._sldIdLst.getparent(), slide_id, rId)
                    if _relationship_references(self.part._element, rId):
                        raise UnsupportedStructureError(
                            "slide deletion refused: relationship %s remains referenced" % rId
                        )
                    self.part.drop_rel(rId)
                return

    def move(self, slide: "Slide | int", to_index: int) -> None:
        """Relocate `slide` to `to_index`, shifting the rest (paper-pptx hardened).

        `slide` may be a |Slide| object belonging to this collection or its
        zero-based index into it (negative indices count from the end, Python
        style). `to_index` is a zero-based position, likewise accepting
        negative offsets. PowerPoint slide order follows the order of `p:sldId`
        children in `p:sldIdLst`, so this detaches the corresponding element
        and re-inserts it at the new position::

            prs.slides.move(0, 2)       # send the first slide to position 2
            prs.slides.move(slide, 0)   # send a Slide object to the front

        Raises |IndexError| if either index is out of range and |ValueError|
        when `slide` is neither a |Slide| nor an int.
        """
        from pptx2._transaction import PackageTransaction

        if isinstance(slide, Slide):
            _require_slide_enrolled(slide)
            current_index = self.index(slide)
            target = slide
        else:
            if isinstance(slide, bool) or not isinstance(slide, int):
                raise ValueError("expected a Slide or int index, got %r" % (slide,))
            target = self[slide]  # -- IndexError on out of range; negatives allowed
            _require_slide_enrolled(target)
            current_index = self._normalized_index(slide, len(self._sldIdLst))
        if isinstance(to_index, bool) or not isinstance(to_index, int):
            raise ValueError("to_index must be an int, got %r" % (to_index,))
        to_index = self._normalized_index(to_index, len(self._sldIdLst))

        sldId = self._sldIdLst.sldId_lst[current_index]
        with PackageTransaction(self.part.package, self, target):
            self._sldIdLst.remove(sldId)
            # -- insert resolves against the post-removal order (lxml insert
            # -- tolerates an index one past the end, appending instead) --
            self._sldIdLst.insert(to_index, sldId)

    def reorder(self, new_order: "Sequence[int | Slide]") -> None:
        """Rearrange the slides into the permutation given by `new_order`.

        `new_order` is a full permutation of the collection, expressed either as
        zero-based indices into the *current* order or as the |Slide| objects
        themselves (the two forms may not be mixed). After the call, the slide
        that was at ``new_order[0]`` becomes the first slide, and so on::

            prs.slides.reorder([2, 0, 1])           # by index
            prs.slides.reorder([s2, s0, s1])        # by Slide object

        Raises |ValueError| if `new_order` is not a permutation of exactly the
        slides in this collection (wrong length, duplicates, or unknown items).
        """
        from pptx2._transaction import PackageTransaction

        count = len(self._sldIdLst)
        order = list(new_order)
        if len(order) != count:
            raise ValueError(
                "new_order must contain exactly %d items, got %d" % (count, len(order))
            )

        sldId_lst = self._sldIdLst.sldId_lst
        indices: list[int] = []
        for item in order:
            if isinstance(item, Slide):
                indices.append(self.index(item))
            elif isinstance(item, bool) or not isinstance(item, int):
                raise ValueError("new_order items must be ints or Slide objects, got %r" % (item,))
            elif not 0 <= item < count:
                raise ValueError("new_order must be a permutation of the slides in this collection")
            else:
                indices.append(item)

        if sorted(indices) != list(range(count)):
            raise ValueError("new_order must be a permutation of the slides in this collection")

        ordered_sldIds = [sldId_lst[i] for i in indices]
        with PackageTransaction(self.part.package, self):
            for sldId in ordered_sldIds:
                self._sldIdLst.remove(sldId)
            for sldId in ordered_sldIds:
                self._sldIdLst.append(sldId)

    def _resolve_slide(self, value: Slide | int) -> Slide:
        """Return the |Slide| in this collection for `value` (a Slide or 0-based index).

        An int resolves with normal indexed-access semantics (|IndexError| when out of
        range); a |Slide| not belonging to this presentation raises |TargetNotFoundError|.
        """
        if isinstance(value, int) and not isinstance(value, bool):
            slide = self[value]
            _require_slide_enrolled(slide)
            return slide
        if isinstance(value, Slide):
            for slide in self:
                if slide == value:
                    _require_slide_enrolled(value)
                    return slide
            raise TargetNotFoundError(
                "slide with id %d is not in this presentation's slide collection" % value.slide_id
            )
        raise ValueError("expected a Slide or int index, got %r" % (value,))

    def get(self, slide_id: int, default: Slide | None = None) -> Slide | None:
        """Return the slide identified by int `slide_id` in this presentation.

        Returns `default` if not found.
        """
        slide = self.part.get_slide(slide_id)
        if slide is None:
            return default
        return slide

    def index(self, slide: Slide) -> int:
        """Map `slide` to its zero-based position in this slide sequence.

        Raises |ValueError| on *slide* not present.
        """
        for idx, this_slide in enumerate(self):
            if this_slide == slide:
                return idx
        raise ValueError("%s is not in slide collection" % slide)

    def remove(self, slide: "Slide | int") -> None:
        """Remove `slide` from this collection, deleting it from the presentation.

        `slide` may be a |Slide| object belonging to this collection or the
        integer slide id of one (the form upstream python-pptx's
        ``remove_slide()`` accepts). The slide's `p:sldId` entry is dropped
        from `p:sldIdLst` and the relationship from the presentation part
        to the slide part is dropped. Because package parts are serialized
        by relationship reachability, that alone removes the slide part —
        and anything only it refers to, such as its notes slide — from the
        package on save. Any section membership referencing the slide is
        purged, leaving no dangling `p14:section` reference. Removing the
        last remaining slide leaves a valid, empty presentation.

        Raises |ValueError| when `slide` is not present in this collection.
        """
        if isinstance(slide, Slide):
            # Match on collection membership rather than slide id: ids are
            # only unique within a deck, so a foreign slide's id could
            # otherwise silently collide with this deck's own.
            sldId = self._sldIdLst.sldId_lst[self.index(slide)]
        else:
            slide_id = int(slide)
            sldId = None
            for candidate in self._sldIdLst.sldId_lst:
                if candidate.id == slide_id:
                    sldId = candidate
                    break
            if sldId is None:
                raise ValueError("no slide with id %d in this presentation" % slide_id)
        # -- detach the p:sldId element *before* dropping the relationship;
        # -- XmlPart.drop_rel only drops rels whose r:id is no longer
        # -- referenced in this part's XML --
        slide_id, rId = sldId.id, sldId.rId
        self._sldIdLst.remove(sldId)
        self.part.drop_rel(rId)
        self._remove_from_sections(slide_id)

    @staticmethod
    def _normalized_index(idx: int, count: int) -> int:
        """Return `idx` resolved against `count`, supporting negative indexing.

        Raises |IndexError| when out of range.
        """
        original = idx
        if idx < 0:
            idx += count
        if idx < 0 or idx >= count:
            raise IndexError("slide index out of range: %r" % original)
        return idx


class HeaderFooters(object):
    """Header/footer placeholder visibility flags of a layout or master (paper-pptx addition).

    Wraps the `p:hf` element. Each property is tri-state: |True|/|False| when the attribute
    is explicit, |None| when it is absent — meaning "inherit" (a layout inherits from its
    master; the schema default is visible). Assigning |None| removes the attribute.
    """

    def __init__(self, owner):
        """Bind to the layout or master that owns these flags."""
        super(HeaderFooters, self).__init__()
        self._owner = owner
        self._element = owner._element

    @property
    def slide_number_visible(self) -> bool | None:
        """Visibility of the slide-number placeholder (`p:hf/@sldNum`)."""
        hf = self._element.hf
        return hf.sldNum if hf is not None else None

    @slide_number_visible.setter
    def slide_number_visible(self, value: bool | None):
        self._set_flag("sldNum", value)

    @property
    def footer_visible(self) -> bool | None:
        """Visibility of the footer placeholder (`p:hf/@ftr`)."""
        hf = self._element.hf
        return hf.ftr if hf is not None else None

    @footer_visible.setter
    def footer_visible(self, value: bool | None):
        self._set_flag("ftr", value)

    @property
    def date_visible(self) -> bool | None:
        """Visibility of the date placeholder (`p:hf/@dt`)."""
        hf = self._element.hf
        return hf.dt if hf is not None else None

    @date_visible.setter
    def date_visible(self, value: bool | None):
        self._set_flag("dt", value)

    def _set_flag(self, attr_name: str, value: "bool | None") -> None:
        """Set one `p:hf` visibility flag inside a transaction.

        Refuses a detached element, and any value that is not True, False, or None.
        """
        from pptx2._ownership import require_element_attached
        from pptx2._transaction import PackageTransaction

        require_element_attached(self._element, self._owner.part, argument="header/footer flags")
        if not isinstance(value, bool) and value is not None:
            raise ValueError("visibility must be True, False, or None, got %r" % (value,))
        with PackageTransaction(self._owner.part.package, self, self._owner):
            if value is None:
                hf = self._element.hf
                if hf is not None:
                    setattr(hf, attr_name, None)
                return
            setattr(self._element.get_or_add_hf(), attr_name, value)


class SlideLayout(_BaseSlide):
    """Slide layout object.

    Provides access to placeholders, regular shapes, and slide layout-level properties.
    """

    part: SlideLayoutPart  # pyright: ignore[reportIncompatibleMethodOverride]

    @property
    def header_footers(self) -> HeaderFooters:
        """|HeaderFooters| flags for this layout (paper-pptx addition)."""
        return HeaderFooters(self)

    def iter_cloneable_placeholders(self) -> Iterator[LayoutPlaceholder]:
        """Generate layout-placeholders on this slide-layout that should be cloned to a new slide.

        Used when creating a new slide from this slide-layout.
        """
        latent_ph_types = (
            PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.FOOTER,
            PP_PLACEHOLDER.SLIDE_NUMBER,
        )
        for ph in self.placeholders:
            if ph.element.ph_type not in latent_ph_types:
                yield ph

    @lazyproperty
    def placeholders(self) -> LayoutPlaceholders:
        """Sequence of placeholder shapes in this slide layout.

        Placeholders appear in `idx` order.
        """
        return LayoutPlaceholders(self._element.spTree, self)

    @lazyproperty
    def shapes(self) -> LayoutShapes:
        """Sequence of shapes appearing on this slide layout."""
        return LayoutShapes(self._element.spTree, self)

    @property
    def slide_master(self) -> SlideMaster:
        """Slide master from which this slide-layout inherits properties."""
        return self.part.slide_master

    @property
    def used_by_slides(self):
        """Tuple of slide objects based on this slide layout."""
        # ---getting Slides collection requires going around the horn a bit---
        slides = self.part.package.presentation_part.presentation.slides
        return tuple(s for s in slides if s.slide_layout == self)


class SlideLayouts(ParentedElementProxy):
    """Sequence of slide layouts belonging to a slide-master.

    Supports indexed access, len(), iteration, index() and remove().
    """

    part: SlideMasterPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldLayoutIdLst: CT_SlideLayoutIdList, parent: SlideMaster):
        super(SlideLayouts, self).__init__(sldLayoutIdLst, parent)
        self._sldLayoutIdLst = sldLayoutIdLst

    def __getitem__(self, idx: int) -> SlideLayout:
        """Provides indexed access, e.g. `slide_layouts[2]`."""
        try:
            sldLayoutId = self._sldLayoutIdLst.sldLayoutId_lst[idx]
        except IndexError:
            raise IndexError("slide layout index out of range")
        return self.part.related_slide_layout(sldLayoutId.rId)

    def __iter__(self) -> Iterator[SlideLayout]:
        """Generate each |SlideLayout| in the collection, in sequence."""
        for sldLayoutId in self._sldLayoutIdLst.sldLayoutId_lst:
            yield self.part.related_slide_layout(sldLayoutId.rId)

    def __len__(self) -> int:
        """Support len() built-in function, e.g. `len(slides) == 4`."""
        return len(self._sldLayoutIdLst)

    def add_layout(self, name: str | None = "Layout %s") -> SlideLayout:
        """Return a newly added blank slide layout.

        The new layout appears at the end of this collection and inherits from this layouts'
        slide master. It contains no shapes; add placeholders to it through
        `layout.shapes.add_placeholder(...)`.

        `name` is assigned to the new layout when provided; a `%s` in `name` is substituted
        with the numeric portion of the new master-to-layout relationship id, so the default
        `"Layout %s"` might yield `"Layout 13"`. Pass |None| to leave the layout unnamed.
        """
        rId, layout = self.part.add_layout()
        self._sldLayoutIdLst.add_sldLayoutId(rId, self.part.next_layout_id)
        id_ = int(rId[3:]) if rId.startswith("rId") and rId[3:].isdigit() else 0
        if name:
            layout.name = name % id_ if "%s" in name else name
        return layout

    def clone(self, slide_layout: SlideLayout, name: str | None = None) -> SlideLayout:
        """Return a clone of `slide_layout` appended to this collection.

        The source layout's XML is deep-copied, so the clone has the same shapes and
        placeholders and renders identically, while later edits to either layout leave the
        other untouched. Dependent parts the layout refers to (images, media, hyperlinks) are
        shared with the source through fresh relationships.

        `name` defaults to a non-colliding variant of the source layout's name ("Title Only"
        becomes "Title Only 2", then "Title Only 3", and so on).

        Raises `ValueError` when `slide_layout` belongs to a different presentation; use
        `Presentation.import_slide()` to bring slides (and their layouts) across packages.
        """
        if slide_layout.part.package is not self.part.package:
            raise ValueError("slide_layout must belong to the same presentation")
        rId, layout = self.part.clone_layout(slide_layout)
        self._sldLayoutIdLst.add_sldLayoutId(rId, self.part.next_layout_id)
        layout.name = self._non_colliding_name(slide_layout.name if name is None else name)
        return layout

    def get_by_name(self, name: str, default: SlideLayout | None = None) -> SlideLayout | None:
        """Return SlideLayout object having `name`, or `default` if not found."""
        for slide_layout in self:
            if slide_layout.name == name:
                return slide_layout
        return default

    def index(self, slide_layout: SlideLayout) -> int:
        """Return zero-based index of `slide_layout` in this collection.

        Raises `ValueError` if `slide_layout` is not present in this collection.
        """
        for idx, this_layout in enumerate(self):
            if slide_layout == this_layout:
                return idx
        raise ValueError("layout not in this SlideLayouts collection")

    def remove(self, slide_layout: SlideLayout) -> None:
        """Remove `slide_layout` from the collection.

        Raises ValueError when `slide_layout` is in use; a slide layout which is the basis for one
        or more slides cannot be removed.

        Refuses with TargetNotFoundError when the layout belongs to another presentation, and with
        UnsupportedStructureError when its part carries inbound relationships beyond this
        collection's own, where dropping it would strand whatever else points at it.
        """
        from pptx2._transaction import PackageTransaction

        # Preserve the established error contract before attachment checks.
        if slide_layout.used_by_slides:
            raise ValueError("cannot remove slide-layout in use by one or more slides")

        # Upstream supports isolated collection proxies in its unit-level API contract.
        # A real presentation always supplies a parent and takes the hardened path below.
        if self._parent is None:
            target_idx = self.index(slide_layout)
            target_sldLayoutId = self._sldLayoutIdLst.sldLayoutId_lst[target_idx]
            self._sldLayoutIdLst.remove(target_sldLayoutId)
            slide_layout.slide_master.part.drop_rel(target_sldLayoutId.rId)
            return

        if not isinstance(slide_layout, SlideLayout):
            raise ValueError("slide_layout must be a SlideLayout, got %r" % (slide_layout,))
        if slide_layout.part.package is not self.part.package:
            raise TargetNotFoundError("slide_layout belongs to a different presentation")
        _require_layout_enrolled(slide_layout)

        # ---target layout is identified by its index in this collection---
        target_idx = self.index(slide_layout)

        # --remove layout from p:sldLayoutIds of its master
        # --this stops layout from showing up, but doesn't remove it from package
        target_sldLayoutId = self._sldLayoutIdLst.sldLayoutId_lst[target_idx]
        rId = target_sldLayoutId.rId
        aliases = [
            (owner, alias_rId)
            for owner, alias_rId, _ in _inbound_relationships(self.part.package, slide_layout.part)
            if owner is not self.part or alias_rId != rId
        ]
        if aliases:
            raise UnsupportedStructureError(
                "slide-layout removal refused: layout part has additional inbound "
                "relationship aliases"
            )

        with PackageTransaction(self.part.package, self, slide_layout):
            self._sldLayoutIdLst.remove(target_sldLayoutId)
            if _relationship_references(self.part._element, rId):
                raise UnsupportedStructureError(
                    "slide-layout removal refused: relationship %s remains referenced" % rId
                )
            # --drop relationship from master to layout
            # --this removes layout from package, along with everything (only) it refers to
            self.part.drop_rel(rId)

    def _non_colliding_name(self, base: str) -> str:
        """Return `base`, or `base N` for the smallest N making it unique in this collection."""
        names = {layout.name for layout in self}
        if base not in names:
            return base
        n = 2
        while "%s %d" % (base, n) in names:
            n += 1
        return "%s %d" % (base, n)


class SlideMaster(_BaseMaster):
    """Slide master object.

    Provides access to slide layouts. Access to placeholders, regular shapes, and slide master-level
    properties is inherited from |_BaseMaster|.
    """

    _element: CT_SlideMaster  # pyright: ignore[reportIncompatibleVariableOverride]

    @property
    def header_footers(self) -> HeaderFooters:
        """|HeaderFooters| flags for this master (paper-pptx addition)."""
        return HeaderFooters(self)

    @lazyproperty
    def slide_layouts(self) -> SlideLayouts:
        """|SlideLayouts| object providing access to this slide-master's layouts."""
        return SlideLayouts(self._element.get_or_add_sldLayoutIdLst(), self)


class SlideMasters(ParentedElementProxy):
    """Sequence of |SlideMaster| objects belonging to a presentation.

    Has list access semantics, supporting indexed access, len(), and iteration.
    """

    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, sldMasterIdLst: CT_SlideMasterIdList, parent: Presentation):
        super(SlideMasters, self).__init__(sldMasterIdLst, parent)
        self._sldMasterIdLst = sldMasterIdLst

    def __getitem__(self, idx: int) -> SlideMaster:
        """Provides indexed access, e.g. `slide_masters[2]`."""
        try:
            sldMasterId = self._sldMasterIdLst.sldMasterId_lst[idx]
        except IndexError:
            raise IndexError("slide master index out of range")
        return self.part.related_slide_master(sldMasterId.rId)

    def __iter__(self):
        """Generate each |SlideMaster| instance in the collection, in sequence."""
        for smi in self._sldMasterIdLst.sldMasterId_lst:
            yield self.part.related_slide_master(smi.rId)

    def __len__(self):
        """Support len() built-in function, e.g. `len(slide_masters) == 4`."""
        return len(self._sldMasterIdLst)


class _Background(ElementProxy):
    """Provides access to slide background properties.

    Note that the presence of this object does not by itself imply an
    explicitly-defined background; a slide with an inherited background still
    has a |_Background| object.
    """

    def __init__(self, cSld: CT_CommonSlideData):
        super(_Background, self).__init__(cSld)
        self._cSld = cSld

    @lazyproperty
    def fill(self):
        """|FillFormat| instance for this background.

        This |FillFormat| object is used to interrogate or specify the fill
        of the slide background.

        Note that accessing this property is potentially destructive. A slide
        background can also be specified by a background style reference and
        accessing this property will remove that reference, if present, and
        replace it with NoFill. This is frequently the case for a slide
        master background.

        This is also the case when there is no explicitly defined background
        (background is inherited); merely accessing this property will cause
        the background to be set to NoFill and the inheritance link will be
        interrupted. This is frequently the case for a slide background.

        Of course, if you are accessing this property in order to set the
        fill, then these changes are of no consequence, but the existing
        background cannot be reliably interrogated using this property unless
        you have already established it is an explicit fill.

        If the background is already a fill, then accessing this property
        makes no changes to the current background.
        """
        bgPr = self._cSld.get_or_add_bgPr()
        return FillFormat.from_fill_parent(bgPr)
