"""Main presentation object."""

from __future__ import annotations

import logging
import os
from contextlib import contextmanager
from typing import IO, TYPE_CHECKING, Literal, cast

from pptx2.section import Sections
from pptx2.shared import PartElementProxy
from pptx2.slide import SlideMasters, Slides
from pptx2.util import lazyproperty

if TYPE_CHECKING:
    from collections.abc import Iterator
    from datetime import datetime

    from pptx2.compose.deck_compose import ImportReport
    from pptx2.enum.presentation import MSO_TRANSITION_TYPE
    from pptx2.lint import LintIssue
    from pptx2.oxml.presentation import CT_Presentation, CT_SlideId
    from pptx2.parts.customprops import CustomProperties
    from pptx2.parts.presentation import PresentationPart
    from pptx2.parts.slide import SlidePart
    from pptx2.slide import NotesMaster, Slide, SlideLayout, SlideLayouts
    from pptx2.util import Length

logger = logging.getLogger(__name__)

# Sentinel used by `set_transition` so callers can distinguish "leave the
# existing value alone" from "explicitly clear it" (which is `None`).
_UNSET = object()

LintOnSaveMode = Literal["off", "warn", "raise"]

#: Accepted values for :attr:`Presentation.lint_on_save`.
_LINT_ON_SAVE_MODES: tuple[LintOnSaveMode, ...] = ("off", "warn", "raise")


class Presentation(PartElementProxy):
    """PresentationML (PML) presentation.

    Not intended to be constructed directly. Use :func:`pptx2.Presentation` to open or
    create a presentation.
    """

    _element: CT_Presentation
    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    # ---- `lint_on_save` is a plain instance attribute, not part of the XML.
    # ---- `PresentationPart.presentation` is a `lazyproperty`, so a package
    # ---- hands back this same proxy object every time and the setting sticks
    # ---- for the life of the package. It is deliberately *not* persisted.
    _lint_on_save: LintOnSaveMode = "off"

    @property
    def core_properties(self):
        """|CoreProperties| instance for this presentation.

        Provides read/write access to the Dublin Core document properties for the presentation.
        """
        return self.part.core_properties

    @property
    def custom_properties(self) -> CustomProperties:
        """|CustomProperties| mapping for this presentation.

        Provides read/write access to the user-defined document properties in
        `/docProps/custom.xml`. Values may be str, int, float, or bool; the first assignment on a
        package without the part creates it.
        """
        return self.part.custom_properties

    @property
    def notes_master(self) -> NotesMaster:
        """Instance of |NotesMaster| for this presentation.

        If the presentation does not have a notes master, one is created from a default template
        and returned. The same single instance is returned on each call.
        """
        return self.part.notes_master

    @property
    def lint_on_save(self) -> LintOnSaveMode:
        """What :meth:`save` does about error-severity lint issues. Read/write.

        One of:

        ``"off"`` *(default)*
            No checks are run at save time; :meth:`save` does no lint work at
            all.  This is the default so that existing code keeps working
            unchanged.

        ``"warn"``
            Every slide is linted before the file is written and each
            error-severity issue is logged (stdlib :mod:`logging`, logger
            ``"pptx2.presentation"``).  The file is still written.

        ``"raise"``
            Every slide is linted *before* the file is written and
            :class:`~pptx2.exc.LintError` is raised if any slide has an
            error-severity issue, so a failing deck never reaches disk.

        Example::

            prs.lint_on_save = "raise"
            prs.save("deck.pptx")   # raises LintError if a shape is off-slide

        This is a setting on the in-memory |Presentation| object; it is not
        stored in the ``.pptx`` file, so a deck re-opened from disk starts
        out at ``"off"`` again.

        Raises:
            ValueError: if assigned anything other than ``"off"``, ``"warn"``,
                or ``"raise"``.
        """
        return self._lint_on_save

    @lint_on_save.setter
    def lint_on_save(self, value: LintOnSaveMode) -> None:
        if value not in _LINT_ON_SAVE_MODES:
            raise ValueError(
                f"lint_on_save must be one of 'off', 'warn', or 'raise', got {value!r}"
            )
        # ---- the `not in` check above narrows `value` to `LintOnSaveMode` ----
        self._lint_on_save = value

    @contextmanager
    def batch(self) -> "Iterator[Presentation]":
        """Validate this deck once at block exit instead of once per mutating call.

        paper-pptx addition. Opt-in; per-edit validation stays the default.

            with prs.batch():
                for slide in prs.slides:
                    slide.shapes[0].text_frame.text = "..."

        Paper's own mutating APIs each run a package transaction that serializes and
        reopens the whole deck before committing. Inside this block those transactions
        nest, and only the outermost validates — one whole-deck check per block rather
        than one per call. Measured 2.8-3.0x on 25-200 slide decks.

        **It also validates operations that had no validation at all.** The mutation
        surface inherited from python-pptx — `text_frame.text`, `add_textbox`,
        `add_picture`, :meth:`Slides.add_slide` and the rest — runs no transaction of its
        own. Inside a block the enclosing transaction covers them, so a sequence that
        previously succeeded and saved an unreadable deck now refuses. That is the
        intended improvement, not a regression, but it does mean working code can start
        refusing once wrapped.

        Costs and limits, all of which argue for scoping a block to a unit of work you
        would be willing to redo:

        - **Entry is not free.** Block entry snapshots every reachable part. An empty
          block costs about what one unbatched edit costs, so a block around a single
          edit breaks even and the gain starts at two.
        - **Rollback granularity is the block.** A refusal discards *every* edit in it,
          not just the offending one.
        - **Only the end state is checked.** A deck that is momentarily invalid inside the
          block but valid at exit commits normally.
        - **Saving inside a block is refused** (|BoundaryViolationError|), because the
          package has not been validated yet and the edits may still roll back. Save after
          the block closes.
        - **Blocks on two different decks must exit in reverse order** of entry, or
          ``RuntimeError``.
        - **Digitally signed decks refuse at block entry**, before any edit runs.
        - :meth:`import_slide` and :meth:`append_deck` return their report before the
          import has been proven reopenable; that proof moves to block exit.

        An exception raised by the caller inside the block rolls the package back and
        propagates unchanged.
        """
        from pptx2._transaction import PackageTransaction

        with PackageTransaction(self.part.package, self):
            yield self

    def save(self, file: str | os.PathLike[str] | IO[bytes]):
        """Writes this presentation to `file`.

        `file` can be either a file-path (a string or `pathlib.Path`) or a
        file-like object open for writing bytes.

        When :attr:`lint_on_save` is ``"warn"`` or ``"raise"``, every slide is
        linted before anything is written; in ``"raise"`` mode a
        :class:`~pptx2.exc.LintError` propagates and no file is written.

        A file-path destination is written atomically, resolving symlinks, so a failure part-way
        through leaves any existing file untouched. A file-like destination is written straight
        through once the whole package has serialized successfully. See
        :meth:`pptx2.opc.package.OpcPackage.save` for what atomic replacement costs.

        :func:`pptx2.package.patch_save` is the narrow-save alternative: atomic too, and it restores
        the original bytes of every part that did not change. Its no-op round trip is byte-identical
        only for a package paper-pptx wrote; see that function for why.

        Refuses with |BoundaryViolationError| while a :meth:`batch` block is open on this package:
        those edits have not been validated yet and may still roll back, so writing them out would
        publish a package the block is not prepared to stand behind. Save once the block has closed.
        """
        from pptx2._transaction import package_has_open_transaction
        from pptx2.errors import BoundaryViolationError

        if package_has_open_transaction(self.part.package):
            raise BoundaryViolationError(
                "cannot save inside an open batch block: the package is not validated "
                "until the block exits and these edits may still roll back; save after "
                "the block closes"
            )
        # ---- the default ("off") does no lint work whatsoever ----
        if self._lint_on_save != "off":
            _lint_before_save(self, self._lint_on_save)
        self.part.save(file)

    def render_thumbnails(self, **kwargs):
        """Render PNG thumbnails for every slide via headless LibreOffice.

        Thin wrapper around :func:`pptx2.render.render_slide_thumbnails` that
        forwards `out_dir`, `slide_indexes`, `soffice_bin`, `timeout`, and
        `return_bytes` keyword arguments.  Requires ``soffice`` (LibreOffice)
        on PATH; raises ``ThumbnailRendererUnavailable`` otherwise.
        """
        from pptx2.render import render_slide_thumbnails

        return render_slide_thumbnails(self, **kwargs)

    def render_contact_sheet(self, out_path, **kwargs):
        """Render every slide and tile the thumbnails into one PNG at `out_path`.

        Thin wrapper around :func:`pptx2.render.render_contact_sheet` — the
        one-image way to eyeball a whole generated deck. Forwards `cols`,
        `thumb_width`, `gap`, `label`, `slides`, `soffice_bin`, `timeout`
        and `dpi`. Requires ``soffice`` (LibreOffice) on PATH.
        """
        from pptx2.render import render_contact_sheet

        return render_contact_sheet(self, out_path, **kwargs)

    @property
    def slide_height(self) -> Length | None:
        """Height of slides in this presentation, in English Metric Units (EMU).

        Returns |None| if no slide width is defined. Read/write.
        """
        sldSz = self._element.sldSz
        if sldSz is None:
            return None
        return sldSz.cy

    @slide_height.setter
    def slide_height(self, height: Length):
        sldSz = self._element.get_or_add_sldSz()
        sldSz.cy = height

    @property
    def slide_layouts(self) -> SlideLayouts:
        """|SlideLayouts| collection belonging to the first |SlideMaster| of this presentation.

        A presentation can have more than one slide master and each master will have its own set
        of layouts. This property is a convenience for the common case where the presentation has
        only a single slide master.
        """
        return self.slide_masters[0].slide_layouts

    @property
    def slide_master(self):
        """
        First |SlideMaster| object belonging to this presentation. Typically,
        presentations have only a single slide master. This property provides
        simpler access in that common case.
        """
        return self.slide_masters[0]

    @lazyproperty
    def slide_masters(self) -> SlideMasters:
        """|SlideMasters| collection of slide-masters belonging to this presentation."""
        return SlideMasters(self._element.get_or_add_sldMasterIdLst(), self)

    @property
    def slide_width(self):
        """
        Width of slides in this presentation, in English Metric Units (EMU).
        Returns |None| if no slide width is defined. Read/write.
        """
        sldSz = self._element.sldSz
        if sldSz is None:
            return None
        return sldSz.cx

    @slide_width.setter
    def slide_width(self, width: Length):
        sldSz = self._element.get_or_add_sldSz()
        sldSz.cx = width

    @property
    def theme(self):
        """Return a |Theme| object providing read-only access to the color palette and fonts.

        Navigates to the theme part of the first slide master, which is where
        Office applications store the active theme.  Returns ``None`` if no
        slide master (and therefore no theme) is present.

        Example::

            from pptx2.enum.dml import MSO_THEME_COLOR

            rgb   = prs.theme.colors[MSO_THEME_COLOR.ACCENT_1]
            major = prs.theme.fonts.major   # e.g. "Calibri"
        """
        return self.slide_master.part.theme

    @lazyproperty
    def slides(self):
        """|Slides| object containing the slides in this presentation."""
        sldIdLst = self._element.get_or_add_sldIdLst()
        self.part.rename_slide_parts([cast("CT_SlideId", sldId).rId for sldId in sldIdLst])
        return Slides(sldIdLst, self)

    @property
    def sections(self) -> Sections:
        """|Sections| collection of named slide groupings in this presentation.

        Sections appear in PowerPoint's outline / slide-sorter pane and are
        stored as a PowerPoint-2010 extension on the presentation part. The
        returned collection supports ``len()``, indexed access, iteration,
        ``.add(name, start_slide_index=None)``, and ``.remove(section)``.

        Reading this property never modifies the deck; the extension
        container is created only when the first section is added.
        """
        return Sections(self._element, self)

    def apply_footers(
        self,
        *,
        footer: str | None = None,
        slide_number: bool = False,
        date_format: str | None = None,
        fixed_date: str | None = None,
        skip_title_slides: bool = False,
        now: "datetime | None" = None,
    ) -> None:
        """Apply the complete footer state to every slide ("Apply to All").

        paper-pptx addition. Persists exactly what PowerPoint's
        Insert > Header & Footer dialog does: materializes minimal `dt`/`ftr`/`sldNum`
        placeholder shapes per slide (binding to the layout furniture by `idx`), writes
        slide numbers and automatic dates as real `a:fld` elements whose cached text
        consumers refresh on open, and *removes* the placeholders for unchecked elements —
        each call sets the full three-element state, like the dialog.

        `footer`: literal footer text, or None to remove the footer placeholder.
        `slide_number`: True writes a `slidenum` field cached with the current position
        (honoring `firstSlideNum`); consumers renumber live after any reorder.
        `date_format`: a `datetime`..`datetime13` token for an automatically-updating date
        field; `fixed_date`: literal date text (the dialog's "Fixed" mode); passing both
        raises |ValueError|. `now` seeds the date field's cached text (None = wall clock);
        the package never vouches for cached values — they are consumer-refreshed hints.
        `skip_title_slides`: the dialog's "Don't show on title slide" — slides on a
        `type="title"` layout get the all-removed state.

        Refuses atomically (|UnsupportedStructureError|, validated deck-wide before the
        first write) when a wanted element has no layout furniture to inherit from, or
        when explicit `p:hf` flags on a layout/master disable it (clear those via
        `header_footers` first — this API never flips them silently).
        """
        from pptx2.hf import apply_presentation_footers

        apply_presentation_footers(
            self,
            footer=footer,
            slide_number=slide_number,
            date_format=date_format,
            fixed_date=fixed_date,
            skip_title_slides=skip_title_slides,
            now=now,
        )

    def append_deck(
        self, source_prs: "Presentation", *, mode: str, notes: bool = True
    ) -> "tuple[ImportReport, ...]":
        """Import every slide of `source_prs`, in order, at the end of this deck.

        paper-pptx addition, built on :meth:`import_slide` — same `mode`
        semantics and refusal ledger. The COMPLETE source deck validates before the first
        write: a refusal on any source slide leaves this presentation untouched. Source
        sections are not copied (this deck's section structure governs — declared).
        """
        from pptx2.compose import append_deck

        return append_deck(self, source_prs, mode=mode, notes=notes)

    def import_slide(
        self,
        source_prs: "Presentation",
        slide: Slide,
        *,
        mode: str,
        position: int | None = None,
        notes: bool = True,
        section: str | None = None,
        section_id: str | None = None,
        target_layout: SlideLayout | None = None,
        placeholder_map="auto",
    ) -> "ImportReport":
        """Import `slide` from `source_prs` into this presentation; return the report.

        paper-pptx addition. `mode` is required — there is no right
        default, the caller chooses consciously:

        - ``"adopt_theme"``: content transplants and rebinds to a destination layout
          (auto only on a unique exact layout name, then a unique exact non-custom layout
          type; `target_layout` overrides; orphan placeholders bake from their
          source-resolved look). The slide takes the house style; every run whose
          resolved values changed is in `run_shifts`.
        - ``"keep_appearance"``: the source layout+master+theme chain transplants,
          fingerprint-deduplicated (ten slides from one source share one master).
        - ``"bake"``: resolvable effective values become explicit local properties,
          furniture placeholders (dt/ftr/sldNum) drop, remaining placeholders become
          free shapes, and the slide attaches to a destination layout selected by the
          same unique name/type tiers, then a unique blank-layout fallback. It never
          falls back to the first layout. Stable look without importing masters.

        The source presentation is never mutated. Media always copies (never shared
        across packages); charts deep-copy with workbooks; SmartArt carries opaquely;
        comments drop (reported); OLE objects, controls, internal slide links, and
        unknown relationship types refuse (`RelationshipPolicyError`) before any write.
        `notes` copies the speaker-notes part re-linked to this deck's notes master.
        `section` selects an existing destination section by unique exact name;
        `section_id` selects by exact stored GUID when names collide. They are mutually
        exclusive. With neither selector, enrollment remains adjacent to the insertion
        point when this deck has sections. Missing selectors raise
        |TargetNotFoundError| and duplicate matches raise |AmbiguousTargetError| before
        any write.
        Multiple candidates at any automatic layout tier raise |AmbiguousTargetError|
        before any write and list the layouts; pass an enrolled destination
        `target_layout` to resolve that choice explicitly.
        Adopt-theme placeholder reconciliation first preserves exact type+idx matches,
        then accepts same-type or compatible-family fallbacks only when unique. Pass a
        partial `placeholder_map={source_idx: target_idx | None}` to resolve ambiguity;
        `None` deliberately orphans and bakes that source placeholder. The argument does
        not apply to keep-appearance or bake imports. Whole-deck append remains automatic-only.
        """
        from pptx2.compose import import_slide

        return import_slide(
            self,
            source_prs,
            slide,
            mode=mode,
            position=position,
            notes=notes,
            section=section,
            section_id=section_id,
            target_layout=target_layout,
            placeholder_map=placeholder_map,
        )

    def apply_template(
        self,
        template_path_or_stream: str | IO[bytes],
    ) -> None:
        """Re-point every slide in this presentation at masters from *template_path_or_stream*.

        After this call every slide inherits theme, fonts, and colours from
        the template.  Slide content (shapes, text, animations) is preserved.

        The template can be a ``.potx`` or any ``.pptx``/``.pptm`` file — the
        master(s) inside are used regardless of the extension.

        Layout matching (in priority order):

        1. Same ``<p:cSld name="…">`` name.
        2. Same layout ``type`` attribute (e.g. ``"title"``, ``"obj"``).
        3. Fall back to the template's first layout.

        Parameters
        ----------
        template_path_or_stream:
            Path to a ``.potx``/``.pptx`` file, or a file-like object.

        Example::

            prs = pptx2.Presentation("deck.pptx")
            prs.apply_template("brand.potx")
            prs.save("branded_deck.pptx")
        """
        from pptx2._template_applier import apply_template as _apply_template
        from pptx2.api import Presentation as _Presentation

        tpl = _Presentation(template_path_or_stream)
        _apply_template(self.part, tpl.part)

    def set_transition(
        self,
        kind: "MSO_TRANSITION_TYPE | None" = cast("MSO_TRANSITION_TYPE", _UNSET),
        *,
        duration: int | None = cast(int, _UNSET),
        advance_on_click: bool | None = cast(bool, _UNSET),
        advance_after: int | None = cast(int, _UNSET),
        force: bool = False,
    ) -> None:
        """Apply a transition to every slide in this presentation.

        Convenience for the common "give the whole deck the same transition"
        case; equivalent to looping over :attr:`slides` and assigning to each
        slide's :attr:`~pptx2.slide.Slide.transition` properties.

        Any argument left unspecified is left untouched on each slide, so
        partial updates (e.g. only changing ``duration``) are safe::

            from pptx2.enum.presentation import MSO_TRANSITION

            prs.set_transition(MSO_TRANSITION.MORPH, duration=750)

            # later, just bump the duration without disturbing the kind
            prs.set_transition(duration=500)

        Passing ``kind=None`` clears the transition element on every slide
        (restoring inheritance/defaults).  Passing ``duration=None``,
        ``advance_on_click=None``, or ``advance_after=None`` clears that
        individual attribute on every slide.

        Per-slide overrides are preserved by default
        ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        When ``kind`` is supplied, slides that already have an explicit
        transition kind are **left untouched** — the deck-wide call no
        longer silently clobbers an earlier ``slide.transition.kind = …``.
        This is the historical footgun that prompted item 2 in
        ``IMPROVEMENT_PLAN.md``.  To restore the old "force every slide to
        match" behaviour, pass ``force=True``::

            slide_2.transition.kind = MSO_TRANSITION_TYPE.MORPH
            prs.set_transition(MSO_TRANSITION_TYPE.FADE)              # slide 2 keeps MORPH
            prs.set_transition(MSO_TRANSITION_TYPE.FADE, force=True)  # slide 2 → FADE

        ``duration``, ``advance_on_click``, and ``advance_after`` are
        always applied to every slide regardless of ``force``; only
        ``kind`` participates in the override-preservation behaviour.
        """
        for slide in self.slides:
            transition = slide.transition
            if kind is not _UNSET:
                # Skip slides that already have an explicit kind unless
                # the caller has opted in to force-overwrite.  The
                # ``transition.kind`` getter returns ``None`` only when
                # no ``<p:transition>`` element is present.  An explicit
                # ``<p:transition/>`` reads back as
                # ``MSO_TRANSITION_TYPE.NONE`` and counts as an existing
                # explicit choice (the slide author asked for "no
                # transition"), so it is preserved unless ``force=True``.
                #
                # ``kind=None`` on the caller side means "clear" — apply
                # unconditionally so callers can still wipe the
                # transition off every slide in one call without an
                # explicit ``force=True``.
                if kind is None or transition.kind is None or force:
                    transition.kind = kind
            if duration is not _UNSET:
                transition.duration = duration
            if advance_on_click is not _UNSET:
                transition.advance_on_click = advance_on_click
            if advance_after is not _UNSET:
                transition.advance_after = advance_after


def _lint_before_save(prs: Presentation, mode: LintOnSaveMode) -> None:
    """Lint every slide of `prs` and warn or raise per `mode`.

    Mirrors the ``lint`` option of :func:`pptx2.compose.from_spec`: only
    ERROR-severity issues are acted on, and ``"raise"`` reports every offending
    issue in a single :class:`~pptx2.exc.LintError`.
    """
    from pptx2.exc import LintError
    from pptx2.lint import LintSeverity

    errors: list[tuple[int, LintIssue]] = [
        (idx, issue)
        for idx, slide in enumerate(prs.slides)
        for issue in slide.lint().issues
        if issue.severity == LintSeverity.ERROR
    ]

    if not errors:
        return

    if mode == "warn":
        for idx, issue in errors:
            logger.warning("pptx lint: slide %d: %s", idx, issue)
        return

    msgs = "; ".join(f"slide {idx}: {issue}" for idx, issue in errors)
    raise LintError(f"Lint errors in presentation: {msgs}")
