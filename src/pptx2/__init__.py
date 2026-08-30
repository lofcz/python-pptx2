"""Initialization module for python-pptx2 package."""

from __future__ import annotations

import sys
from typing import TYPE_CHECKING

import pptx2.exc as exceptions
from pptx2.api import Presentation
from pptx2.audit import AuditReport, audit
from pptx2.geometry import BBox
from pptx2.design.components import (
    ArticleCard,
    Gauge,
    KpiCard,
    ProgressBar,
    StatStrip,
    StatusPill,
    add_article_card,
    add_gauge,
    add_kpi_card,
    add_progress_bar,
    add_stat_strip,
    add_status_pill,
)
from pptx2.design.figures import (
    FigureBackendUnavailable,
    add_html_figure,
    add_matplotlib_figure,
    add_plotly_figure,
    add_svg_figure,
)
from pptx2.math import MathBackendUnavailable
from pptx2.opc.constants import CONTENT_TYPE as CT
from pptx2.opc.package import PartFactory
from pptx2.parts.chart import ChartPart
from pptx2.parts.coreprops import CorePropertiesPart
from pptx2.parts.customprops import CustomPropertiesPart
from pptx2.parts.diagram import (
    DiagramColorsPart,
    DiagramDataPart,
    DiagramLayoutPart,
    DiagramStylePart,
)
from pptx2.parts.image import ImagePart
from pptx2.parts.media import MediaPart
from pptx2.parts.presentation import PresentationPart
from pptx2.parts.slide import (
    NotesMasterPart,
    NotesSlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    SlidePart,
    ThemePart,
)

if TYPE_CHECKING:
    from pptx2.opc.package import Part

__version__ = "2.16.0"

sys.modules["pptx2.exceptions"] = exceptions
del sys

__all__ = [
    "Presentation",
    # First-class rectangular region value object (immutable, splattable
    # into add_shape / add_textbox / add_picture).
    "BBox",
    # One-call deck audit (lint + picture + empty-slide + font + size).
    "audit",
    "AuditReport",
    # Figure adapters — embed Plotly / Matplotlib / SVG / HTML output as
    # slide pictures. Third-party deps are imported lazily on first call.
    "add_plotly_figure",
    "add_matplotlib_figure",
    "add_svg_figure",
    "add_html_figure",
    "FigureBackendUnavailable",
    "MathBackendUnavailable",
    # Shape-level building blocks built on the design tokens.
    "add_kpi_card",
    "add_progress_bar",
    "add_gauge",
    "add_status_pill",
    "add_stat_strip",
    "add_article_card",
    "KpiCard",
    "ProgressBar",
    "Gauge",
    "StatusPill",
    "StatStrip",
    "ArticleCard",
]

content_type_to_part_class_map: dict[str, type[Part]] = {
    CT.PML_PRESENTATION_MAIN: PresentationPart,
    CT.PML_PRES_MACRO_MAIN: PresentationPart,
    CT.PML_TEMPLATE_MAIN: PresentationPart,
    CT.PML_SLIDESHOW_MAIN: PresentationPart,
    CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
    CT.OFC_CUSTOM_PROPERTIES: CustomPropertiesPart,
    CT.PML_NOTES_MASTER: NotesMasterPart,
    CT.PML_NOTES_SLIDE: NotesSlidePart,
    CT.PML_SLIDE: SlidePart,
    CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
    CT.PML_SLIDE_MASTER: SlideMasterPart,
    CT.OFC_THEME: ThemePart,
    CT.DML_CHART: ChartPart,
    CT.DML_DIAGRAM_DATA: DiagramDataPart,
    CT.DML_DIAGRAM_LAYOUT: DiagramLayoutPart,
    CT.DML_DIAGRAM_STYLE: DiagramStylePart,
    CT.DML_DIAGRAM_COLORS: DiagramColorsPart,
    CT.BMP: ImagePart,
    CT.GIF: ImagePart,
    CT.JPEG: ImagePart,
    CT.MS_PHOTO: ImagePart,
    CT.PNG: ImagePart,
    CT.TIFF: ImagePart,
    CT.X_EMF: ImagePart,
    CT.X_WMF: ImagePart,
    CT.ASF: MediaPart,
    CT.AVI: MediaPart,
    CT.MOV: MediaPart,
    CT.MP4: MediaPart,
    CT.MPG: MediaPart,
    CT.MS_VIDEO: MediaPart,
    CT.SWF: MediaPart,
    CT.VIDEO: MediaPart,
    CT.WMV: MediaPart,
    CT.X_MS_VIDEO: MediaPart,
    # -- accommodate "image/jpg" as an alias for "image/jpeg" --
    "image/jpg": ImagePart,
}

PartFactory.part_type_for.update(content_type_to_part_class_map)

del (
    ChartPart,
    CorePropertiesPart,
    CustomPropertiesPart,
    DiagramColorsPart,
    DiagramDataPart,
    DiagramLayoutPart,
    DiagramStylePart,
    ImagePart,
    MediaPart,
    SlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    ThemePart,
    PresentationPart,
    CT,
    PartFactory,
)
