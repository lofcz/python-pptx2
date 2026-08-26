"""Native PowerPoint equations from LaTeX.

PowerPoint stores editable equations as Office Math (OMML) wrapped in an
``a14:m`` marker inside a DrawingML paragraph. This module does not compile
LaTeX itself: it calls ``latex2mathml`` then ``mathml2omml`` and wraps the
result so PowerPoint will open it.

Install the converters with::

    pip install "python-pptx2[math]"
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from lxml import etree

from pptx2.oxml import parse_xml
from pptx2.oxml.ns import nsuri, qn

if TYPE_CHECKING:
    from pptx2.oxml.text import CT_OfficeMathMarker


A14_NS = nsuri("a14")
M_NS = nsuri("m")

_WRAPPER_PATTERNS = (
    re.compile(r"^\s*\$\$(.*)\$\$\s*$", re.DOTALL),
    re.compile(r"^\s*\\\[(.*)\\\]\s*$", re.DOTALL),
    re.compile(r"^\s*\\\((.*)\\\)\s*$", re.DOTALL),
    re.compile(r"^\s*\$(.*)\$\s*$", re.DOTALL),
    re.compile(
        r"^\s*\\begin\{(equation\*?|align\*?|displaymath|math)\}"
        r"(.*)\\end\{\1\}\s*$",
        re.DOTALL,
    ),
)

_DISPLAY_JC = {
    "left": "left",
    "right": "right",
    "center": "centerGroup",
    "centre": "centerGroup",
    "justify": "centerGroup",
}


class MathBackendUnavailable(ImportError):
    """Raised when the optional LaTeX → OMML converters are not installed."""


def strip_math_delimiters(source: str) -> str:
    """Return *source* without a single outer ``$…$`` / ``\\[…\\]`` / env wrapper."""
    text = (source or "").strip()
    for pattern in _WRAPPER_PATTERNS:
        match = pattern.match(text)
        if match:
            return match.group(match.lastindex or 1).strip()
    return text


def latex_to_omml(latex: str) -> str:
    """Convert a LaTeX math fragment to an ``<m:oMath>…`` OMML string.

    Accepts a bare fragment or a common wrapper (``$…$``, ``$$…$$``,
    ``\\(…\\)``, ``\\[…\\]``, ``equation`` / ``align`` environments).
    """
    fragment = strip_math_delimiters(latex)
    if not fragment:
        raise ValueError("latex must be a non-empty math fragment")

    converter, mathml2omml = _import_math_backends()
    try:
        mathml = converter.convert(fragment)
    except Exception as exc:
        raise ValueError(f"LaTeX could not be converted to MathML: {exc}") from exc
    try:
        omml = mathml2omml.convert(mathml)
    except Exception as exc:
        raise ValueError(f"MathML could not be converted to Office Math: {exc}") from exc
    if not isinstance(omml, str) or not omml.strip():
        raise ValueError("mathml2omml returned empty OMML")
    return omml.strip()


def office_math_element(
    latex: str,
    *,
    display: bool = False,
    align: str | None = None,
) -> CT_OfficeMathMarker:
    """Return an ``a14:m`` element ready to append to a DrawingML paragraph."""
    payload = _omml_payload(latex_to_omml(latex), display=display, align=align)
    xml = f'<a14:m xmlns:a14="{A14_NS}" xmlns:m="{M_NS}">{payload}</a14:m>'
    return parse_xml(xml)


def style_office_math(
    marker: Any,
    *,
    size_pt: float | None = None,
    color: Any = None,
    font: str | None = None,
) -> None:
    """Write run-level size / color / typeface onto each ``m:r`` inside *marker*."""
    if size_pt is None and color is None and font is None:
        return

    from pptx2._color import coerce_color

    sz = None if size_pt is None else str(int(round(float(size_pt) * 100)))
    rgb = None if color is None else str(coerce_color(color))

    for run in marker.iter(qn("m:r")):
        rPr = run.find(qn("a:rPr"))
        if rPr is None:
            rPr = etree.Element(qn("a:rPr"))
            text = run.find(qn("m:t"))
            if text is not None:
                text.addprevious(rPr)
            else:
                run.append(rPr)
        if sz is not None:
            rPr.set("sz", sz)
        if font is not None:
            latin = rPr.find(qn("a:latin"))
            if latin is None:
                latin = etree.SubElement(rPr, qn("a:latin"))
            latin.set("typeface", font)
        if rgb is not None:
            solid = rPr.find(qn("a:solidFill"))
            if solid is None:
                solid = etree.Element(qn("a:solidFill"))
                rPr.insert(0, solid)
            srgb = solid.find(qn("a:srgbClr"))
            if srgb is None:
                srgb = etree.SubElement(solid, qn("a:srgbClr"))
            srgb.set("val", rgb)


def _omml_payload(omml: str, *, display: bool, align: str | None) -> str:
    if omml.startswith("<m:oMathPara"):
        return omml
    inner = omml if omml.startswith("<m:oMath") else f"<m:oMath>{omml}</m:oMath>"
    if not display:
        return inner
    jc = _DISPLAY_JC.get((align or "center").lower(), "centerGroup")
    return (
        "<m:oMathPara>"
        f'<m:oMathParaPr><m:jc m:val="{jc}"/></m:oMathParaPr>'
        f"{inner}"
        "</m:oMathPara>"
    )


def _import_math_backends():
    try:
        import latex2mathml.converter as latex2mathml_converter
        import mathml2omml
    except ImportError as exc:
        raise MathBackendUnavailable(
            "Native PowerPoint equations require the latex2mathml and "
            "mathml2omml packages. Install with: pip install 'python-pptx2[math]'"
        ) from exc
    return latex2mathml_converter, mathml2omml


__all__ = (
    "MathBackendUnavailable",
    "latex_to_omml",
    "office_math_element",
    "strip_math_delimiters",
    "style_office_math",
)
