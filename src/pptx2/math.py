"""Native PowerPoint equations from LaTeX.

PowerPoint stores editable equations as Office Math (OMML) wrapped in an
``a14:m`` marker inside a DrawingML paragraph. This module does not compile
LaTeX itself: ``latex2mathml`` turns the fragment into MathML and the
bundled ``pptx2.mathml2omml`` port (from mathml2omml-plus / ECMA-376 §7.1)
turns that into Office Math.

Install the LaTeX front-end with::

    pip install "python-pptx2[math]"
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from lxml import etree

from pptx2.mathml2omml import convert as mathml_to_omml
from pptx2.oxml import parse_xml
from pptx2.oxml.ns import nsuri, qn

if TYPE_CHECKING:
    from pptx2.oxml.text import CT_OfficeMathMarker


A14_NS = nsuri("a14")
M_NS = nsuri("m")
W_NS = nsuri("w")
MC_NS = nsuri("mc")

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

_SLD_OPEN_RE = re.compile(
    r"^(<\?xml[^>]*\?>\s*)?(<p:sld\b)([^>]*)(>)",
    re.DOTALL,
)
_IGNORABLE_RE = re.compile(r'\smc:Ignorable="([^"]*)"')


class MathBackendUnavailable(ImportError):
    """Raised when the optional LaTeX → MathML converter is not installed."""


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

    converter = _import_latex2mathml()
    try:
        mathml = converter.convert(fragment)
    except Exception as exc:
        raise ValueError(f"LaTeX could not be converted to MathML: {exc}") from exc
    try:
        omml = mathml_to_omml(mathml)
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
    xml = (
        f'<a14:m xmlns:a14="{A14_NS}" xmlns:m="{M_NS}" xmlns:w="{W_NS}">'
        f"{payload}</a14:m>"
    )
    return parse_xml(xml)


def style_office_math(
    marker: Any,
    *,
    size_pt: float | None = None,
    color: Any = None,
    font: str | None = None,
) -> None:
    """Write size / color / typeface onto every ``m:r`` and ``m:ctrlPr``."""
    if size_pt is None and color is None and font is None:
        return

    from pptx2._color import coerce_color

    sz = None if size_pt is None else str(int(round(float(size_pt) * 100)))
    rgb = None if color is None else str(coerce_color(color))
    r_pr_xml = _drawing_rpr(sz, rgb, font)

    for run in marker.iter(qn("m:r")):
        _stamp_arpr(run, r_pr_xml, before=qn("m:t"))
    for ctrl in marker.iter(qn("m:ctrlPr")):
        _stamp_arpr(ctrl, r_pr_xml, before=None)


def prepare_slide_xml_for_math(xml: bytes) -> bytes:
    """Ensure the ``p:sld`` root declares math host namespaces.

    PowerPoint requires ``xmlns:m``, ``xmlns:a14``, ``xmlns:mc`` and
    ``mc:Ignorable="a14"`` on the slide (MS-PPTX §2.2.8). Bare ``m:oMath``
    without ``a14:m`` is stripped on open.
    """
    match = _SLD_OPEN_RE.match(xml.decode("utf-8"))
    if match is None:
        return xml
    decl, start, attrs, close = match.group(1) or "", match.group(2), match.group(3), match.group(4)
    for prefix, uri in (
        ("xmlns:m", M_NS),
        ("xmlns:a14", A14_NS),
        ("xmlns:mc", MC_NS),
        ("xmlns:w", W_NS),
    ):
        if f'{prefix}="' not in attrs:
            attrs += f' {prefix}="{uri}"'
    ignorable = _IGNORABLE_RE.search(attrs)
    if ignorable is None:
        attrs += ' mc:Ignorable="a14"'
    else:
        tokens = ignorable.group(1).split()
        if "a14" not in tokens:
            tokens.append("a14")
            attrs = _IGNORABLE_RE.sub(f' mc:Ignorable="{" ".join(tokens)}"', attrs, count=1)
    rewritten = decl + start + attrs + close + xml.decode("utf-8")[match.end() :]
    return rewritten.encode("utf-8")


def _drawing_rpr(sz: str | None, rgb: str | None, font: str | None) -> etree._Element:
    attrs = {"dirty": "0"}
    if sz is not None:
        attrs["sz"] = sz
    r_pr = etree.Element(qn("a:rPr"), attrs)
    if rgb is not None:
        solid = etree.SubElement(r_pr, qn("a:solidFill"))
        etree.SubElement(solid, qn("a:srgbClr")).set("val", rgb)
    if font is not None:
        etree.SubElement(r_pr, qn("a:latin")).set("typeface", font)
    return r_pr


def _stamp_arpr(parent: Any, template: etree._Element, before: str | None) -> None:
    existing = parent.find(qn("a:rPr"))
    stamped = etree.fromstring(etree.tostring(template))
    if existing is not None:
        parent.replace(existing, stamped)
        return
    if before is not None:
        sibling = parent.find(before)
        if sibling is not None:
            sibling.addprevious(stamped)
            return
    parent.append(stamped)


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


def _import_latex2mathml():
    try:
        import latex2mathml.converter as latex2mathml_converter
    except ImportError as exc:
        raise MathBackendUnavailable(
            "Native PowerPoint equations require latex2mathml. "
            "Install with: pip install 'python-pptx2[math]'"
        ) from exc
    return latex2mathml_converter


__all__ = (
    "MathBackendUnavailable",
    "latex_to_omml",
    "mathml_to_omml",
    "office_math_element",
    "prepare_slide_xml_for_math",
    "strip_math_delimiters",
    "style_office_math",
)
