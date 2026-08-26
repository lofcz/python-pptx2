"""MathML → Office Math (OMML), ported from ``mathml2omml-plus``.

Faithful Python port of https://github.com/lofcz/mathml2omml (ECMA-376
Part 1 §7.1). This module emits ``<m:oMath>`` only. The host
(``pptx2.math``) wraps the fragment in PowerPoint's ``a14:m`` marker.

Do not invent a second math grammar here — walk MathML and emit the
OMML elements the spec names.
"""

from __future__ import annotations

import re
from typing import Any
from xml.sax.saxutils import escape as _xml_escape

from lxml import etree

MATH_NS = "http://www.w3.org/1998/Math/MathML"
OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"

_NARY_RE = re.compile(r"^[\u220f-\u2211]|[\u2229-\u2233]|[\u22c0-\u22c3]$")
_GROW_RE = re.compile(
    r"^\u220f|\u2211|[\u2229-\u222b]|\u222e|\u222f|\u2232|\u2233|[\u22c0-\u22c3]$"
)
_SKIPPED_SUBTREES = frozenset({"annotation", "annotation-xml"})
_ARG_PARENTS = frozenset(
    {"m:deg", "m:den", "m:e", "m:fName", "m:lim", "m:num", "m:sub", "m:sup"}
)
_STYLES = {"bold": "b", "italic": "i", "bold-italic": "bi"}
_UPPER_COMBINATION = {
    "\u2190": "\u20d6",
    "\u27f5": "\u20d6",
    "\u2192": "\u20d7",
    "\u27f6": "\u20d7",
    "\u00b4": "\u0301",
    "\u02dd": "\u030b",
    "\u02d8": "\u0306",
    "ˇ": "\u030c",
    "\u00b8": "\u0312",
    "\u005e": "\u0302",
    "\u00a8": "\u0308",
    "\u02d9": "\u0307",
    "\u0060": "\u0300",
    "\u002d": "\u0305",
    "\u00af": "\u0305",
    "\u2212": "\u0305",
    "\u002e": "\u0307",
    "\u007e": "\u0303",
    "\u02dc": "\u0303",
}

_HANDLERS: dict[str, Any] = {}


class Node:
    """Mutable tree node matching the JS converter's object shape."""

    __slots__ = (
        "name",
        "type",
        "attribs",
        "children",
        "data",
        "is_nary",
        "style",
        "has_mglyph_child",
        "pending_brk",
        "void_element",
        "comment",
    )

    def __init__(
        self,
        name: str = "",
        type: str = "tag",
        attribs: dict[str, str] | None = None,
        children: list[Node] | None = None,
        data: str = "",
    ) -> None:
        self.name = name
        self.type = type
        self.attribs = attribs if attribs is not None else {}
        self.children = children if children is not None else []
        self.data = data
        self.is_nary = False
        self.style: dict[str, str] | None = None
        self.has_mglyph_child = False
        self.pending_brk = False
        self.void_element = False
        self.comment = ""


def convert(mathml: str) -> str:
    """Convert a MathML document or fragment to an ``<m:oMath>`` string."""
    if not isinstance(mathml, str) or not mathml.strip():
        raise ValueError("MathML must be a non-empty string")
    roots = _parse_mathml(mathml)
    out = Node(type="root")
    _walk(Node(type="root", children=roots), out)
    if out.type != "tag" or not out.name:
        raise ValueError("mathml2omml produced no m:oMath")
    xml = _stringify(out)
    if "<m:oMath" not in xml:
        raise ValueError("mathml2omml produced no m:oMath")
    return xml


def _parse_mathml(source: str) -> list[Node]:
    text = source.strip()
    if not text.startswith("<"):
        text = f'<math xmlns="{MATH_NS}">{text}</math>'
    try:
        root = etree.fromstring(text.encode("utf-8"))
    except etree.XMLSyntaxError as exc:
        raise ValueError(f"MathML is not well-formed XML: {exc}") from exc
    return [_elem_to_node(root)]


def _local_name(tag: str) -> str:
    if isinstance(tag, str) and tag.startswith("{"):
        return tag[tag.rfind("}") + 1 :]
    return str(tag)


def _elem_to_node(el: etree._Element) -> Node:
    attribs: dict[str, str] = {}
    for key, value in el.attrib.items():
        if key == f"{{{XML_NS}}}space":
            attribs["xml:space"] = value
        else:
            attribs[_local_name(str(key))] = value
    node = Node(name=_local_name(el.tag), type="tag", attribs=attribs)
    if el.text:
        node.children.append(Node(type="text", data=el.text))
    for child in el:
        node.children.append(_elem_to_node(child))
        if child.tail:
            node.children.append(Node(type="text", data=child.tail))
    return node


def _escape_text(value: str) -> str:
    return (
        _xml_escape(str(value), {"\u200b": "&#x200B;"})
        .replace("\u200b", "&#x200B;")
    )


def _escape_attr(value: str) -> str:
    return _xml_escape(str(value), {'"': "&quot;"}).replace('"', "&quot;")


def _stringify(doc: Node) -> str:
    if doc.type == "text":
        return _escape_text(doc.data)
    if doc.type == "comment":
        return f"<!--{doc.comment}-->"
    if doc.type != "tag":
        return "".join(_stringify(child) for child in doc.children)
    attrs = "".join(f' {key}="{_escape_attr(val)}"' for key, val in doc.attribs.items())
    void_element = doc.void_element or (
        not doc.children and doc.attribs.get("xml:space") != "preserve"
    )
    if void_element:
        return f"<{doc.name}{attrs}/>"
    inner = "".join(_stringify(child) for child in doc.children)
    return f"<{doc.name}{attrs}>{inner}</{doc.name}>"


def _tag(
    name: str,
    attribs: dict[str, str] | None = None,
    children: list[Node] | None = None,
) -> Node:
    return Node(name=name, type="tag", attribs=attribs, children=children)


def _text_content(node: Node | None, trim: bool = True) -> str:
    if node is None:
        return ""
    if node.type == "text":
        text = node.data.replace("\u2062", "").replace("\u200b", "")
        return text.strip() if trim else text
    return "".join(_text_content(child, trim) for child in node.children)


def _attr_true(element: Node, name: str) -> bool:
    return (element.attribs.get(name) or "").lower() == "true"


def _nary_char(node: Node) -> str | None:
    text = _text_content(node)
    if _NARY_RE.search(text):
        return text
    return None


def _nary_base_arg(parent: Node) -> Node | None:
    if not parent.children:
        return None
    last = parent.children[-1]
    if last.name != "m:nary":
        return None
    return last.children[-1] if last.children else None


def _nary_target(
    nary_char: str,
    element: Node,
    lim_loc: str,
    sub_hide: bool = False,
    sup_hide: bool = False,
) -> Node:
    stretchy = element.attribs.get("stretchy")
    if stretchy == "true":
        grow = "on"
    elif stretchy == "false":
        grow = "off"
    else:
        grow = "on" if _GROW_RE.search(nary_char) else "off"
    return _tag(
        "m:nary",
        children=[
            _tag(
                "m:naryPr",
                children=[
                    _tag("m:chr", {"m:val": nary_char}),
                    _tag("m:limLoc", {"m:val": lim_loc}),
                    _tag("m:grow", {"m:val": grow}),
                    _tag("m:subHide", {"m:val": "on" if sub_hide else "off"}),
                    _tag("m:supHide", {"m:val": "on" if sup_hide else "off"}),
                ],
            )
        ],
    )


def _add_scriptlevel(target: Node, ancestors: list[Node]) -> None:
    scriptlevel = None
    for ancestor in ancestors:
        if ancestor.attribs.get("scriptlevel") is not None:
            scriptlevel = ancestor.attribs.get("scriptlevel")
            break
    if scriptlevel in {"0", "1", "2"}:
        target.children.insert(
            0,
            _tag("m:argPr", children=[_tag("m:scrLvl", {"m:val": scriptlevel})]),
        )


def _style_of(
    element: Node, ancestors: list[Node], previous_style: dict[str, str] | None
) -> dict[str, str]:
    previous_style = previous_style or {}

    def inherited(attr: str, mstyle_attr: str | None = None) -> str:
        if element.attribs.get(attr):
            return element.attribs[attr]
        key = mstyle_attr or attr
        for ancestor in ancestors:
            if ancestor.name == "mstyle" and ancestor.attribs.get(key):
                return ancestor.attribs[key]
        return ""

    color = inherited("mathcolor", "color")
    size = inherited("mathsize")
    scriptlevel = inherited("scriptlevel")
    background = inherited("mathbackground")
    variant = inherited("mathvariant")
    if variant == "b-i":
        variant = "bold-italic"
    fontweight = inherited("fontweight")
    if fontweight == "bold" and variant not in {"bold", "bold-italic"}:
        variant = "bold-italic" if "italic" in variant else "bold"
    elif fontweight == "normal" and variant in {"bold", "bold-italic"}:
        variant = "italic" if "italic" in variant else ""
    fontstyle = inherited("fontstyle")
    if fontstyle == "italic" and variant not in {"italic", "bold-italic"}:
        variant = "bold-italic" if "bold" in variant else "italic"
    elif fontstyle == "normal" and variant in {"italic", "bold-italic"}:
        variant = "bold" if "bold" in variant else ""
    if not element.attribs.get("mathvariant"):
        text = _text_content(element)
        if previous_style.get("variant") == "" and (
            (element.name == "mi" and len(text) > 1)
            or (element.name == "mn" and not re.match(r"^\d+\.\d+$", text))
        ):
            variant = ""
        elif element.name in {"mi", "mn", "mo"} and previous_style.get("variant") in {
            "italic",
            "bold-italic",
        }:
            variant = "bold-italic" if fontweight == "bold" else "italic"
    return {
        "color": color,
        "variant": variant,
        "size": size,
        "scriptlevel": scriptlevel,
        "background": background,
        "fontstyle": fontstyle,
    }


def _ensure_math_text_target(target_parent: Node) -> Node:
    if target_parent.name == "m:t":
        return target_parent
    if target_parent.children:
        last = target_parent.children[-1]
        if last.name == "m:r" and last.children:
            t = last.children[-1]
            if t.name == "m:t":
                return t
    t = _tag("m:t", {"xml:space": "preserve"})
    target_parent.children.append(_tag("m:r", children=[t]))
    return t


def _consume_pending_brk(target_parent: Node, r_element: Node) -> None:
    if not target_parent.pending_brk:
        return
    target_parent.pending_brk = False
    r_pr = next((c for c in r_element.children if c.name == "m:rPr"), None)
    if r_pr is None:
        r_pr = _tag("m:rPr")
        wr_index = next(
            (i for i, c in enumerate(r_element.children) if c.name == "w:rPr"), -1
        )
        r_element.children.insert(wr_index + 1, r_pr)
    r_pr.children.append(_tag("m:brk"))


def _wrap_last_child_in_break_box(target_parent: Node) -> None:
    if not target_parent.pending_brk or not target_parent.children:
        return
    last = target_parent.children[-1]
    if last.name == "m:r":
        return
    target_parent.pending_brk = False
    target_parent.children[-1] = _tag(
        "m:box",
        children=[
            _tag("m:boxPr", children=[_tag("m:brk")]),
            _tag("m:e", children=[last]),
        ],
    )


def _walk(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None = None,
    next_sibling: Node | None = None,
    ancestors: list[Node] | None = None,
) -> None:
    ancestors = ancestors or []
    if element.name in _SKIPPED_SUBTREES:
        return
    if previous_sibling is not None and previous_sibling.is_nary:
        e = _nary_base_arg(target_parent)
        if e is not None:
            target_parent = e
    if not previous_sibling and target_parent.name in _ARG_PARENTS:
        _add_scriptlevel(target_parent, ancestors)
    name_or_type = element.name or element.type
    handler = _HANDLERS.get(name_or_type)
    if handler is not None:
        target_element = handler(
            element, target_parent, previous_sibling, next_sibling, ancestors
        )
    else:
        target_element = target_parent
    if target_element is None:
        return
    if element.children:
        next_ancestors = [element, *ancestors]
        for i, child in enumerate(element.children):
            prev = element.children[i - 1] if i else None
            nxt = element.children[i + 1] if i + 1 < len(element.children) else None
            _walk(child, target_element, prev, nxt, next_ancestors)
            _wrap_last_child_in_break_box(target_element)


def _handle_math(
    element: Node, target_parent: Node, *_args: Any
) -> Node:
    target_parent.name = "m:oMath"
    target_parent.type = "tag"
    target_parent.attribs = {"xmlns:m": OMML_NS, "xmlns:w": W_NS}
    target_parent.children = []
    return target_parent


def _handle_passthrough(
    element: Node, target_parent: Node, *_args: Any
) -> Node:
    return target_parent


def _handle_mrow(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    *_args: Any,
) -> Node:
    if previous_sibling is not None and previous_sibling.is_nary:
        e = _nary_base_arg(target_parent)
        if e is not None:
            return e
    return target_parent


def _handle_mfrac(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node | None:
    if len(element.children) != 2:
        return target_parent
    numerator, denumerator = element.children
    num_target = _tag("m:num")
    den_target = _tag("m:den")
    next_ancestors = [element, *ancestors]
    _walk(numerator, num_target, None, None, next_ancestors)
    _walk(denumerator, den_target, None, None, next_ancestors)
    linethickness = (element.attribs.get("linethickness") or "").strip()
    if linethickness and re.match(r"^0+(\.0*)?([a-z%]+)?$", linethickness, re.I):
        frac_type = "noBar"
    elif element.attribs.get("bevelled") == "true":
        frac_type = "skw"
    else:
        frac_type = "bar"
    target_parent.children.append(
        _tag(
            "m:f",
            children=[
                _tag("m:fPr", children=[_tag("m:type", {"m:val": frac_type})]),
                num_target,
                den_target,
            ],
        )
    )
    return None


def _handle_msqrt(
    element: Node, target_parent: Node, *_args: Any
) -> Node:
    inner = _tag("m:e")
    target_parent.children.append(
        _tag(
            "m:rad",
            children=[
                _tag("m:radPr", children=[_tag("m:degHide", {"m:val": "on"})]),
                _tag("m:deg"),
                inner,
            ],
        )
    )
    return inner


def _handle_mroot(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node | None:
    if len(element.children) != 2:
        return target_parent
    next_ancestors = [element, *ancestors]
    base, root = element.children
    base_target = _tag("m:e")
    root_target = _tag("m:deg")
    _walk(base, base_target, None, None, next_ancestors)
    _walk(root, root_target, None, None, next_ancestors)
    hide = "off" if _text_content(root) else "on"
    target_parent.children.append(
        _tag(
            "m:rad",
            children=[
                _tag("m:radPr", children=[_tag("m:degHide", {"m:val": hide})]),
                root_target,
                base_target,
            ],
        )
    )
    return None


def _script_pr(name: str) -> Node:
    return _tag(name, children=[_tag("m:ctrlPr")])


def _handle_msub(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node | None:
    if len(element.children) != 2:
        return target_parent
    next_ancestors = [element, *ancestors]
    base, subscript = element.children
    nary = _nary_char(base)
    if nary and not _attr_true(element, "accent") and not _attr_true(element, "accentunder"):
        top = _nary_target(nary, element, "subSup", False, True)
        element.is_nary = True
    else:
        base_target = _tag("m:e")
        _walk(base, base_target, None, None, next_ancestors)
        top = _tag("m:sSub", children=[_script_pr("m:sSubPr"), base_target])
    sub_target = _tag("m:sub")
    _walk(subscript, sub_target, None, None, next_ancestors)
    top.children.append(sub_target)
    if element.is_nary:
        top.children.append(_tag("m:sup"))
        top.children.append(_tag("m:e"))
    target_parent.children.append(top)
    return None


def _handle_msup(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node | None:
    if len(element.children) != 2:
        return target_parent
    next_ancestors = [element, *ancestors]
    base, superscript = element.children
    nary = _nary_char(base)
    if nary and not _attr_true(element, "accent") and not _attr_true(element, "accentunder"):
        top = _nary_target(nary, element, "subSup", True)
        element.is_nary = True
        top.children.append(_tag("m:sub"))
    else:
        base_target = _tag("m:e")
        _walk(base, base_target, None, None, next_ancestors)
        top = _tag("m:sSup", children=[_script_pr("m:sSupPr"), base_target])
    sup_target = _tag("m:sup")
    _walk(superscript, sup_target, None, None, next_ancestors)
    top.children.append(sup_target)
    if element.is_nary:
        top.children.append(_tag("m:e"))
    target_parent.children.append(top)
    return None


def _handle_msubsup(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node | None:
    if len(element.children) != 3:
        return target_parent
    next_ancestors = [element, *ancestors]
    base, subscript, superscript = element.children
    nary = _nary_char(base)
    if nary and not _attr_true(element, "accent") and not _attr_true(element, "accentunder"):
        top = _nary_target(nary, element, "subSup")
        element.is_nary = True
    else:
        base_target = _tag("m:e")
        _walk(base, base_target, None, None, next_ancestors)
        top = _tag("m:sSubSup", children=[_script_pr("m:sSubSupPr"), base_target])
    sub_target = _tag("m:sub")
    sup_target = _tag("m:sup")
    _walk(subscript, sub_target, None, None, next_ancestors)
    _walk(superscript, sup_target, None, None, next_ancestors)
    top.children.append(sub_target)
    top.children.append(sup_target)
    if element.is_nary:
        top.children.append(_tag("m:e"))
    target_parent.children.append(top)
    return None


def _handle_under_or_over(
    element: Node,
    target_parent: Node,
    ancestors: list[Node],
    direction: str,
) -> Node | None:
    if len(element.children) != 2:
        return target_parent
    next_ancestors = [element, *ancestors]
    base, script = element.children
    nary = _nary_char(base)
    if nary and not _attr_true(element, "accent") and not _attr_true(element, "accentunder"):
        top = _nary_target(
            nary, element, "undOvr", direction == "over", direction == "under"
        )
        element.is_nary = True
        sub_target = _tag("m:sub")
        sup_target = _tag("m:sup")
        _walk(
            script,
            sub_target if direction == "under" else sup_target,
            None,
            None,
            next_ancestors,
        )
        top.children.extend([sub_target, sup_target, _tag("m:e")])
        target_parent.children.append(top)
        return None

    script_text = _text_content(script)
    base_target = _tag("m:e")
    _walk(base, base_target, None, None, next_ancestors)

    if (
        direction == "under"
        and script.name == "mo"
        and script_text in {"\u0332", "\u005f"}
    ) or (
        direction == "over"
        and script.name == "mo"
        and script_text in {"\u0305", "\u00af"}
    ):
        script_name = "m:sSub" if direction == "under" else "m:sSup"
        script_pr = "m:sSubPr" if direction == "under" else "m:sSupPr"
        target_parent.children.append(
            _tag(
                "m:bar",
                children=[
                    _tag(
                        "m:barPr",
                        children=[
                            _tag(
                                "m:pos",
                                {"m:val": "bot" if direction == "under" else "top"},
                            )
                        ],
                    ),
                    _tag(
                        "m:e",
                        children=[
                            _tag(
                                script_name,
                                children=[
                                    _script_pr(script_pr),
                                    base_target,
                                    _tag("m:sub"),
                                ],
                            )
                        ],
                    ),
                ],
            )
        )
        return None

    accent_ok = (
        direction == "under"
        and _attr_true(element, "accentunder")
        and script.name == "mo"
        and len(script_text) < 2
    ) or (
        direction == "over"
        and _attr_true(element, "accent")
        and script.name == "mo"
        and len(script_text) < 2
    )
    if accent_ok:
        target_parent.children.append(
            _tag(
                "m:acc",
                children=[
                    _tag(
                        "m:accPr",
                        children=[
                            _tag(
                                "m:chr",
                                {
                                    "m:val": _UPPER_COMBINATION.get(
                                        script_text, script_text
                                    )
                                },
                            )
                        ],
                    ),
                    base_target,
                ],
            )
        )
        return None

    if (
        not _attr_true(element, "accent")
        and not _attr_true(element, "accentunder")
        and script.name == "mo"
        and base.name == "mrow"
        and len(script_text) == 1
    ):
        target_parent.children.append(
            _tag(
                "m:groupChr",
                children=[
                    _tag(
                        "m:groupChrPr",
                        children=[
                            _tag("m:chr", {"m:val": script_text}),
                            _tag(
                                "m:pos",
                                {"m:val": "bot" if direction == "under" else "top"},
                            ),
                        ],
                    ),
                    base_target,
                ],
            )
        )
        return None

    script_target = _tag("m:lim")
    _walk(script, script_target, None, None, next_ancestors)
    target_parent.children.append(
        _tag(
            "m:limLow" if direction == "under" else "m:limUpp",
            children=[base_target, script_target],
        )
    )
    return None


def _handle_munder(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node | None:
    return _handle_under_or_over(element, target_parent, ancestors, "under")


def _handle_mover(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node | None:
    return _handle_under_or_over(element, target_parent, ancestors, "over")


def _handle_munderover(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node | None:
    if len(element.children) != 3:
        return target_parent
    next_ancestors = [element, *ancestors]
    base, underscript, overscript = element.children
    nary = _nary_char(base)
    if nary and not _attr_true(element, "accent") and not _attr_true(element, "accentunder"):
        top = _nary_target(nary, element, "undOvr")
        element.is_nary = True
        sub_target = _tag("m:sub")
        sup_target = _tag("m:sup")
        _walk(underscript, sub_target, None, None, next_ancestors)
        _walk(overscript, sup_target, None, None, next_ancestors)
        top.children.extend([sub_target, sup_target, _tag("m:e")])
        target_parent.children.append(top)
        return None
    base_target = _tag("m:e")
    _walk(base, base_target, None, None, next_ancestors)
    under_target = _tag("m:lim")
    over_target = _tag("m:lim")
    _walk(underscript, under_target, None, None, next_ancestors)
    _walk(overscript, over_target, None, None, next_ancestors)
    target_parent.children.append(
        _tag(
            "m:limUpp",
            children=[
                _tag("m:e", children=[_tag("m:limLow", children=[base_target, under_target])]),
                over_target,
            ],
        )
    )
    return None


def _handle_mtable(
    element: Node, target_parent: Node, *_args: Any
) -> Node:
    cells = max((len(row.children) for row in element.children), default=0)
    for row in element.children:
        while len(row.children) < cells:
            row.children.append(_tag("mtd"))
    matrix = _tag(
        "m:m",
        children=[
            _tag(
                "m:mPr",
                children=[
                    _tag("m:baseJc", {"m:val": "center"}),
                    _tag("m:plcHide", {"m:val": "on"}),
                    _tag(
                        "m:mcs",
                        children=[
                            _tag(
                                "m:mc",
                                children=[
                                    _tag(
                                        "m:mcPr",
                                        children=[
                                            _tag("m:count", {"m:val": str(cells)}),
                                            _tag("m:mcJc", {"m:val": "center"}),
                                        ],
                                    )
                                ],
                            )
                        ],
                    ),
                ],
            )
        ],
    )
    target_parent.children.append(matrix)
    return matrix


def _handle_mtr(element: Node, target_parent: Node, *_args: Any) -> Node:
    row = _tag("m:mr")
    target_parent.children.append(row)
    return row


def _handle_mtd(element: Node, target_parent: Node, *_args: Any) -> Node:
    cell = _tag("m:e")
    target_parent.children.append(cell)
    return cell


def _hide(name: str) -> Node:
    return _tag(name, {"m:val": "on"})


def _handle_menclose(
    element: Node, target_parent: Node, *_args: Any
) -> Node:
    notation = (element.attribs.get("notation") or "longdiv").split(" ")[0]
    inner = _tag("m:e")
    if notation == "longdiv":
        target_parent.children.append(
            _tag(
                "m:rad",
                children=[
                    _tag("m:radPr", children=[_tag("m:degHide", {"m:val": "on"})]),
                    _tag("m:deg"),
                    inner,
                ],
            )
        )
        return inner
    hide = {
        "t": _hide("m:hideTop"),
        "b": _hide("m:hideBot"),
        "l": _hide("m:hideLeft"),
        "r": _hide("m:hideRight"),
    }
    pr = _tag("m:borderBoxPr")
    box = _tag("m:borderBox")
    if notation in {"actuarial", "radical", "box"}:
        box.children = [inner]
    elif notation in {"left", "roundedbox"}:
        pr.children = [hide["t"], hide["b"], hide["r"]]
        box.children = [pr, inner]
    elif notation in {"right", "circle"}:
        pr.children = [hide["t"], hide["b"], hide["l"]]
        box.children = [pr, inner]
    elif notation == "top":
        pr.children = [hide["b"], hide["l"], hide["r"]]
        box.children = [pr, inner]
    elif notation == "bottom":
        pr.children = [hide["t"], hide["l"], hide["r"]]
        box.children = [pr, inner]
    elif notation == "updiagonalstrike":
        pr.children = [hide["t"], hide["b"], hide["l"], hide["r"], _hide("m:strikeBLTR")]
        box.children = [pr, inner]
    elif notation == "downdiagonalstrike":
        pr.children = [hide["t"], hide["b"], hide["l"], hide["r"], _hide("m:strikeTLBR")]
        box.children = [pr, inner]
    elif notation == "verticalstrike":
        pr.children = [hide["t"], hide["b"], hide["l"], hide["r"], _hide("m:strikeV")]
        box.children = [pr, inner]
    elif notation == "horizontalstrike":
        pr.children = [hide["t"], hide["b"], hide["l"], hide["r"], _hide("m:strikeH")]
        box.children = [pr, inner]
    else:
        pr.children = [hide["t"], hide["b"], hide["l"], hide["r"]]
        box.children = [pr, inner]
    target_parent.children.append(box)
    return inner


def _handle_mmultiscripts(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node | None:
    if not element.children:
        return None
    base = element.children[0]
    post_subs: list[Node] = []
    post_supers: list[Node] = []
    pre_subs: list[Node] = []
    pre_supers: list[Node] = []
    divider = False
    for index, child in enumerate(element.children[1:]):
        if child.name == "mprescripts":
            divider = True
        elif child.name != "none":
            if index % 2:
                (pre_subs if divider else post_supers).append(child)
            else:
                (pre_supers if divider else post_subs).append(child)
    next_ancestors = [element, *ancestors]
    temp = Node()
    _walk(base, temp, None, None, next_ancestors)
    top = temp.children[0] if temp.children else _tag("m:r")
    if post_subs or post_supers:
        sub_target = _tag("m:sub")
        for item in post_subs:
            _walk(item, sub_target, None, None, next_ancestors)
        sup_target = _tag("m:sup")
        for item in post_supers:
            _walk(item, sup_target, None, None, next_ancestors)
        wrapped = _tag("", children=[_tag("m:e", children=[top])])
        if post_subs and post_supers:
            wrapped.name = "m:sSubSup"
            wrapped.children.extend([sub_target, sup_target])
        elif post_subs:
            wrapped.name = "m:sSub"
            wrapped.children.append(sub_target)
        else:
            wrapped.name = "m:sSup"
            wrapped.children.append(sup_target)
        top = wrapped
    if pre_subs or pre_supers:
        pre_sub = _tag("m:sub")
        for item in pre_subs:
            _walk(item, pre_sub, None, None, next_ancestors)
        pre_sup = _tag("m:sup")
        for item in pre_supers:
            _walk(item, pre_sup, None, None, next_ancestors)
        top = _tag("m:sPre", children=[_tag("m:e", children=[top]), pre_sub, pre_sup])
    target_parent.children.append(top)
    return None


def _handle_mglyph(element: Node, target_parent: Node, *_args: Any) -> None:
    alt = element.attribs.get("alt")
    if alt:
        dest = _ensure_math_text_target(target_parent)
        dest.children.append(Node(type="text", data=alt))
    return None


def _handle_mspace(element: Node, target_parent: Node, *_args: Any) -> None:
    if element.attribs.get("linebreak") == "newline":
        target_parent.pending_brk = True
        return None
    target_parent.children.append(
        _tag(
            "m:r",
            children=[_tag("m:t", {"xml:space": "preserve"}, [Node(type="text", data=" ")])],
        )
    )
    return None


def _handle_text(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node:
    text = re.sub(r"[\u2061-\u2064\u200B]", "", element.data)
    if any(a.name in {"mi", "mn", "mo"} for a in ancestors):
        text = re.sub(r"\s", "", text)
    else:
        ms = next((a for a in ancestors if a.name == "ms"), None)
        if ms is not None:
            text = (ms.attribs.get("lquote") or '"') + text + (ms.attribs.get("rquote") or '"')
    if text:
        dest = _ensure_math_text_target(target_parent)
        if dest.children and dest.children[-1].type == "text":
            dest.children[-1].data += text
        else:
            dest.children.append(Node(type="text", data=text))
        return dest
    return target_parent


def _text_container(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    ancestors: list[Node],
    text_type: str,
) -> Node:
    if previous_sibling is not None and previous_sibling.is_nary:
        e = _nary_base_arg(target_parent)
        if e is not None:
            target_parent = e
    has_mglyph = any(child.name == "mglyph" for child in element.children)
    style = _style_of(
        element, ancestors, previous_sibling.style if previous_sibling else None
    )
    element.style = style
    element.has_mglyph_child = has_mglyph
    prev_style = previous_sibling.style if previous_sibling else None
    style_same = bool(prev_style) and all(
        style[key] == prev_style.get(key) for key in style
    ) and (previous_sibling.has_mglyph_child if previous_sibling else None) == has_mglyph
    same_group = previous_sibling is not None and (
        text_type == previous_sibling.name
        or (
            text_type in {"mi", "mn", "mo"}
            and previous_sibling.name in {"mi", "mn", "mo"}
        )
    )
    last = target_parent.children[-1] if target_parent.children else None
    if (
        same_group
        and style_same
        and not has_mglyph
        and not target_parent.pending_brk
        and last is not None
        and last.name == "m:r"
        and last.children
    ):
        return last.children[-1]
    r_element = _tag("m:r")
    variant = style.get("variant") or ""
    if variant:
        mr_pr = _tag("m:rPr")
        style_value = _STYLES.get(variant)
        if style_value:
            mr_pr.children.append(_tag("m:sty", {"m:val": style_value}))
        else:
            mr_pr.children.append(_tag("m:nor"))
        r_element.children.append(mr_pr)
        wr_pr = _tag("w:rPr")
        if "bold" in variant:
            wr_pr.children.append(_tag("w:b"))
        if "italic" in variant:
            wr_pr.children.append(_tag("w:i"))
        if wr_pr.children:
            r_element.children.append(wr_pr)
    elif has_mglyph or text_type == "mtext":
        r_element.children.append(_tag("m:rPr", children=[_tag("m:nor")]))
    elif style.get("fontstyle") == "normal" or (
        text_type == "ms" and style.get("fontstyle") == ""
    ):
        r_element.children.append(_tag("m:rPr", children=[_tag("m:sty", {"m:val": "p"})]))
    target = _tag("m:t", {"xml:space": "preserve"})
    if not element.children:
        target.children.append(Node(type="text", data="\u200b"))
    _consume_pending_brk(target_parent, r_element)
    r_element.children.append(target)
    target_parent.children.append(r_element)
    return target


def _handle_mtext(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node:
    return _text_container(element, target_parent, previous_sibling, ancestors, "mtext")


def _handle_mi(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node:
    return _text_container(element, target_parent, previous_sibling, ancestors, "mi")


def _handle_mn(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node:
    return _text_container(element, target_parent, previous_sibling, ancestors, "mn")


def _handle_mo(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node:
    return _text_container(element, target_parent, previous_sibling, ancestors, "mo")


def _handle_ms(
    element: Node,
    target_parent: Node,
    previous_sibling: Node | None,
    next_sibling: Node | None,
    ancestors: list[Node],
) -> Node:
    return _text_container(element, target_parent, previous_sibling, ancestors, "ms")


_HANDLERS.update(
    {
        "math": _handle_math,
        "semantics": _handle_passthrough,
        "mstyle": _handle_passthrough,
        "mrow": _handle_mrow,
        "mfrac": _handle_mfrac,
        "msqrt": _handle_msqrt,
        "mroot": _handle_mroot,
        "msub": _handle_msub,
        "msup": _handle_msup,
        "msubsup": _handle_msubsup,
        "munder": _handle_munder,
        "mover": _handle_mover,
        "munderover": _handle_munderover,
        "mtable": _handle_mtable,
        "mtr": _handle_mtr,
        "mtd": _handle_mtd,
        "menclose": _handle_menclose,
        "mmultiscripts": _handle_mmultiscripts,
        "mglyph": _handle_mglyph,
        "mspace": _handle_mspace,
        "mtext": _handle_mtext,
        "mi": _handle_mi,
        "mn": _handle_mn,
        "mo": _handle_mo,
        "ms": _handle_ms,
        "text": _handle_text,
    }
)
