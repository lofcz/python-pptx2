"""lxml custom element classes for custom properties-related XML elements."""

from __future__ import annotations

from typing import Any, Callable, cast

from lxml.etree import _Element  # pyright: ignore[reportPrivateUsage]

from pptx2.oxml import parse_xml
from pptx2.oxml.ns import nsdecls, qn
from pptx2.oxml.xmlchemy import BaseOxmlElement

#: Format id required on every ``property`` element. This is the well-known
#: "User-Defined Property" GUID (the same one Office writes).
_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

#: Python type -> local name of the ``vt:`` child element that carries it.
_VT_TAG_FOR_TYPE: dict[type, str] = {bool: "bool", int: "i4", float: "r8", str: "lpstr"}

#: Local name of a ``vt:`` child element -> callable that produces the Python
#: value from its text. Types we never write are still parsed leniently so
#: real-world files open without surprises.
_VT_PARSER_FOR_TAG: dict[str, Callable[[str], Any]] = {
    "lpstr": str,
    "lpwstr": str,
    "bstr": str,
    "i1": int,
    "i2": int,
    "i4": int,
    "i8": int,
    "int": int,
    "ui1": int,
    "ui2": int,
    "ui4": int,
    "ui8": int,
    "uint": int,
    "r4": float,
    "r8": float,
    "decimal": float,
    "bool": lambda text: text.strip() in ("true", "1"),
}

_I4_MIN, _I4_MAX = -(2**31), 2**31 - 1


def _coerce_text(value: Any, vt_tag: str) -> str:
    """Return the lexical form of `value` for a `vt:` element of `vt_tag`."""
    if vt_tag == "bool":
        return "true" if value else "false"
    if vt_tag == "r8":
        return repr(float(value))
    return str(value)


class CT_CustomProperties(BaseOxmlElement):
    """`Properties` element.

    The root element of the Custom Properties part stored as `/docProps/custom.xml`. Each child
    `property` element carries a name plus a single typed `vt:` value child (`vt:lpstr` for str,
    `vt:i4` for int, `vt:r8` for float, and `vt:bool` for bool).
    """

    @staticmethod
    def new() -> CT_CustomProperties:
        """Return a new empty `Properties` element."""
        return cast(
            CT_CustomProperties,
            parse_xml(
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
                'custom-properties" %s/>\n' % nsdecls("vt")
            ),
        )

    def get_property(self, name: str, default: Any = None) -> Any:
        """Return typed value of property `name`, or `default` if no such property."""
        prop = self.property_by_name(name)
        if prop is None:
            return default
        return self._value_of_prop(prop)

    @property
    def property_elms(self) -> list[_Element]:
        """List of `property` child elements in document order."""
        return self.findall(qn("custom:property"))

    def property_names(self) -> list[str]:
        """List the names of the `property` child elements in document order."""
        return [prop.get("name") for prop in self.property_elms]

    def property_by_name(self, name: str) -> _Element | None:
        """Return the `property` element named `name`, or None if not present."""
        for prop in self.property_elms:
            if prop.get("name") == name:
                return prop
        return None

    def remove_property(self, name: str) -> bool:
        """Remove property `name`; return True if a property was removed."""
        prop = self.property_by_name(name)
        if prop is None:
            return False
        self.remove(prop)
        return True

    def set_property(self, name: str, value: Any) -> None:
        """Set property `name` to `value`, adding it if not already present.

        `value` must be a str, int, float, or bool (checked in this order, so a bool is not
        mistaken for an int). An existing property of a different type is retyped.
        """
        # -- bool must be tested before int: bool is a subclass of int --
        for py_type in (bool, int, float, str):
            if isinstance(value, py_type):
                vt_tag = _VT_TAG_FOR_TYPE[py_type]
                break
        else:
            tmpl = (
                "custom property value must be a str, int, float, or bool, got %s "
                "(%r); datetime and other VT types are not supported"
            )
            raise TypeError(tmpl % (type(value).__name__, value))
        if not isinstance(name, str) or not name:
            raise ValueError("custom property name must be a non-empty string, got %r" % (name,))
        if vt_tag == "i4" and not _I4_MIN <= value <= _I4_MAX:
            raise ValueError(
                "int custom property value must fit in 32 bits (vt:i4), got %r" % (value,)
            )
        prop = self.property_by_name(name)
        if prop is None:
            prop = self._add_property(name)
        for child in list(prop):
            prop.remove(child)
        vt = prop.makeelement(qn("vt:%s" % vt_tag), {})
        prop.append(vt)
        vt.text = _coerce_text(value, vt_tag)

    def _add_property(self, name: str) -> _Element:
        """Append and return a new `property` element for `name`."""
        prop = self.makeelement(qn("custom:property"), {})
        self.append(prop)
        prop.set("fmtid", _FMTID)
        prop.set("pid", str(self._next_pid))
        prop.set("name", name)
        return prop

    @property
    def _next_pid(self) -> int:
        """One greater than the highest pid in use (2 for an empty element).

        The count is based on the maximum rather than the number of properties so pids are never
        reused after a deletion. Office reserves pid 1, so numbering starts at 2.
        """
        pids = [1]
        for prop in self.property_elms:
            try:
                pids.append(int(prop.get("pid")))  # pyright: ignore[reportArgumentType]
            except (TypeError, ValueError):
                continue
        return max(pids) + 1

    def _value_of_prop(self, prop: _Element) -> Any:
        """Return the Python value carried by `prop`'s `vt:` child element."""
        if len(prop) == 0:
            return None
        vt = prop[0]
        vt_tag = vt.tag.rsplit("}", 1)[-1]
        parser = _VT_PARSER_FOR_TAG.get(vt_tag, str)
        if vt.text is None:
            return None
        return parser(vt.text)
