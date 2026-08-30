"""Custom properties part, corresponds to ``/docProps/custom.xml`` part in package."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Iterator

from pptx2.opc.constants import CONTENT_TYPE as CT
from pptx2.opc.constants import RELATIONSHIP_TYPE as RT
from pptx2.opc.package import XmlPart
from pptx2.opc.packuri import PackURI
from pptx2.oxml.customprops import CT_CustomProperties

if TYPE_CHECKING:
    from pptx2.package import Package


class CustomPropertiesPart(XmlPart):
    """Corresponds to part named `/docProps/custom.xml`.

    Contains the user-defined (custom) document properties for this document package and supports
    mapping-style access (`part[name]`) to them.
    """

    _element: CT_CustomProperties

    @classmethod
    def default(cls, package: Package) -> CustomPropertiesPart:
        """Return a new, empty |CustomPropertiesPart| for `package`.

        Provides a base for adding custom properties to a package that doesn't yet have any.
        """
        return cls._new(package)

    def __contains__(self, name: str) -> bool:
        return name in self._element.property_names()

    def __delitem__(self, name: str) -> None:
        if not self._element.remove_property(name):
            raise KeyError(name)

    def __getitem__(self, name: str) -> Any:
        names = self._element.property_names()
        if name not in names:
            raise KeyError(name)
        return self._element.get_property(name)

    def __iter__(self) -> Iterator[str]:
        return iter(self._element.property_names())

    def __len__(self) -> int:
        return len(self._element.property_elms)

    def __setitem__(self, name: str, value: Any) -> None:
        self._element.set_property(name, value)

    def keys(self) -> list[str]:
        """List the names of the custom properties in document order."""
        return self._element.property_names()

    @classmethod
    def _new(cls, package: Package) -> CustomPropertiesPart:
        """Return a new empty |CustomPropertiesPart| instance."""
        return CustomPropertiesPart(
            PackURI("/docProps/custom.xml"),
            CT.OFC_CUSTOM_PROPERTIES,
            package,
            CT_CustomProperties.new(),
        )


class CustomProperties(object):
    """Mapping-style access to the custom document properties of a package.

    A property name maps to a str, int, float, or bool value, stored in ``/docProps/custom.xml``
    as ``vt:lpstr``, ``vt:i4``, ``vt:r8``, and ``vt:bool`` respectively. Property names are
    stored verbatim (no character restrictions are imposed). Reading from a package that has no
    custom-properties part behaves like reading an empty mapping; the part (and its relationship
    from the package root) is created lazily by the first assignment.
    """

    def __init__(self, package: Package) -> None:
        super(CustomProperties, self).__init__()
        self._package = package

    def __contains__(self, name: str) -> bool:
        part = self._part()
        return part is not None and name in part

    def __delitem__(self, name: str) -> None:
        part = self._part()
        if part is None:
            raise KeyError(name)
        del part[name]

    def __getitem__(self, name: str) -> Any:
        part = self._part()
        if part is None:
            raise KeyError(name)
        return part[name]

    def __iter__(self) -> Iterator[str]:
        part = self._part()
        return iter(()) if part is None else iter(part)

    def __len__(self) -> int:
        part = self._part()
        return 0 if part is None else len(part)

    def __setitem__(self, name: str, value: Any) -> None:
        self._get_or_add_part()[name] = value

    def keys(self) -> list[str]:
        """List the names of the custom properties in document order."""
        return list(self)

    def _get_or_add_part(self) -> CustomPropertiesPart:
        """Return the custom-properties part, creating and relating it if need be."""
        part = self._part()
        if part is None:
            part = CustomPropertiesPart.default(self._package)
            self._package.relate_to(part, RT.CUSTOM_PROPERTIES)
        return part

    def _part(self) -> CustomPropertiesPart | None:
        """Return the custom-properties part of the package, or None if there is none."""
        try:
            return self._package.part_related_by(RT.CUSTOM_PROPERTIES)
        except KeyError:
            return None
