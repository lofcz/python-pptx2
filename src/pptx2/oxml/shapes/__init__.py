"""Base shape-related objects such as BaseShape."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import TypeAlias

    from pptx2.oxml.shapes.autoshape import CT_Shape
    from pptx2.oxml.shapes.connector import CT_Connector
    from pptx2.oxml.shapes.graphfrm import CT_GraphicalObjectFrame
    from pptx2.oxml.shapes.groupshape import CT_GroupShape
    from pptx2.oxml.shapes.picture import CT_Picture


ShapeElement: TypeAlias = (
    "CT_Connector | CT_GraphicalObjectFrame |  CT_GroupShape | CT_Picture | CT_Shape"
)
