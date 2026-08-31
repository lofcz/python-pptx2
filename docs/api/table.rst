.. _table_api:

Table-related objects
=====================


.. currentmodule:: pptx2.table


|Table| objects
----------------

A |Table| object is added to a slide using the
:meth:`~.SlideShapes.add_table` method on |SlideShapes|.

Table content and shape geometry have different owners. Group-aware
:meth:`~pptx2.shapes.shapetree.SlideShapes.table_by_name` returns a |Table| directly and is the
shortest path for cell, row, and column work. When the workflow also needs position, width, height,
rotation, or other shape properties, use
:meth:`~pptx2.shapes.shapetree.SlideShapes.shape_by_name`, verify the returned graphic frame has a
table, and access its :attr:`~pptx2.shapes.graphfrm.GraphicFrame.table`. Geometry stays on that
graphic frame; |Table| does not provide owner-navigation or geometry aliases.

.. autoclass:: Table()
   :members:
   :inherited-members:
   :exclude-members:
      notify_height_changed, notify_width_changed, part
   :undoc-members:


Direct-format column templates
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

:meth:`Table.insert_column` can copy direct cell formatting from an existing column. The source
index always names a column in the table *before* insertion. Each inserted cell receives the
complete direct cell-properties subtree from the source cell in the same row, including fills,
borders, margins, anchors, and producer extension content. Text and merge state are not copied;
the inserted cells are empty and unmerged.

This is a direct-format copy, not an effective-format calculation. Appearance supplied only by the
table style or theme is not materialized into the new cells. If the caller supplies a width, it
wins. Otherwise a selected source column supplies the new width; without a source, the existing
neighbor-width behavior applies.


|_ColumnCollection| objects
---------------------------

.. autoclass:: _ColumnCollection()
   :members:
   :member-order: bysource
   :undoc-members:


|_Column| objects
-----------------

.. autoclass:: _Column()
   :members:
   :member-order: bysource
   :undoc-members:


|_RowCollection| objects
------------------------

.. autoclass:: _RowCollection()
   :members:
   :member-order: bysource
   :undoc-members:


|_Row| objects
--------------

.. autoclass:: _Row()
   :members:
   :member-order: bysource
   :undoc-members:


|_Cell| objects
---------------

A |_Cell| object represents a single table cell at a particular row/column
location in the table. |_Cell| objects are not constructed directly. A
reference to a |_Cell| object is obtained using the :meth:`Table.cell` method,
specifying the cell's row/column location. A cell object can also be obtained
using the :attr:`_Row.cells` collection.

.. autoclass:: _Cell
   :members:
   :member-order: bysource
   :undoc-members:
