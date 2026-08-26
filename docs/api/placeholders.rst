.. _placeholder_api:

Placeholders
============

The following classes represent placeholder shapes. A placeholder most
commonly appears on a slide, but also appears on a slide layout and a slide
master. The role of a master placeholder and layout placeholder differs from
that of a slide placeholder and these roles are reflected in the distinct
classes for each.

There are a larger variety of slide placeholders to accomodate their more
complex and varied behaviors.


|MasterPlaceholder| objects
---------------------------

.. autoclass:: pptx2.shapes.placeholder.MasterPlaceholder()
   :members:
   :exclude-members:
      adjustments, get_or_add_ln, has_chart, has_table, idx, ln, orient,
      part, ph_type, shape_type, sz
   :inherited-members:
   :undoc-members:


|LayoutPlaceholder| objects
---------------------------

.. autoclass:: pptx2.shapes.placeholder.LayoutPlaceholder()
   :members:
   :undoc-members:


ChartPlaceholder objects
------------------------

.. autoclass:: pptx2.shapes.placeholder.ChartPlaceholder()
   :members:
   :exclude-members:
      has_chart, has_table, has_text_frame, part
   :inherited-members:
   :undoc-members:


PicturePlaceholder objects
--------------------------

.. autoclass:: pptx2.shapes.placeholder.PicturePlaceholder()
   :members:
   :exclude-members:
      has_chart, has_table, has_text_frame, part
   :inherited-members:
   :undoc-members:


TablePlaceholder objects
------------------------

.. autoclass:: pptx2.shapes.placeholder.TablePlaceholder()
   :members:
   :inherited-members:
   :exclude-members:
      has_chart, has_table, has_text_frame, is_placeholder,
      part
   :undoc-members:


PlaceholderGraphicFrame objects
-------------------------------

.. autoclass:: pptx2.shapes.placeholder.PlaceholderGraphicFrame()
   :members:
   :inherited-members:
   :exclude-members:
      chart_part, has_text_frame, is_placeholder, part
   :undoc-members:


PlaceholderPicture objects
--------------------------

.. autoclass:: pptx2.shapes.placeholder.PlaceholderPicture()
   :members:
   :inherited-members:
   :exclude-members:
      get_or_add_ln, has_chart, has_table, has_text_frame, ln,
      part
   :undoc-members:


_PlaceholderFormat objects
--------------------------

.. autoclass:: pptx2.shapes.base._PlaceholderFormat()
   :members:
   :inherited-members:
   :undoc-members:
