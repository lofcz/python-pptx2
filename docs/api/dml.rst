.. _dml_api:

DrawingML objects
=================

Low-level drawing elements like fill and color that appear repeatedly in
various aspects of shapes.


|ChartFormat| objects
---------------------

.. autoclass:: pptx2.dml.chtfmt.ChartFormat
   :members:


|FillFormat| objects
--------------------

.. autoclass:: pptx2.dml.fill.FillFormat
   :members:
   :exclude-members: from_fill_parent
   :undoc-members:


|LineFormat| objects
--------------------

.. autoclass:: pptx2.dml.line.LineFormat
   :members:
   :undoc-members:


|LineFormat| line ends
~~~~~~~~~~~~~~~~~~~~~~

.. autoclass:: pptx2.dml.line.LineEndFormat
   :members:
   :undoc-members:


|ColorFormat| objects
---------------------

.. autoclass:: pptx2.dml.color.ColorFormat
   :members: brightness, rgb, theme_color, type, alpha
   :undoc-members:


|RGBColor| objects
------------------

.. autoclass:: pptx2.dml.color.RGBColor
   :members: from_string, from_hex
   :undoc-members:


Effect proxies
--------------

|ShadowFormat| objects
~~~~~~~~~~~~~~~~~~~~~~

.. autoclass:: pptx2.dml.effect.ShadowFormat
   :members:
   :undoc-members:


.. autoclass:: pptx2.dml.effect.GlowFormat
   :members:
   :undoc-members:


.. autoclass:: pptx2.dml.effect.SoftEdgeFormat
   :members:
   :undoc-members:


.. autoclass:: pptx2.dml.effect.BlurFormat
   :members:
   :undoc-members:


.. autoclass:: pptx2.dml.effect.ReflectionFormat
   :members:
   :undoc-members:


Picture effects
---------------

.. autoclass:: pptx2.dml.picture.PictureEffects
   :members:
   :undoc-members:
