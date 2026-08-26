.. _slides_api:

Slides
======

|Slides| objects
-----------------

The |Slides| object is accessed using the
:attr:`~pptx2.presentation.Presentation.slides` property of |Presentation|. It
is not intended to be constructed directly.

.. autoclass:: pptx2.slide.Slides()
   :members:
   :member-order: bysource
   :undoc-members:


|Slide| objects
---------------

An individual |Slide| object is accessed by index from |Slides| or as the
return value of :meth:`add_slide`.

.. autoclass:: pptx2.slide.Slide()
   :members:
   :exclude-members: part
   :inherited-members:
   :undoc-members:


|SlideLayouts| objects
----------------------

The |SlideLayouts| object is accessed using the
:attr:`~pptx2.slide.SlideMaster.slide_layouts` property of |SlideMaster|, typically::

    >>> from pptx2 import Presentation
    >>> prs = Presentation()
    >>> slide_layouts = prs.slide_master.slide_layouts

As a convenience, since most presentations have only a single slide master, the
|SlideLayouts| collection for the first master may be accessed directly from the
|Presentation| object::

    >>> slide_layouts = prs.slide_layouts

This class is not intended to be constructed directly.

.. autoclass:: pptx2.slide.SlideLayouts()
   :members:
   :exclude-members: element, parent
   :inherited-members:
   :undoc-members:


|SlideLayout| objects
---------------------

.. autoclass:: pptx2.slide.SlideLayout
   :members:
   :exclude-members: iter_cloneable_placeholders


|SlideMasters| objects
----------------------

The |SlideMasters| object is accessed using the
:attr:`~pptx2.presentation.slide_masters` property of |Presentation|, typically::

    >>> from pptx2 import Presentation
    >>> prs = Presentation()
    >>> slide_masters = prs.slide_masters

As a convenience, since most presentations have only a single slide master, the
first master may be accessed directly from the |Presentation| object without indexing
the collection::

    >>> slide_master = prs.slide_master

This class is not intended to be constructed directly.

.. autoclass:: pptx2.slide.SlideMasters()
   :members:
   :exclude-members: element, parent
   :inherited-members:
   :undoc-members:

|SlideMaster| objects
---------------------

.. autoclass:: pptx2.slide.SlideMaster
   :members:
   :exclude-members: related_slide_layout, sldLayoutIdLst


|SlidePlaceholders| objects
---------------------------

.. autoclass:: pptx2.shapes.shapetree.SlidePlaceholders
   :members:


|NotesSlide| objects
--------------------

.. autoclass:: pptx2.slide.NotesSlide
   :members:
   :exclude-members: clone_master_placeholders
   :inherited-members:


|SlideTransition| objects
-------------------------

The :attr:`~pptx2.slide.Slide.transition` property of |Slide| returns a
|SlideTransition| proxy backed by ``<p:transition>``. Reads on an unset
transition return |None| and never mutate XML, keeping theme inheritance
intact.

.. autoclass:: pptx2.slide.SlideTransition
   :members:
   :undoc-members:
