"""High-level composition entry points.

This package collects the cross-presentation operations: JSON-driven
authoring (:func:`from_spec`), validated slide import across decks
(:func:`import_slide`, :func:`append_deck`), and bulk template re-pointing
(:func:`apply_template`).

Two import engines live here:

* :mod:`pptx2.compose.deck_compose` — the paper-pptx validating engine.
  ``import_slide(dest_prs, source_prs, slide, mode=...)`` validates the
  complete operation up front and returns an |ImportReport|; a refusal
  leaves the destination byte-for-byte unchanged. This is the engine behind
  ``Presentation.import_slide``.
* ``pptx2._slide_importer`` — the older low-level part-level engine
  (``import_slide(src_slide_part, dst_prs_part, merge_master=...)``
  returning the new |Slide|), kept as a private implementation detail.

The paper-pptx API is the public surface re-exported here::

    from pptx2.compose import from_spec, import_slide, append_deck, apply_template
"""

from __future__ import annotations

from pptx2._template_applier import apply_template
from pptx2.compose.deck_compose import ImportReport, append_deck, import_slide
from pptx2.compose.from_spec import from_spec, from_yaml

__all__ = [
    "ImportReport",
    "append_deck",
    "apply_template",
    "from_spec",
    "from_yaml",
    "import_slide",
]
