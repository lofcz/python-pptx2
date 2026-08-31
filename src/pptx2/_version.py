"""Paper-pptx provenance version.

paper-pptx is the upstream this feature set was ported from; ``__paper_version__``
records which paper-pptx release the port corresponds to. python-pptx2 keeps its
own ``__version__`` in ``pptx2/__init__.py``.
"""

from __future__ import annotations

__paper_version__ = "0.1.3"


def assert_distribution_identity() -> None:
    """No-op in python-pptx2.

    paper-pptx uses this to fail early when both python-pptx and paper-pptx are
    installed and both claim the ``pptx`` import name. python-pptx2 imports as
    ``pptx2`` and cannot collide with either distribution, so the guard has
    nothing to check here; it is retained so ported call sites keep working.
    """
