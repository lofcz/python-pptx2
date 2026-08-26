python-pptx2
=============

|PyPI| |PyPI - Python Versions| |CI| |License| |Documentation|

.. |PyPI| image:: https://img.shields.io/pypi/v/python-pptx2.svg
   :target: https://pypi.org/project/python-pptx2/
   :alt: PyPI

.. |PyPI - Python Versions| image:: https://img.shields.io/pypi/pyversions/python-pptx2.svg
   :target: https://pypi.org/project/python-pptx2/
   :alt: Python 3.9 – 3.13

.. |CI| image:: https://github.com/lofcz/python-pptx2/actions/workflows/ci.yml/badge.svg
   :target: https://github.com/lofcz/python-pptx2/actions/workflows/ci.yml
   :alt: CI status

.. |License| image:: https://img.shields.io/badge/license-MIT-blue.svg
   :target: https://github.com/lofcz/python-pptx2/blob/master/LICENSE
   :alt: MIT License

.. |Documentation| image:: https://img.shields.io/badge/docs-github.com%2Flofcz%2Fpython--pptx2-blue.svg
   :target: https://github.com/lofcz/python-pptx2
   :alt: Documentation

**PowerPoint decks from Python, that actually fit.**

*python-pptx2* is a fork of `power-pptx`_ (Daniel Halwell) and, through
it, of `python-pptx`_ by `Steve Canny`_. It is a Python library for
creating, reading, and updating PowerPoint (.pptx) files, plus native
LaTeX equations.

+----------------+----------------------------------------------------------+
| **Package**    | ``python-pptx2`` on PyPI (imports as ``pptx2``)          |
+----------------+----------------------------------------------------------+
| **Python**     | 3.9 – 3.13                                               |
+----------------+----------------------------------------------------------+
| **License**    | MIT                                                      |
+----------------+----------------------------------------------------------+
| **Source**     | https://github.com/lofcz/python-pptx2                    |
+----------------+----------------------------------------------------------+

.. _`power-pptx`: https://github.com/CodeHalwell/power-pptx

A typical use is generating a PowerPoint presentation from dynamic
content such as a database query, an analytics output, an LLM payload,
or a JSON spec — and downloading the generated .pptx file. It runs on
any Python-capable platform (macOS, Linux, Windows) and does not
require Microsoft PowerPoint to be installed or licensed.

**Why this fork exists: space-aware authoring.**  The headline
proposition is that text doesn't overflow its container and shapes
don't slide off the edges of the slide.  Three layered tools used
together catch ~all real-world layout issues:

1. ``TextFrame.fit_text(...)`` measures with Pillow font metrics and
   bakes a fitting size into the XML *before* save.
2. ``text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`` lets
   PowerPoint shrink at render time as a fallback.
3. ``slide.lint()`` catches what slipped through; ``auto_fix()`` (or
   the one-call ``slide.tidy()``) nudges off-slide shapes back inside.

Reach for python-pptx2 whenever the deck is generated dynamically and
has to look right without manual cleanup.

Installation
------------

Install from PyPI::

    pip install python-pptx2

Then in Python::

    from pptx2 import Presentation

``python-pptx2`` imports as ``pptx2`` so it can sit beside upstream
``python-pptx`` (``pptx``) and the parent fork ``power-pptx``
(``power_pptx``). To migrate from those packages, replace
``from pptx import`` / ``from power_pptx import`` with
``from pptx2 import``.

Optional dependencies:

* ``cairosvg`` — install only if you want ``add_svg_picture(...)`` to
  auto-rasterise the PNG fallback.
* ``pyyaml`` — install only if you want ``DesignTokens.from_yaml``.
* ``python-pptx2[math]`` (``latex2mathml``) — required
  for ``slide.shapes.add_equation(...)`` / ``paragraph.add_math(...)``.
  MathML → OMML is bundled.
* ``soffice`` (LibreOffice) on PATH — required for
  ``Presentation.render_thumbnails()``.
* ``pdftoppm`` (Poppler) or ``pypdfium2`` — required for the
  PDF→PNG split path in the thumbnail renderer.

Claude Code skill
-----------------

python-pptx2 ships a Claude Code skill alongside the library — pip-install
the package and the skill files are already on disk inside it.  Install
the skill into your local Claude Code skills directory with::

    python -m pptx2.skill install

This copies ``SKILL.md`` and the ``references/`` directory into
``~/.claude/skills/python-pptx2/``.  Claude Code (and any compatible
Claude Agent SDK harness) will pick it up automatically the next time
it starts.

Other useful commands::

    # Print the skill source path inside the installed package
    python -m pptx2.skill path

    # Install into a custom directory
    python -m pptx2.skill install --target /path/to/skills/python-pptx2

    # Refuse to overwrite an existing install
    python -m pptx2.skill install --no-overwrite

The skill documents the headline space-aware-authoring workflow, the
``BBox`` value object, the one-call ``add_text`` / ``add_arrow``
helpers, the diagram recipes (``horizontal_pipeline``,
``hub_and_spoke``, ``cycle``, ``decision_tree``,
``comparison_columns``), and 16 focused reference docs covering
effects, animations, transitions, theming, charts, tables, 3D,
SmartArt, and rendering.  It also includes an *anti-patterns* section
calling out the mistakes LLMs commonly make (mis-comparing wrapper
objects, assuming ``add_connector`` puts an arrowhead on the line,
sizing a diagram to a broken picture's bbox rather than its enclosing
card, …).

Quick start
-----------

A minimal end-to-end deck-generation pattern::

    from pptx2 import Presentation, BBox, audit
    from pptx2.diagrams import horizontal_pipeline
    from pptx2.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Pipeline overview"

    horizontal_pipeline(
        slide,
        BBox.from_inches(0.5, 2.5, 9, 2.2),
        steps=["Extract", "Classify", "Enrich", "Output"],
        accent="#0B5CFF",
    )

    slide.shapes.add_text(
        BBox.from_inches(0.5, 5.5, 9, 1),
        text="Four-stage data pipeline.",
        align="center", size_pt=14, color="#666666",
    )

    # Lint + safe auto-fixes
    slide.tidy()

    # Optional: full-deck audit
    print(audit(prs).markdown())

    prs.save("out.pptx")

What's new in the fork
----------------------

The fork extends the 1.0.2 surface with features the upstream roadmap
did not cover.  All additions are drop-in compatible — existing
scripts keep working — and every new feature ships with a round-trip
regression test.

* **Space-aware authoring** — ``TextFrame.fit_text`` bakes a fitting
  size before save; ``auto_size`` flags shrink at render time; the
  layout linter reports text overflow / off-slide shapes /
  collisions.  Three-tier safety so generated decks look right
  without manual cleanup.

* **Geometry and convenience helpers (v2.8)** — first-class
  ``BBox`` value object with ``inset`` / ``split_h`` / ``split_v`` /
  ``grid`` / ``contains`` / ``intersection``.  One-call
  ``slide.shapes.add_text(bb, text=..., color="#0B5CFF",
  align="center")`` collapses the historical seven-line styling
  ritual.  Hex-string shortcuts (``shape.fill_hex("#0B5CFF")``,
  ``shape.line_hex(...)``).  ``shape.set_text_preserving_format(new)``
  for templated placeholders.

* **Real arrows** —
  ``slide.shapes.add_arrow(start=a, end=b, head="triangle",
  color="#0B5CFF", inset_pt=6)`` produces a connector with an
  arrowhead and auto-routed endpoints (mid-edge of target shape,
  pulled back by ``inset_pt``).  No XML required.

* **Diagram recipes** — ``horizontal_pipeline``, ``vertical_pipeline``,
  ``hub_and_spoke``, ``cycle``, ``decision_tree``,
  ``comparison_columns`` from ``pptx2.diagrams`` cover ~80% of
  architecture-deck patterns.  Each takes a slide, a ``BBox``, and a
  content spec.

* **Picture and slide helpers** — ``picture.replace_with(builder,
  padding=...)`` deletes a broken / sub-quality picture and calls a
  builder in its place.  ``picture.enclosing_container()`` finds the
  surrounding card so the rebuild fills the right area.
  ``slide.tidy()`` is the one-call lint + safe auto-fix.
  ``slide.find_empty_region(...)`` returns an unused area for
  greenfield placement.

* **Whole-deck audit** — ``pptx2.audit(prs)`` returns an
  ``AuditReport`` with lint issues, broken pictures, empty slides,
  uncommon-font warnings, and oversized-picture warnings.  Markdown
  output for chat replies.

* **Visual effects** — outer shadow, glow, soft edges, blur, and
  reflection exposed as non-mutating proxies on every shape;
  alpha-tinted colours (``RGBColor.alpha``); gradient fills with
  ``linear`` / ``radial`` / ``rectangular`` / ``shape`` kinds and
  mutable stops; line ends, caps, joins, and compound lines.

* **Animations and transitions** — preset entrance / exit / emphasis
  effects; motion-path presets; per-paragraph reveal; sequencing
  context manager; per-slide and deck-wide transitions including
  Morph and the other ``p14:`` extension transitions.

* **JSON authoring** — ``pptx2.compose.from_spec(...)`` builds a
  deck from a JSON-shaped spec; ``import_slide`` and
  ``apply_template`` cover cross-presentation operations.

* **Theme reader and writer** — read theme colours and fonts; write
  fresh ``<a:srgbClr>`` values into the clrScheme; apply a theme
  imported from a ``.potx``.

* **Picture effects** — transparency, brightness, contrast, recolor
  (grayscale, sepia, washout, duotone); native SVG embedding with
  PNG fallback.

* **Design-system layer** — ``DesignTokens`` (palette, typography,
  shadows, radii, spacings) loadable from a dict, YAML, or a
  ``.pptx``; a token-resolving ``shape.style`` facade; ``Grid`` /
  ``Stack`` layout primitives; opinionated slide recipes
  (``title``, ``bullet``, ``kpi``, ``quote``, ``image_hero``); a
  starter pack of three example token sets.

* **Shape-level building blocks** — ``add_kpi_card``,
  ``add_progress_bar``, ``add_gauge``, ``add_status_pill``,
  ``add_stat_strip``, ``add_article_card``: token-driven cards that
  return small dataclasses exposing the constituent shapes for
  further tweaks.

* **Charting** — chart palette presets independent of
  ``chart_style``; ten quick-layout presets; per-series gradient
  and pattern fills.

* **3D primitives and SmartArt text substitution** — bevel and
  extrusion via ``shape.three_d``;
  ``slide.smart_art[i].set_text([...])``.

* **Slide thumbnails** — ``Presentation.render_thumbnails()`` or
  ``pptx2.render.render_slides(prs, slides=[0, 1, 2],
  name_template="slide-{:02d}.png")`` shells out to LibreOffice for
  PNG previews.

See ``HISTORY.rst`` for the full changelog and ``ROADMAP.md`` for the
broader plan.

Attribution
-----------

This project is a fork of `scanny/python-pptx`_, originally created and
maintained by Steve Canny under the MIT License. The original
copyright notice is preserved in ``LICENSE``. Sincere thanks to Steve
and to all the upstream contributors whose work this project builds
on.

The fork was created to continue development of features the upstream
roadmap did not cover (notably effects, transitions, animations, theme
customization, and a higher-level design layer).  See ``HISTORY.rst``
for the divergence point and changelog from there forward.

This project is **not** affiliated with or endorsed by Microsoft.
"PowerPoint" is a trademark of Microsoft Corporation; it is used here
only descriptively to identify the file format the library reads and
writes.

Documentation
-------------

**Project site:** https://github.com/lofcz/python-pptx2 — the
Astro/React documentation site (source in ``site/``), plus the
Sphinx docs under ``docs/``.

The Sphinx documentation lives under ``docs/`` and covers both the
inherited 1.0.2 API and every feature added by the fork.  Browse
`examples with screenshots`_ to get a quick idea what you can do.

The bundled Claude Code skill (``python -m pptx2.skill install``)
is the most up-to-date entry point for using the post-fork APIs.

.. _`python-pptx`:
   https://github.com/scanny/python-pptx
.. _`scanny/python-pptx`:
   https://github.com/scanny/python-pptx
.. _`Steve Canny`:
   https://github.com/scanny
.. _`examples with screenshots`:
   https://python-pptx.readthedocs.org/en/latest/user/quickstart.html
