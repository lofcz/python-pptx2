---
name: python-pptx2
description: Build PowerPoint (.pptx) decks from Python with the python-pptx2 library — a fork of power-pptx / python-pptx. Use this skill whenever the user wants to generate, mutate, lint, theme, animate, or render PowerPoint decks programmatically. The headline reason this line exists is **space-awareness**: text that doesn't overflow its box and shapes that don't slide off the edges of the slide. Reach for it especially when generation is dynamic (LLM, DB, CLI, JSON spec) and the deck has to look right without manual cleanup. Other features include native LaTeX equations, visual effects, animations, transitions, theme writer, design tokens, slide recipes, slide thumbnails, chart palettes, SVG embedding, 3D, and SmartArt text substitution.
---

# python-pptx2

`python-pptx2` is a fork of `power-pptx` (and, through it, `python-pptx`),
distributed on PyPI as `python-pptx2` and imported as `import pptx2`.
Use it for every PowerPoint generation / mutation task.

## The headline: space-aware authoring

The single biggest reason this fork exists is to make programmatic
decks **physically correct**: text doesn't overflow its container,
shapes don't sit off the slide, and elements that overlap do so on
purpose. Three layered tools — used together — catch ~all real-world
issues:

1. **`TextFrame.fit_text(...)`** measures with Pillow font metrics
   and bakes a fitting size into the XML *before* save.
2. **`text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`** lets
   PowerPoint shrink at render time as a fallback.
3. **`slide.lint()`** catches what slipped through; `auto_fix()`
   nudges off-slide shapes back inside.

**Read `references/space-aware-authoring.md` first** if the user is
generating decks from any dynamic input. It's the reason this skill
exists.

## The second headline: decks that look designed

Physically correct is the floor. When a deck will be *seen* — a lesson,
a talk, a pitch — **read `references/slide-design.md`** before writing
slides. It gives you the working vocabulary of a designed deck and the
helpers that implement it:

- one **palette** (`PALETTES["slate"]` → `P.paper / P.ink / P.accent /
  P.surface / P.muted`; `P.dark()` for title and section slides),
- one **grid** (`TITLE` band + `BODY` region; `BODY.columns(n, gap=…)`),
- a **type scale** read from the back row (title 32–36, body 22–26, card
  body 18–20, caption 14 — nothing under 14 pt),
- **surfaces done well** — `add_card` is one tinted rounded rectangle with
  padded text inside, nothing attached to it,
- real **bullets** with hanging indents (`add_bullets`), pictures that
  keep their aspect (`add_picture_fit`), **code listings** the editor
  recognises as code blocks (`add_code`), palette-driven diagrams and
  tables,
- a catalog of **slide archetypes** (title, section, statement, big
  number, bullets + picture, three cards, process, comparison, quote,
  full-bleed photo, question, summary) — planned as a sequence before
  slide 2, consecutive slides never sharing one, dense alternating with
  light,
- a **visual per slide** chosen for its content: a photo for a real
  thing, a chart for numbers, a diagram for a process, a table for a
  comparison, typeset math for formulas,
- and a **visual check**: `prs.render_contact_sheet("preview.png")` — one
  PNG of every slide, looked at once before shipping.

The whole upstream 1.0.2 API still works — the rest of this skill
focuses on the post-fork additions because they're what's most often
missed by snippets pulled from the wider internet.

## Cheat sheet (most common operations)

The 25 calls that cover ~90% of deck-generation tasks. Reach for
`references/geometry-and-arrows.md` for the full surface; this is the
working set:

```python
from pptx2 import Presentation, BBox, audit, PALETTES
from pptx2 import add_card, add_bullets, add_picture_fit, add_code
from pptx2.diagrams import horizontal_pipeline, hub_and_spoke, cycle
from pptx2.enum.shapes import MSO_SHAPE
from pptx2.util import Inches, Pt

# --- open / save ---
prs = Presentation()                       # new blank deck
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)   # 16:9
prs = Presentation("file.pptx")            # open existing
prs.save("out.pptx")

# --- slides ---
slide = prs.slides.add_slide(prs.slide_layouts[5])   # Title Only
slide = prs.slides.add_slide(prs.slide_layouts[6])   # Blank

# --- one palette for the whole deck ---
P = PALETTES["slate"]                      # .paper .surface .line .ink .muted .accent .accent_soft
D = P.dark()                               # same hues on dark paper — title / section slides
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = P.paper

# --- geometry (BBox is splattable into add_*) ---
bb = BBox.from_inches(1, 2, 8, 4)
left, right = bb.split_h([7, 5], gap=Inches(0.6))
inner = bb.inset(all=Inches(0.2))

# --- n-up rows/grids: never hand-compute (avail - (n-1)*gap) / n ---
for cell in bb.columns(3, gap=Inches(0.5)):   # equal columns; .rows(n) too
    add_card(slide, cell, title="Step", body="What happens here.",
             fill=P.surface, title_color=P.ink, body_color=P.ink)

# --- text (one call) ---
slide.shapes.add_text(bb, text="Hello",
                      size_pt=24, bold=True,
                      color=P.ink, align="center")

# --- real bullets (hanging indent, spacing, fitted to the box) ---
add_bullets(slide, left, items=["First point", "Second point"],
            size_pt=22, color=P.ink)               # numbered=True for 1. 2. 3.

# --- picture that keeps its aspect ---
add_picture_fit(slide, "photo.jpg", right, mode="contain",   # or "cover" to fill + crop
                caption="Source: Wikimedia", caption_color=P.muted)

# --- source code: monospace listing, never a plain text box ---
add_code(slide, right, code='for i in range(3):\n    print(i)', language="python",
         size_pt=18)                               # theme="github-dark" | "github-light" | …; line_numbers=True
# the shape is named fika:code:<lang>:<theme>:<0|1> — editors with a native
# code block (Fika) turn it into a highlighted, editable element on import

# --- native equation from LaTeX (pip install "python-pptx2[math]") ---
slide.shapes.add_equation(bb, latex=r"\frac{a}{b}", size_pt=28)
para = slide.shapes.add_text(bb, text="Euler: ").text_frame.paragraphs[0]
para.add_math(r"e^{i\pi}+1=0")

# --- shape with chainable colour ---
slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *bb) \
    .fill_hex("#FFFFFF").line_hex("#0D0D0D", weight_pt=1.25)

# --- hand-rolled flat surface (add_card does all of this for you) ---
card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, *inner)
card.fill_hex(P.surface); card.line.fill.background()
card.shadow.clear()                        # kills the inherited effectRef too
card.corner_radius = Pt(10)                # not adjustments[0]

# --- arrow with proper triangular head + auto edge routing ---
start_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *left)
end_shape   = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *right)
slide.shapes.add_arrow(start=start_shape, end=end_shape,
                       head="triangle", color="#0B5CFF", weight_pt=1.5)

# --- format-preserving text replacement (templated placeholders) ---
title_shape.set_text_preserving_format("New title")

# --- picture replacement (broken / sub-quality picture → native shapes) ---
picture.replace_with(lambda slide, bbox: ..., padding=Inches(0.1))

# --- diagram recipes ---
horizontal_pipeline(slide, bb, steps=["Extract", "Classify", "Enrich"])
hub_and_spoke(slide, bb, centre="Agent",
              spokes=["Memory", "Tools", "Planning"])

# --- table styling without dropping to raw fill/font loops ---
table.format_cells(rows=0, fill="#1F2937", color="#FFFFFF", bold=True)
table.format_cells(rows=slice(1, None), size_pt=11, align="right")

# --- space-aware fit (returns the size it applied) ---
size = tf.fit_text(font_family="Inter", max_size=24)

# --- single-call cleanup before save ---
slide.tidy()                               # lints + safe auto-fixes

# --- tell the linter an overlap is deliberate (widest -> narrowest) ---
slide.lint_group_overlaps(photo, band, title)    # one visual cluster (hero image)
badge.allow_overlap_with(card)                   # exactly this one pair
card.layer, badge.layer_above = "card", "card"   # also asserts z-order
with slide.design_group("hero"): ...             # tags everything created inside

# --- validate at save time (off by default) ---
prs.lint_on_save = "raise"                 # or "warn"; raises before writing

# --- whole-deck audit (markdown summary) ---
print(audit(prs).markdown())

# --- LOOK at the deck: one PNG with every slide, numbered ---
prs.render_contact_sheet("preview.png", cols=4, thumb_width=480)

# --- or per-slide thumbnails ---
from pptx2.render import render_slides
render_slides(prs, slides=[0, 1, 2], out_dir="thumbs",
              name_template="slide-{:02d}.png")
```

That whole sheet fits on one screen — it's the working set.

## When to use this skill

- The user wants to **generate a deck** from Python or a JSON / dict spec
- The user is concerned about **text overflow** or **layout correctness**
  in generated decks (lead with `space-aware-authoring.md`)
- The user wants to **add visual effects** (shadow, glow, soft edges,
  blur, reflection, alpha) to shapes
- The user wants **animations**, **transitions**, or **motion paths**
- The user wants to **read or write a theme** (palette + fonts), or
  apply one from a `.potx`
- The user wants to **lint / auto-fix** geometry issues
- The user wants to **import a slide** between decks or **apply a template**
- The user wants a **design system** (tokens, recipes, Grid/Stack layout)
- The user wants **chart palettes**, **quick layouts**, or per-series
  gradient/pattern fills
- The user wants **slide thumbnails** rendered to PNG
- The user wants **3D** primitives (bevels / extrusion) or **SmartArt
  text substitution**
- The user wants **native SVG embedding** with PNG fallback

## Install

```bash
pip install python-pptx2
```

The `cairosvg` dependency is optional — install only if you want
`add_svg_picture(...)` to auto-rasterise the PNG fallback. `pyyaml` is
optional too — install only if you want `DesignTokens.from_yaml`.

## Reference snippets

This skill ships a `references/` directory with focused recipe
collections. Read just the file you need — they're self-contained.

| File | What it covers |
|---|---|
| `references/space-aware-authoring.md` | **READ THIS FIRST.** Pre-flight measurement (`fit_text`, `TextFitter.best_fit_font_size`), `auto_size` flags, the linter, and a robust layout pattern. **Phase 2 + Phase 6 text-fit estimator.** |
| `references/slide-design.md` | **READ THIS WHEN THE DECK WILL BE SEEN.** Grid and margins, type scale, one-palette colour (`PALETTES`, `.dark()`), surfaces done well (`add_card`), lists (`add_bullets`), pictures (`add_picture_fit`), tables/diagrams/charts styling, deliberate vs accidental overlaps, a 12-archetype slide catalog with runnable code, and the visual check (`render_contact_sheet`). **v3.1.** |
| `references/geometry-and-arrows.md` | `BBox` value object (`columns`/`rows`/`split_h`/`grid`), `add_text` / `add_arrow` / `fill_hex` / `line_hex` convenience, `set_text_preserving_format`, `Picture.replace_with`, `Slide.tidy()`, diagram recipes (`horizontal_pipeline`, `hub_and_spoke`, `cycle`, `decision_tree`, `comparison_columns`), `audit(prs)`. **v2.8.** |
| `references/lint.md` | Detail on `slide.lint()`, issue types, `auto_fix`, and the `from_spec(..., lint="raise")` hook. **Phase 2.** |
| `references/design.md` | `DesignTokens`, `shape.style` facade, `Grid` / `Stack` layout primitives (geometry-safe placement), slide recipes (`title_slide`, `bullet_slide`, `kpi_slide`, `quote_slide`, `image_hero_slide`), starter pack. **Phase 9.** |
| `references/math.md` | Native PowerPoint equations from LaTeX (`add_equation`, `paragraph.add_math`). Requires `python-pptx2[math]`. **v2.14.** |
| `references/basics.md` | The 1.0.2 surface: `Presentation`, slides, placeholders, shapes, textboxes, tables, pictures, charts. Quick-reference cheatsheet. |
| `references/effects.md` | **The canonical fills-and-effects reference.** Shadow (including `shadow.clear()`), glow, soft edges, blur, reflection, alpha-tinted colors, gradient fills (linear / radial / rectangular / shape), line ends/caps/joins/compound. Other files link here rather than repeat it. **Phase 3 + Phase 6.** |
| `references/animations.md` | `Entrance` / `Exit` / `Emphasis` presets, triggers, by-paragraph reveal, sequencing context manager, motion paths. **Phase 5.** |
| `references/transitions.md` | Per-slide and deck-wide transitions including Morph and other `p14:` extensions. **Phase 4.** |
| `references/compose.md` | `from_spec` (JSON authoring with built-in lint), `import_slide` / `append_deck` (paper-pptx validating engine + the low-level part engine), `apply_template`. **Phase 2 + Phase 7 + paper-pptx.** |
| `references/editing.md` | **Changing existing decks safely** (paper-pptx): `inspect_text` / `effective_font` with provenance, `replace_text`, `diff_decks`, `Slides.clone/delete/move`, `shapes.*_by_name`, table surgery (`insert_row` / `extend_merge`), `replace_image`, `replace_data_safe`, `normalize_autofit`, `Presentation.batch()`, typed refusals. |
| `references/theme.md` | Reading + writing the theme palette and fonts; `theme.apply(...)`; theme-aware color resolution via `pptx2.inherit.resolve_color`. **Phase 6 + Phase 7.** |
| `references/picture-effects.md` | Picture transparency / brightness / contrast / recolor (grayscale / sepia / washout / duotone) and SVG embedding. **Phase 6.** |
| `references/charts.md` | Chart palettes, quick layouts, per-series gradient/pattern fills, plus the inherited chart API. **Phase 10.** |
| `references/render.md` | Slide thumbnails via LibreOffice. **Phase 10.** |
| `references/three-d.md` | Bevels and extrusion via `shape.three_d`. **Phase 8.** |
| `references/smart-art.md` | Text substitution inside an existing template's SmartArt. **Phase 8.** |
| `references/tables.md` | The inherited table API, plus `cell.format(...)` / `table.format_cells(...)` styling, `Cell.borders`, and `fit_to_box`. |
| `references/end-to-end-deck.md` | A complete worked example: tokens, recipes, animations, transitions, charts, **and a lint pass before save**. |

## Top-level imports beyond `Presentation`

These are stable package-root re-exports — prefer them over deeper
import paths:

```python
from pptx2 import (
    Presentation,
    # Immutable rectangular region; splats into add_* APIs.
    BBox,
    # One-call deck audit (lint + picture + empty-slide + font checks).
    audit, AuditReport,
    # Figure adapters — Plotly / Matplotlib / SVG / HTML → slide picture.
    # Third-party deps are imported lazily; missing deps surface a clear
    # FigureBackendUnavailable with the right pip install command.
    add_plotly_figure, add_matplotlib_figure,
    add_svg_figure,    add_html_figure,
    FigureBackendUnavailable,
    MathBackendUnavailable,
    # Everyday slide blocks (hex-driven, no token setup). Each returns the
    # shapes it made and tags them as one lint group.
    add_card, add_bullets, add_picture_fit,
    Card, FittedPicture,
    # Curated colour sets: PALETTES["slate"], palette("linen"), P.dark().
    PALETTES, Palette, palette,
    # Dashboard-style components (token-driven).
    add_kpi_card, add_progress_bar,
    KpiCard,      ProgressBar,
)
```

## House rules for code you write

1. **Always `from pptx2 import Presentation`** — never invent another
   import path.
2. **Default to space-aware patterns** for any text the user controls
   at runtime: `fit_text` *or* `auto_size = TEXT_TO_FIT_SHAPE`, plus a
   `slide.lint()` pass before save.
3. **Reads should not mutate.** All effect / color / line proxies in
   python-pptx2 return `None` for unset properties; assign `None` to
   clear.
4. **Use EMU through helpers**: `Inches`, `Pt`, `Emu`, `Cm` from
   `pptx2.util`. Never write raw EMU integers when a helper exists.
5. **Use `BBox` / `Grid` / `Stack` for placement** when you have more
   than two shapes on a slide — they compute geometry from the slide's
   real dimensions (or a region you hand them), so you can't
   accidentally walk off the right edge, and there's no column
   arithmetic to get wrong.
6. **Prefer recipes for whole-slide layouts** when the user wants a
   "good enough" pitch deck; drop down to direct `add_shape` /
   `add_textbox` only when the recipes don't fit.
7. **One palette, one accent.** Pick `PALETTES[...]` (or build a
   `Palette`) once and take every colour from its roles. The accent goes
   on the single most important thing on a slide.
8. **Surfaces are `add_card`, lists are `add_bullets`, pictures are
   `add_picture_fit`.** They encode the padding, indents, fitting and
   lint-grouping that hand-rolled versions tend to miss, and the card's
   silhouette stays one clean rounded rectangle.
9. **Look before you ship.** `prs.render_contact_sheet("preview.png")`
   once after `tidy()`, fix what you see, render once more.
10. **Save once at the end** — build the deck in memory, then call
    `prs.save(...)`. Don't open and re-save inside loops.
11. **For released-version constraints**: pin `python-pptx2>=3.1.0`
    when generating requirements files — that's the minimum that
    ships `add_card` / `add_bullets` / `add_picture_fit`, `PALETTES`
    and `render_contact_sheet` used in this skill.

## A space-aware mini-template

The pattern you'll reach for most often:

```python
from pptx2 import Presentation
from pptx2.enum.text import MSO_AUTO_SIZE
from pptx2.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
slide.shapes.title.text = "Q4 Review"

# Body box that has to swallow runtime-supplied text
box = slide.shapes.add_textbox(Inches(0.6), Inches(1.6),
                                Inches(12), Inches(5))
tf = box.text_frame
tf.word_wrap = True
tf.text = USER_SUPPLIED_BODY

# Belt: pick a determined size now using Pillow font metrics
tf.fit_text(font_family="Inter", max_size=24)

# Braces: let PowerPoint shrink on the way down if a user later edits
tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

# Catch anything that slipped through. auto_fix() mutates the slide
# (nudges OffSlide shapes in, autofits overflowing frames, restacks
# contradicted layer declarations) and refreshes report.issues itself.
slide.lint().auto_fix()
report = slide.lint()
errors = [i for i in report.issues if i.severity.value == "error"]
if errors:
    raise RuntimeError("\n".join(str(e) for e in errors))

prs.save("out.pptx")
```

## Recent additions worth knowing

These changes ship after v2.5 and are easy to miss:

- **`add_card` / `add_bullets` / `add_picture_fit`** (v3.1, package
  root) — the everyday blocks of a content slide, hex-driven. `add_card`
  is one flat rounded surface with padded, fitted title/body (body may be
  a list → bullets); `add_bullets` writes real `a:buChar` / `a:buAutoNum`
  bullets with hanging indents, spacing and fit-to-box; `add_picture_fit`
  letter-boxes (`contain`) or crops (`cover`) a picture into a box with an
  optional caption. All three tag their shapes as one lint group.
- **`PALETTES` / `Palette` / `palette()`** (v3.1) — curated colour sets
  with named roles (`paper`, `surface`, `line`, `ink`, `muted`, `accent`,
  `accent_soft`, `accent_ink`) that pass contrast checks; `P.dark()`
  derives a dark variant for title / section slides.
- **`prs.render_contact_sheet("preview.png")`** (v3.1) — every slide
  tiled into one numbered PNG for a single visual check.
- **`Chart.recolour(palette)`** is the recommended single entry
  point — auto-dispatches per chart type (per-point on pie /
  doughnut, per-series otherwise). `apply_palette` warns and
  routes when called on a doughnut.
- **`Chart.line_color`** and **`Chart.apply_dark_theme(text=, line=)`**
  pin axis lines + gridlines for dark-deck styling.
- **Horizontal bar charts (`BAR_*`)** now default to top-to-bottom
  reading order (`reverse_order=True`). Override with
  `chart.category_axis.reverse_order = False` for legacy ordering.
  Column charts are unaffected.
- **`anchor=` keyword on `add_picture` / `add_shape` / `add_textbox`**
  collapses corner / centre placement to one call (see
  `references/basics.md`).
- **`add_table(..., style="clean")`** disables every inherited style
  flag — use it whenever you'll set custom cell borders or fills.
- **`add_kpi_card(slide, ...)` / `add_progress_bar(slide, ...)`** —
  shape-level building blocks beneath the slide-level recipes
  (see `references/design.md`).
- **`shape.shadow.clear()`** is the way to guarantee *no* shadow.
  Clearing the individual shadow properties (or the deprecated
  `shadow.inherit = False`) leaves the shape's `<a:effectRef idx="2"/>`
  in place — a soft drop shadow in most themes — so cards keep a
  phantom shadow nobody asked for. `clear()` drops the explicit shadow
  elements *and* re-points the effect reference, keeping glow / soft
  edges / blur / reflection intact.
- **`shape.corner_radius`** reads and writes a rounded rectangle's
  radius as a real length (`card.corner_radius = Pt(6)`), instead of the
  `adjustments[0]` fraction-of-the-shorter-side that has to be eyeballed
  per shape size.
- **`BBox.columns(n, gap=...)` / `BBox.rows(n, gap=...)`** replace the
  `(available - (n - 1) * gap) / n` loop for card rows, stat grids, and
  panels. `Grid.from_box(box, cols=..., rows=...)` puts a full grid over
  a region rather than the whole slide.
- **`cell.format(...)` / `table.format_cells(rows=..., cols=..., ...)`**
  style table cells with the same keywords as `add_text` (`fill`,
  `color`, `bold`, `size_pt`, `align`, `anchor`, `margin`) — no more
  `cell.fill.solid(); cell.fill.fore_color.rgb = ...` loops.
- **`fit_text` is explicit about fallback metrics.** It returns the size
  it applied, and warns (`FontMetricsWarning`) when a *named* family
  isn't installed and measurement silently drops to Pillow's default
  font — naming one is what makes it audible, so omitting `font_family`
  stays quiet while `font_family="Calibri"` does not. `strict=True`
  raises instead; `pptx2.text.fonts.font_is_installed("Inter")`
  checks up front.
- **Float coordinates from arithmetic are coerced** at constructor
  entry and at `shape.left/top/width/height` setters, so
  `(Inches(N) - gutter) / 2` style expressions can be passed straight
  through. Pre-2.6.1 these produced float-valued `<a:off>` / `<a:ext>`
  attributes that PowerPoint rejected with the "Repair?" dialog.

## Anti-patterns to avoid

LLM-generated python-pptx2 code falls into the same handful of traps.
Flagging them up front saves the trial-and-error round.

- **Don't** access `tf.paragraphs` twice and compare wrapper objects
  with `is`. The property returns *fresh* `_Paragraph` objects every
  call, so `p is not para` is always true — your filter will remove
  the wrong paragraph. Use `set_text_preserving_format(new_text)` for
  the common "replace text, keep formatting" case.
- **Don't** assume `add_connector(MSO_CONNECTOR.STRAIGHT, ...)` puts
  an arrowhead on the line. It produces a bare line. Use
  `slide.shapes.add_arrow(start, end, head="triangle")` — it sets the
  arrowhead, inset, edge routing, and colour in one call.
- **Don't** size a diagram to a broken picture's bbox when there's an
  enclosing card. The card area is what you want. Use
  `picture.enclosing_container()` to find the right box, then
  `picture.replace_with(builder)`.
- **Don't** delete a picture and then assume sibling shape indexes
  are stable. Process in reverse index order, or capture the shapes
  to mutate *before* iterating.
- **Don't** import `RGBColor` / `PP_ALIGN` / `MSO_VERTICAL_ANCHOR`
  for every styling call. Use the hex-string and short-name kwargs:
  `slide.shapes.add_text(bb, text="…", color="#0B5CFF", align="center",
  anchor="middle")`. Hex strings, tuples, and `RGBColor` all work
  everywhere a colour is accepted.
- **Don't** write raw EMU integers. `BBox.from_inches(1, 2, 8, 4)`
  for regions, `Inches(1)` / `Pt(12)` for individual lengths. Float
  arithmetic on EMU is fine — coordinates are coerced at the setter.
- **Don't** lint + auto_fix + lint again to clear safe issues. Use
  `slide.tidy()` — it's the one-call wrapper.
- **Don't** hand-roll `col_w = (avail - (n - 1) * gap) / n` and a running
  cursor for card rows or stat grids. `bb.columns(n, gap=Pt(16))` (and
  `bb.rows(...)`, `bb.grid(cols, rows)`, `Grid.from_box(bb, cols=...)`)
  return exact, drift-free boxes.
- **Don't** hand-roll a `card()` helper that draws a rounded rectangle
  and then a second thin rectangle on its edge, a badge in its corner, or
  a rail down the side of the slide. Those extra shapes are what make
  generated decks look generated (and they trip `ShapeCollision` for a
  reason — a square stripe pokes out of a rounded corner). A card is
  `add_card(...)`: one surface, padded text, done. Emphasis is the
  fill tint (`P.accent_soft`) or the title weight.
- **Don't** write bullets as `"•  " + text` in a plain paragraph. That
  bullet does not hang, wrapped lines run under it, and the list
  misaligns. `add_bullets(slide, bb, items=[...])` writes real bullets.
- **Don't** draw `add_picture(path, left, top, width, height)` with both
  dimensions from a box — that stretches the photo.
  `add_picture_fit(slide, path, bb, mode="contain"|"cover")` keeps the
  aspect.
- **Don't** pick colours ad hoc per slide (a teal here, a navy there, an
  orange badge). Choose one `Palette` and take every colour from its
  roles; one accent per slide.
- **Don't** try to remove a shadow by assigning `None` to
  `shadow.blur_radius` / `distance`, or by `shadow.inherit = False`.
  Neither touches the theme effect style, so the shadow is still
  rendered. Use `shape.shadow.clear()` — the only call that does.
- **Don't** set a corner radius by guessing `adjustments[0]` (it's a
  fraction of the shorter side, so the same value means a different
  radius on every differently-sized card). Use
  `shape.corner_radius = Pt(6)`.
- **Don't** invent OMML / `a14:m` XML, and don't rasterise ordinary
  formulas with matplotlib just to get them on a slide. Use
  `slide.shapes.add_equation(bb, latex=r"\frac{a}{b}")` or
  `paragraph.add_math(...)` (`pip install "python-pptx2[math]"`).
- **Don't** style table cells with `cell.fill.solid()` +
  `cell.fill.fore_color.rgb` + per-run font loops. Use
  `cell.format(...)` / `table.format_cells(...)` — and note that a
  cell's anchor and insets belong on the cell, not on its text frame.
- **Don't** trust `fit_text` when the font isn't installed. The
  measurement falls back to Pillow's default metrics and the result is
  an estimate — bundle the `.ttf` and pass `font_file=`, or pass
  `strict=True` so the build fails instead of shipping a guess.

## Common pitfalls

- **Calling `shape.shadow.inherit`** raises `DeprecationWarning`. Read
  individual properties (`blur_radius`, `distance`, `direction`,
  `color`) and check for `None` instead; to *remove* a shadow call
  `shape.shadow.clear()`.
- **`fit_text` degrades to a best guess for uninstalled fonts.** Font
  metrics come from the machine running the build, so a brand face that
  isn't installed (the usual case in a container or CI) is measured with
  Pillow's default font. Check with
  `pptx2.text.fonts.font_is_installed(...)`, pass `font_file=` with
  the real `.ttf`, or use `strict=True`. `slide.lint()` is unaffected —
  its overflow check is font-agnostic — which is why the lint pass is
  worth keeping even when the metrics are exact.
- **Bare-int sizes in `DesignTokens` typography** are interpreted as
  **EMU**, not points. Use floats (`44.0`) or `Pt(44)` to mean
  44-point font.
- **Recipes use the Blank layout**, so `slide.shapes.title` is `None`.
  Address shapes by index (`slide.shapes[0]`, `slide.shapes[1]`, …).
- **`add_svg_picture` without `cairosvg` and without a `png_fallback`**
  raises `CairoSvgUnavailable`. Either install cairosvg or supply a
  pre-rasterised PNG.
- **`auto_fix()` repairs geometry, not judgment.** It clamps `OffSlide`,
  flips `TextOverflow` frames to `TEXT_TO_FIT_SHAPE`, snaps
  `OffGridDrift`, and restacks `LayerOrderViolation`. It will never
  touch `ShapeCollision`, `LowContrast`, `MinFontSize` or
  `ZOrderAnomaly` — those need a designer. Prefer `tf.fit_text(...)`
  *before* save over relying on the overflow fix.
- **Slide thumbnails require `soffice` on PATH** (LibreOffice).
  Otherwise you get `ThumbnailRendererUnavailable`.
- **`MSO_PATTERN_TYPE.ERCENT_40`** is the upstream typo and emits a
  `DeprecationWarning`. Use `PERCENT_40`.
- **Calling `chart.apply_palette` on a pie / doughnut** emits a
  `UserWarning` and routes through `color_by_category`. Use
  `chart.recolour(palette)` directly to silence it.

## Where to look in the project

If the user has the `python-pptx2` repo checked out alongside this
skill, these paths are useful for source-of-truth lookup:

- `src/pptx2/lint.py` — `SlideLintReport`, `TextOverflow`, `OffSlide`,
  `ShapeCollision`, `LayerOrderViolation`, `LintSeverity`.
- `src/pptx2/shapes/base.py` — `lint_group`, `lint_skip`,
  `allow_overlap_with`, `layer` / `layer_above` (overlap intent).
- `src/pptx2/text/text.py`, `src/pptx2/text/layout.py` — `fit_text`,
  `TextFitter`, `_best_fit_font_size`.
- `src/pptx2/animation.py` — `Entrance`, `Exit`, `Emphasis`,
  `MotionPath`, `SlideAnimations`.
- `src/pptx2/compose/` — `from_spec`, plus the `import_slide` /
  `apply_template` re-exports.
- `src/pptx2/theme.py`, `src/pptx2/inherit.py` — theme reader/writer and
  `resolve_color`.
- `src/pptx2/dml/effect.py`, `src/pptx2/dml/picture.py`,
  `src/pptx2/dml/line.py` — Phase 3/6 visual effects, picture filters,
  line-end formatting.
- `src/pptx2/design/` — `tokens`, `style`, `layout`, `recipes`,
  `components`, `blocks` (`add_card` / `add_bullets` / `add_picture_fit`),
  `palettes` (`PALETTES`, `Palette`).
- `src/pptx2/chart/palettes.py`, `src/pptx2/chart/quick_layouts.py`.
- `src/pptx2/render.py` — slide-thumbnail renderer + `render_contact_sheet`.
- `src/pptx2/smart_art.py`, `src/pptx2/_svg.py`.
- `examples/starter_pack/` — three example token sets and a build script.

The user-facing Sphinx documentation under `docs/user/` mirrors the
sections in this skill and is a good source of additional prose.
