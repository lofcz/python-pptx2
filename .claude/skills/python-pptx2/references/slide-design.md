# Slide design: making a generated deck look designed

`space-aware-authoring.md` makes a deck *physically* correct — nothing
overflows, nothing falls off the slide. This file is about the other
half: a deck that looks like a designer laid it out. Everything here is
built from the public API (`BBox`, `add_text`, `add_card`, `add_bullets`,
`add_picture_fit`, `PALETTES`, `pptx2.diagrams`, `render_contact_sheet`),
so the code samples run as written.

Read this when the deck will be *seen* — a lesson, a talk, a pitch — and
not only parsed.

## The one-paragraph brief

A slide that looks designed has **one idea**, **one focal point**, **one
palette**, **one type family**, generous **air** around everything, and
text that **fits at a readable size**. The viewer knows within a second
where to look first. Every element on the slide is *content* (a title, a
claim, a list, a picture, a table, a diagram, an equation) or *structure*
(a surface, a rule, a caption); there is nothing whose only job is to
decorate. When the content is strong the slide needs nothing else, and
when the content is thin, ornament will not save it.

## 1. Start from the content

Before drawing anything, write down for each slide:

- **Purpose** — one sentence: what the viewer should know or feel after it.
- **Carrier** — the single element that delivers the purpose: a
  statement, a short list, a picture, a table, a process, an equation.
- **Support** — at most two more elements that help the carrier (a
  caption, a source, a question, a picture beside a list).
- **Archetype** — pick one from the catalog in §10. Reusing a handful of
  archetypes is what makes a deck feel consistent.

A deck of 8–12 slides typically needs 4–5 distinct archetypes. Alternate
dense and light slides: after a table or a three-card slide, give the
audience a statement or a full-bleed picture.

## 2. Canvas, margins, grid

Work in 16:9 at 13.333 × 7.5 in, and set it explicitly:

```python
from pptx2 import Presentation, BBox
from pptx2.util import Inches

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
W, H = 13.333, 7.5
BLANK = prs.slide_layouts[6]
```

Then let two regions govern every content slide:

```python
TITLE = BBox.from_inches(0.8, 0.55, W - 1.6, 0.9)     # title band
BODY  = BBox.from_inches(0.8, 1.7,  W - 1.6, H - 2.4)  # content area (ends at 6.8")
```

- **Margins**: 0.8 in left/right, 0.55 in top, 0.7 in bottom. Nothing but
  a full-bleed picture or a coloured background touches the slide edge.
- **Gutters**: 0.4–0.6 in between columns and cards. Use one value per deck.
- **Columns**: `BODY.columns(3, gap=Inches(0.5))`,
  `BODY.split_h([7, 5], gap=Inches(0.6))`, `BODY.rows(2, gap=...)`. Every
  card, column and picture on a slide starts at the same left edge as the
  title, or at a grid line derived from it.
- **Optical center**: single statements and hero titles sit slightly
  above the geometric middle (`anchor="middle"` in a box whose top is
  ~1.6 in and bottom ~5.4 in).

## 3. Typography

One family for the whole deck. Noto Sans (or whatever the environment's
default sans is) is a fine choice; the design comes from *scale* and
*spacing*, not from the font.

| Role | Size | Weight | Notes |
|---|---|---|---|
| Hero title (title / section slide) | 44–54 pt | bold | 1–2 lines |
| Statement (the one-sentence slide) | 36–44 pt | bold | ≤ 20 words, centred or left |
| Slide title | 30–36 pt | bold | one line; sentence case |
| Body / bullets | 20–24 pt | regular | ≤ 6 items, ≤ 12 words each |
| Card title | 20–24 pt | bold | |
| Card body | 16–18 pt | regular | 2–4 lines |
| Table cells | 16–20 pt | header bold | |
| Caption / source / footer | 12–14 pt | regular or italic, `muted` | |

Rules that make type look intentional:

- **Sentence case** for titles ("Co rostlina potřebuje", not "CO ROSTLINA
  POTŘEBUJE"). Uppercase belongs to acronyms.
- **Left-align running text.** Centre only short statements, titles on
  hero slides, and labels inside diagram nodes.
- **Breathing room**: `line_spacing = 1.1–1.2` for body, `space_after =
  8–10 pt` between bullets (`add_bullets(gap_pt=8)` does this).
- **Zero text-box margins** when the box is already placed by the grid
  (`margin_pt=0`) so the text edge *is* the grid line.
- **Hierarchy by size and weight, not by colour.** Reserve colour for the
  one thing on the slide that should pop (§4).
- **Fit, then check**: `add_bullets` and `add_card` shrink text to the box
  and never below 12 pt. If they had to shrink, there is too much text —
  split the slide.

```python
slide.shapes.add_text(TITLE, text="Co rostlina potřebuje",
                      size_pt=32, bold=True, color=P.ink,
                      anchor="middle", margin_pt=0)
```

## 4. Colour

Pick **one** palette and use only its roles:

```python
from pptx2 import PALETTES

P = PALETTES["slate"]      # slate · linen · forest · plum · ember · navy · graphite
D = P.dark()               # same hues on a dark paper — title, section, closing slides
```

| Role | Use |
|---|---|
| `P.paper` | slide background (`slide.background.fill.solid(); ...fore_color.rgb = P.paper`) |
| `P.ink` | all headings and body text |
| `P.muted` | captions, sources, secondary labels, table grid text |
| `P.surface` | card fills, table zebra rows, picture frames |
| `P.line` | hairline rules and table borders |
| `P.accent` | **one** emphasis per slide: the key number, the active step, a highlighted word, the header row |
| `P.accent_soft` | fill of a callout card / highlighted row; text on it stays `P.ink` |
| `P.accent_ink` | text placed *on* `P.accent` (white on a dark accent) |

- **One accent per slide.** Two accents split attention; three look like a
  toy. Cards in a row all share `P.surface`; the one to emphasise (if any)
  gets `P.accent_soft`.
- **Dark slides** for the title, section dividers and the closing slide
  give the deck rhythm. Use `D = P.dark()` so the hues stay in the family.
- **Contrast**: every text/background pair ≥ 4.5:1 (the linter reports
  `LowContrast`); the shipped palettes satisfy this for `ink` on `paper`,
  `surface` and `accent_soft`, for `muted` on `paper`, and for `accent_ink`
  on `accent`.
- **Photos bring their own colour.** Beside a photo, keep the rest of the
  slide to `paper` / `ink` / `muted`.

## 5. Surfaces: cards, callouts, panels

A card is one tinted rounded rectangle with padded text inside it, and
that is the whole card. Its silhouette stays a clean rectangle from every
angle; emphasis comes from the tint (or, when the tint is unwanted, a
hairline outline — one or the other, never both). Padding of 18–28 pt on
every side gives the text room to sit rather than press against the edge.

```python
from pptx2 import add_card

for cell, (title, body) in zip(BODY.columns(3, gap=Inches(0.5)), items):
    add_card(slide, cell, title=title, body=body,
             fill=P.surface, title_color=P.ink, body_color=P.ink,
             title_size_pt=22, body_size_pt=18, pad_pt=22)
```

Composition rules for groups of cards:

- **2–4 cards per row**, equal size, equal gap, same fill, same radius.
  Five or more become a table.
- **Same content shape in every card**: if one has a title and two lines,
  all have a title and two lines. Trim copy to match.
- **A callout** (question, key definition, "remember this") is one wide
  card in `P.accent_soft`, inset from the body region so it reads as a
  deliberate object, not as a background.
- **Cards hold content, backgrounds hold cards.** A card sits directly on
  the paper; it does not sit on another panel.
- **Numbering** belongs *in* the title text ("1 · Zachycení"), where it
  aligns with everything else, rather than in a separate badge shape.

`add_card` returns `card.inner` (the padded box) so a picture, equation
or diagram can be placed inside the same surface:

```python
c = add_card(slide, right_col, fill=P.surface)
slide.shapes.add_equation(c.inner, latex=r"6\,CO_2 + 6\,H_2O \rightarrow C_6H_{12}O_6 + 6\,O_2", size_pt=28)
```

## 6. Whitespace and alignment

- 30–40 % of a content slide is empty. That is the design working.
- **One left edge.** Title, bullets, cards and pictures share the `0.8 in`
  margin (or a grid line derived from it). Right edges align to the
  mirror margin.
- **Consistent vertical rhythm.** Title band → 0.25 in gap → body. Every
  content slide places its body at the same `BODY.top`, so flipping
  through the deck feels steady.
- **Balance, not symmetry.** A picture on the right with a list on the
  left is balanced when the list's text block is roughly as tall as the
  picture. Use `split_h([7, 5])` for text-heavy pairs, `[1, 1]` for a
  true comparison.
- **Empty is fine, unbalanced is not.** A slide whose right half is blank
  because the list is short: make the list larger (24 pt), add a picture,
  or narrow the text column and centre the pair.

## 7. Pictures

Photos carry emotion and specificity — one good photo does more than three
small ones.

```python
from pptx2 import add_picture_fit

# Beside text: fit inside the column, never distort, caption in muted.
add_picture_fit(slide, "leaf.jpg", right_col, mode="contain",
                caption="Buk lesní · Wikimedia Commons", caption_color=P.muted)

# Full-bleed hero: crop to fill the whole slide, then a dark band for text.
with slide.design_group("hero"):             # photo + band + title are one deliberate stack
    add_picture_fit(slide, "forest.jpg", BBox.from_inches(0, 0, W, H), mode="cover")
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *BBox.from_inches(0, H - 2.2, W, 2.2))
    band.fill_hex(D.paper); band.line.fill.background(); band.shadow.clear()
    band.fill.fore_color.alpha = 0.8         # translucent, photo shows through
    slide.shapes.add_text(BBox.from_inches(0.8, H - 1.9, W - 1.6, 1.6),
                          text="Les jako plíce planety", size_pt=44, bold=True,
                          color=D.ink, anchor="middle", margin_pt=0)
```

- **`mode="contain"`** when the whole picture matters (diagrams, maps,
  artworks); **`mode="cover"`** for atmosphere (backgrounds, hero slides,
  square tiles in a grid).
- **Picture grids**: `BODY.columns(3)` → three `mode="cover"` tiles of
  equal size, one caption each.
- **Caption every photo** that is not purely decorative: what it shows and
  the source, 12–14 pt, `P.muted`.
- Leave pictures unframed. A `P.surface` background behind a letter-boxed
  `contain` picture is the one exception, so the empty band reads as a
  frame rather than a gap.

## 8. Tables, charts, diagrams

**Tables** — structure carries the styling:

```python
tbl = slide.shapes.add_table(rows, cols, *BODY.inset(bottom=Inches(1.0)), style="clean").table
tbl.format_cells(rows=0, fill=P.ink, color=P.paper, bold=True, size_pt=18)
tbl.format_cells(rows=slice(1, None), size_pt=18, color=P.ink)
tbl.format_cells(rows=slice(1, None, 2), fill=P.surface)      # zebra
tbl.format_cells(rows=slice(1, None), cols=0, bold=True)      # row labels
```

≤ 6 body rows and ≤ 4 columns on a slide; more belongs in a handout. The
header row is the accent of the slide — use `P.ink` or `P.accent`, not
both.

**Diagrams** — `pptx2.diagrams` recipes already space nodes and route
arrows; feed them the palette:

```python
from pptx2.diagrams import horizontal_pipeline, cycle, hub_and_spoke, comparison_columns

horizontal_pipeline(slide, BODY.inset(top=Inches(1.4), bottom=Inches(1.4)),
                    steps=["Vstupy", "Světlo + chlorofyl", "Výstupy"],
                    accent=P.accent, fill=P.surface, text_color=P.ink,
                    card_line=None, card_radius=10, size_pt=20)
```

A pipeline with 3–5 steps, a cycle with 3–6 nodes, a hub with 3–6 spokes.
Node labels are 2–4 words; explanations go in a list or a card beside the
diagram.

**Charts** — `chart.recolour([P.accent, P.muted, P.line])`, gridlines in
`P.line`, one series in `P.accent` and the rest in greys when there is a
story to tell. Title the chart with the takeaway ("Kyslík roste s
osvětlením"), not the variable names.

**Equations** — `slide.shapes.add_equation(bb, latex=..., size_pt=28)` or
`paragraph.add_math(...)`; display equations at 28–36 pt, centred, with
air around them.

## 9. Overlaps: deliberate and accidental

Overlap is a normal design move — text on a card, a caption on a photo, a
translucent band over a hero image, a number disc on a timeline. The
linter's `ShapeCollision` warning asks one question: *did you mean it?*

- **Meant it** → say so, and the warning disappears:

  ```python
  with slide.design_group("hero"):           # everything created inside is one cluster
      add_picture_fit(slide, "forest.jpg", BBox.from_slide(slide), mode="cover")
      band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *BBox.from_inches(0, H - 2.2, W, 2.2))
      slide.shapes.add_text(BBox.from_inches(0.8, H - 1.9, W - 1.6, 1.6), text="…")

  slide.lint_group_overlaps(card, picture, caption)   # shapes already made, widest first
  badge.allow_overlap_with(card)                      # exactly this pair
  card.layer, badge.layer_above = "card", "card"      # also asserts z-order
  ```

  `add_card`, `add_picture_fit`, the components and the diagram recipes
  already tag their own shapes, so a card and its text never trigger it.

- **Did not mean it** → it is a layout bug: two text boxes crossing, a
  card poking into its neighbour, a picture over a title. Fix the
  geometry (use `columns()` / `split_h()` instead of hand arithmetic) and
  re-lint.

Deliberate stacking looks deliberate when the upper shape is clearly
*inside* the lower one (inset on every side) or clearly *anchored* to one
of its corners/edges. A shape that half-crosses a boundary reads as a
mistake even when it is intentional.

`slide.tidy()` runs the safe auto-fixes (`OffSlide`, `TextOverflow`,
`OffGridDrift`); collisions, contrast and minimum size stay as warnings
because they need a judgment call. Make it, then move on.

## 10. Archetype catalog

Each archetype below is complete: the code runs given the setup in §2 and
`P`/`D` from §4. Mix 4–5 of them across a deck.

```python
def paper(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color

def content_slide(title):
    """Title band + body region. Every content archetype starts here."""
    s = prs.slides.add_slide(BLANK)
    paper(s, P.paper)
    s.shapes.add_text(TITLE, text=title, size_pt=32, bold=True,
                      color=P.ink, anchor="middle", margin_pt=0)
    return s
```

**Title (dark)** — hero title, one-line subtitle, nothing else.

```python
s = prs.slides.add_slide(BLANK); paper(s, D.paper)
s.shapes.add_text(BBox.from_inches(0.9, 2.2, W - 1.8, 1.6), text="Fotosyntéza",
                  size_pt=54, bold=True, color=D.ink, anchor="bottom", margin_pt=0)
s.shapes.add_text(BBox.from_inches(0.9, 3.9, W - 1.8, 0.8),
                  text="Jak rostliny mění světlo v život · 7. ročník",
                  size_pt=22, color=D.muted, margin_pt=0)
```

**Section divider (dark)** — number + section name, optically centred.

```python
s = prs.slides.add_slide(BLANK); paper(s, D.paper)
s.shapes.add_text(BBox.from_inches(0.9, 2.4, W - 1.8, 0.7), text="Část 2",
                  size_pt=20, color=D.accent, margin_pt=0)
s.shapes.add_text(BBox.from_inches(0.9, 3.0, W - 1.8, 1.4), text="Co se děje v listu",
                  size_pt=44, bold=True, color=D.ink, margin_pt=0)
```

**Statement** — one sentence the audience should remember; optional
follow-up question in the accent.

```python
s = prs.slides.add_slide(BLANK); paper(s, P.paper)
s.shapes.add_text(BBox.from_inches(1.2, 1.8, W - 2.4, 3.2),
                  text="Každý nádech kyslíku kdysi vyrobila rostlina.",
                  size_pt=40, bold=True, color=P.ink, anchor="middle", align="center")
s.shapes.add_text(BBox.from_inches(1.2, 5.2, W - 2.4, 0.6),
                  text="Proč je fotosyntéza nejdůležitější reakce na Zemi?",
                  size_pt=18, color=P.accent, align="center")
```

**Bullets + picture** — the workhorse. List left (7/12), picture right
(5/12), caption under the picture.

```python
s = content_slide("Co rostlina potřebuje")
left, right = BODY.split_h([7, 5], gap=Inches(0.6))
add_bullets(s, left, items=[
    "Světlo — energii dodává Slunce",
    "Vodu — kořeny ji přivádějí z půdy",
    "Oxid uhličitý — vstupuje průduchy",
    "Chlorofyl — zelené barvivo, které světlo zachytí",
], size_pt=22, color=P.ink)
add_picture_fit(s, "leaf.jpg", right, mode="contain",
                caption="List buku · Wikimedia Commons", caption_color=P.muted)
```

**Three cards** — parallel concepts, equal weight.

```python
s = content_slide("Tři fáze v jednom listu")
for cell, (t, b) in zip(BODY.columns(3, gap=Inches(0.5)), [
    ("1 · Zachycení", "Chlorofyl v chloroplastech pohltí světelnou energii."),
    ("2 · Rozklad vody", "Energie rozdělí vodu na kyslík a vodík."),
    ("3 · Tvorba cukru", "Vodík se spojí s CO₂ a vzniká glukóza."),
]):
    add_card(s, cell, title=t, body=b, fill=P.surface,
             title_color=P.ink, body_color=P.ink, title_size_pt=22, body_size_pt=18)
```

**Process** — steps with arrows, vertically centred in the body.

```python
s = content_slide("Rovnice fotosyntézy jako proces")
horizontal_pipeline(s, BODY.inset(top=Inches(1.4), bottom=Inches(1.4)),
                    steps=["6 CO₂ + 6 H₂O", "světlo + chlorofyl", "C₆H₁₂O₆ + 6 O₂"],
                    accent=P.accent, fill=P.surface, text_color=P.ink,
                    card_line=None, card_radius=10, size_pt=20)
```

**Comparison table** — two things side by side, row labels bold.

```python
s = content_slide("Fotosyntéza vs. dýchání")
rows = [("", "Fotosyntéza", "Dýchání"), ("Kde", "chloroplasty", "mitochondrie"),
        ("Kdy", "jen na světle", "stále"), ("Vstup", "CO₂ + H₂O", "glukóza + O₂"),
        ("Výstup", "glukóza + O₂", "CO₂ + H₂O + energie")]
tbl = s.shapes.add_table(len(rows), 3, *BODY.inset(bottom=Inches(1.0)), style="clean").table
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        tbl.cell(r, c).text = val
tbl.format_cells(rows=0, fill=P.ink, color=P.paper, bold=True, size_pt=18)
tbl.format_cells(rows=slice(1, None), size_pt=18, color=P.ink)
tbl.format_cells(rows=slice(1, None), cols=0, bold=True)
```

**Two-column contrast** — "before / after", "myth / fact", "pros / cons".

```python
s = content_slide("Mýtus a skutečnost")
a, b = BODY.columns(2, gap=Inches(0.6))
add_card(s, a, title="Mýtus", body="Rostliny dýchají jen v noci.",
         fill=P.surface, title_color=P.muted, body_color=P.ink, title_size_pt=20, body_size_pt=20)
add_card(s, b, title="Skutečnost", body="Dýchají nepřetržitě; ve dne fotosyntéza převažuje.",
         fill=P.accent_soft, title_color=P.ink, body_color=P.ink, title_size_pt=20, body_size_pt=20)
```

**Full-bleed picture with a title band** — see §7 (`mode="cover"` + a
translucent band in `D.paper` + title in `D.ink`).

**Quote** — large italic statement, attribution in `muted`.

```python
s = prs.slides.add_slide(BLANK); paper(s, P.paper)
s.shapes.add_text(BBox.from_inches(1.4, 1.8, W - 2.8, 3.0),
                  text="„Rostlina je stroj na slunce.“", size_pt=40, italic=True,
                  color=P.ink, anchor="middle", align="center")
s.shapes.add_text(BBox.from_inches(1.4, 4.9, W - 2.8, 0.6), text="— Jan Evangelista Purkyně",
                  size_pt=18, color=P.muted, align="center")
```

**Question / activity** — one callout card, numbered prompts.

```python
s = content_slide("Zamyslete se")
add_card(s, BODY.inset(left=Inches(1.5), right=Inches(1.5), top=Inches(0.6), bottom=Inches(0.6)),
         title="Otázka do dvojic", body=[
             "Proč jsou listy zelené, když černá pohltí nejvíce světla?",
             "Co by se stalo s kyslíkem, kdyby zmizely oceánské řasy?",
         ], numbered=True, fill=P.accent_soft, title_color=P.ink, body_color=P.ink,
         title_size_pt=24, body_size_pt=20, pad_pt=28)
```

**Summary / closing (dark)** — three takeaways at 24 pt, then stop.

```python
s = prs.slides.add_slide(BLANK); paper(s, D.paper)
s.shapes.add_text(BBox.from_inches(0.9, 1.4, W - 1.8, 1.0), text="Shrnutí",
                  size_pt=36, bold=True, color=D.ink, margin_pt=0)
add_bullets(s, BBox.from_inches(0.9, 2.6, W - 1.8, 3.6), items=[
    "Rostliny přeměňují světlo, vodu a CO₂ na cukr a kyslík.",
    "Děje se to v chloroplastech díky chlorofylu.",
    "Bez fotosyntézy by nebyl kyslík ani potrava.",
], size_pt=24, color=D.ink)
```

## 11. Look at it before you ship it

Text previews cannot show a clipped title or a lopsided slide. Render the
deck once and look:

```python
for s in prs.slides:
    s.tidy()
prs.save("output/deck.pptx")
prs.render_contact_sheet("preview.png", cols=4, thumb_width=480)   # one PNG, every slide
```

What to check on the sheet, in order:

1. **Fit** — every title on one line, no text touching a box edge, no
   list shrunk to small type.
2. **Rhythm** — titles at the same height on every content slide; the same
   left edge everywhere; dark slides where you intended them.
3. **Density** — no slide markedly fuller than its neighbours; no half-empty
   slide next to a packed one.
4. **Colour** — one accent per slide; nothing outside the palette except
   photos.
5. **Pictures** — undistorted, cropped where intended, captioned.

Fix what you see, re-render once, save. Two passes are enough; polishing
beyond that is time better spent on the words.
