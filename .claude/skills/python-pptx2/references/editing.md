# Editing existing decks: inspect, edit, diff, batch (paper-pptx)

Ported from paper-pptx. These APIs exist to change an *existing* deck
without silently corrupting it: every operation validates fully before
the first mutation, returns typed machine-readable outcomes, and a
refusal leaves the package byte-identical.

```python
from pptx2 import Presentation
from pptx2.inspect import inspect_text, inspect_deck, effective_font
from pptx2.edit import replace_text
from pptx2.diff import diff_decks
```

## Reading what the deck actually renders

```python
prs = Presentation("deck.pptx")

# resolved size/typeface/color with provenance (run -> paragraph ->
# shape list style -> placeholder/layout/master -> theme)
font = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0].effective_font()
font.size.value_pt          # e.g. 18.0 — resolved, never None-guessed
font.size.provenance        # which rung of the chain supplied it

# every text block, including nested groups and table cells
report = inspect_text(prs.slides[0])
for block in report.blocks:
    block.text, block.anchor, block.fields

# deterministic structural manifest of the whole deck
manifest = inspect_deck(prs)
```

Also: `pptx2.inspect.effective_paragraph_format(paragraph)` (alignment,
line spacing, and the *inherited bullet* with its own typeface/size
chains) and `effective_shape_format(shape)`.

## Editing text without flattening formatting

```python
# replaces every "FY25" with "FY26", keeping each run's formatting;
# runs fragment around the match and untouched runs are left alone
replace_text(prs, "FY25", "FY26")
```

Edits are anchored by structural identity (shape/table-cell anchor plus
content fingerprint), so duplicated text cannot send an edit to the
wrong twin. Invalid anchors raise typed refusals
(`pptx2.errors.UnsupportedStructureError` etc.), never guess.

## Slide-level surgery

```python
prs.slides.clone(3)                 # policy-governed deep copy
prs.slides.delete(2)                # sections/custom shows maintained
prs.slides.move(slide, 0)           # or move(2, 0) by index
prs.slides.reorder([2, 0, 3, 1])

slide.shapes.add_copy(shape)        # relationship-safe shape copy
slide.shapes.delete(shape)          # drops now-unreferenced rels
slide.shapes.shape_by_name("Total") # strict: unique or AmbiguousTargetError
slide.shapes.chart_by_name("rev")   # / picture_by_name / table_by_name

slide.rebind_layout(new_layout)     # template migration primitive
slide.apply_footers(footer="ACME", slide_number=True)
slide.read_notes_text()             # never creates a notes slide
slide.replace_notes_text("new note")
```

## Tables, pictures, charts

```python
table = slide.shapes.table_by_name("Q3")
table.insert_row(0, copy_format_from=0)     # clone formatting, not text
table.insert_column(2, copy_format_from=1)
table.delete_row(5); table.delete_column(0) # merge-aware refusals
cell.extend_merge(other_cell)               # grow a rectangular merge

picture.replace_image("new.png")            # geometry preserved
chart.replace_data_safe(["A", "B"], [("S1", (1.0, 2.0))])

tf.normalize_autofit(min_font_size=Pt(11))  # freeze normAutofit shrink
tf.font_scale                                # read the recorded scale
paragraph.bullet.set_character(char="•")    # real a:buChar markup
paragraph.add_slide_number_field()          # real a:fld, not a literal
```

## Diffing two decks

```python
from pptx2.diff import diff_decks

delta = diff_decks("old.pptx", "new.pptx", detail="text")
delta.slide_changes    # matched by stable identity, not position
delta.package_changes  # part-level adds/removes/changes
```

## Batch: validate once per block

```python
with prs.batch():
    for slide in prs.slides:
        slide.shapes[0].text_frame.text = "..."
# deck validated once at exit; a refusal rolls back the whole block.
# save() inside the block raises BoundaryViolationError.
prs.save("out.pptx")
```

## Error contract

All refusals derive from `pptx2.errors.PaperRefusal`
(`UnsupportedStructureError`, `TargetNotFoundError`,
`AmbiguousTargetError`, `RelationshipPolicyError`,
`PackageLimitError`, `BoundaryViolationError`, ...). Catch
`PaperRefusal` at the boundary; a refusal never leaves a half-applied
edit.
