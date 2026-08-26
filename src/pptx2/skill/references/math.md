# Native equations from LaTeX

PowerPoint stores editable equations as Office Math (OMML) inside an
`a14:m` marker. python-pptx2 does **not** compile LaTeX itself. It calls
two libraries and wraps the result:

1. [`latex2mathml`](https://pypi.org/project/latex2mathml/) — LaTeX → MathML
2. [`mathml2omml`](https://pypi.org/project/mathml2omml/) — MathML → OMML

Install them with:

```bash
pip install "python-pptx2[math]"
```

Missing converters raise `MathBackendUnavailable` with that install line.

## Display equation — `slide.shapes.add_equation`

Same calling convention as `add_text`: a `BBox` or `(left, top, width, height)`.

```python
from pptx2 import Presentation, BBox
from pptx2.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

slide.shapes.add_equation(
    BBox.from_inches(1, 2, 8, 1.5),
    latex=r"\frac{-b \pm \sqrt{b^2-4ac}}{2a}",
    size_pt=28,
    color="#111827",
    align="center",
)

slide.shapes.add_equation(
    Inches(1), Inches(4), Inches(8), Inches(1),
    latex=r"E = mc^2",
)
```

The shape is a text box. The equation is editable in PowerPoint's
equation editor. `display=True` (the default) wraps OMML in
`m:oMathPara`.

Keyword args: `latex` (required), `display`, `font`, `size_pt`,
`color`, `align`, `anchor`, `margin_pt`.

## Inline equation — `paragraph.add_math`

Sits between ordinary runs in the same paragraph:

```python
box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
p = box.text_frame.paragraphs[0]
p.add_run().text = "Euler's identity "
p.add_math(r"e^{i\pi} + 1 = 0")
p.add_run().text = " holds for all real π."
```

`display=False` (the default here) emits inline `m:oMath`. Pass
`size_pt` / `color` / `font` to style the math runs.

`paragraph.clear()` and assigning `paragraph.text` remove the equation
along with the runs.

## Bare converter

```python
from pptx2.math import latex_to_omml

omml = latex_to_omml(r"\sum_{i=1}^{n} i")
# '<m:oMath>…</m:oMath>'
```

Wrappers (`$…$`, `$$…$$`, `\(…\)`, `\[…\]`, `equation` / `align`
environments) are stripped before conversion.

## What this is not

- Not a TeX engine. Environments like `tikzpicture` will not render.
- Not Microsoft's `MML2OMML.XSL` (that stylesheet cannot be
  redistributed). The MathML → OMML step is the `mathml2omml` package.
- Not an image. If you need a PNG of a formula, use
  `add_matplotlib_figure` with mathtext instead.
