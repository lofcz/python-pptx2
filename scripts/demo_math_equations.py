"""Write a demo deck of native PowerPoint equations and print its path."""

from __future__ import annotations

import sys
from pathlib import Path

from pptx2 import BBox, Presentation
from pptx2.dml.color import RGBColor
from pptx2.enum.text import PP_ALIGN
from pptx2.util import Inches, Pt

INK = RGBColor(0x11, 0x18, 0x27)
MUTED = RGBColor(0x4B, 0x55, 0x63)

FORMULAS = (
    (r"\frac{a}{b}", "Proper fraction"),
    (r"\frac{7}{10}", "Numeric fraction"),
    (r"\frac{a}{b} \in \mathbb{N}", "Fraction in a set"),
    (r"\sqrt{2}", "Square root"),
    (r"\sqrt[3]{8}", "Cube root"),
    (r"\frac{-b \pm \sqrt{b^2-4ac}}{2a}", "Quadratic formula"),
    (r"e^{i\pi}+1=0", "Euler identity"),
    (r"\sum_{k=1}^{n} \frac{1}{k^2}", "N-ary sum + fraction"),
    (r"\int_{0}^{1} \sqrt{1-x^2}\,dx", "Integral + radical"),
    (r"\lim_{x \to 0} \frac{\sin x}{x} = 1", "Limit of a fraction"),
    (
        r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}",
        "2×2 matrix",
    ),
    (r"\overline{AB} \parallel \vec{v}", "Overline + vector"),
)


def _add_label(slide, box: BBox, text: str) -> None:
    shape = slide.shapes.add_textbox(box.left, box.top, box.width, box.height)
    p = shape.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = text
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = MUTED


def main() -> Path:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path.home() / "Desktop" / "pptx2-math-demo.pptx"
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    title = prs.slides.add_slide(prs.slide_layouts[6])
    heading = title.shapes.add_textbox(Inches(0.7), Inches(0.5), Inches(12), Inches(0.7))
    hp = heading.text_frame.paragraphs[0]
    hr = hp.add_run()
    hr.text = "python-pptx2 native equations"
    hr.font.size = Pt(32)
    hr.font.bold = True
    hr.font.color.rgb = INK
    sub = title.shapes.add_textbox(Inches(0.7), Inches(1.2), Inches(12), Inches(0.5))
    sp = sub.text_frame.paragraphs[0]
    sr = sp.add_run()
    sr.text = "Fractions, radicals, n-ary ops, limits, matrices — editable OMML"
    sr.font.size = Pt(16)
    sr.font.color.rgb = MUTED

    cols = 2
    rows = 3
    cell_w, cell_h = 5.8, 1.7
    origin_x, origin_y = 0.7, 1.9
    for i, (latex, label) in enumerate(FORMULAS):
        if i > 0 and i % (cols * rows) == 0:
            title = prs.slides.add_slide(prs.slide_layouts[6])
        slide = title
        idx = i % (cols * rows)
        col, row = idx % cols, idx // cols
        left = origin_x + col * (cell_w + 0.5)
        top = origin_y + row * (cell_h + 0.15)
        _add_label(slide, BBox.from_inches(left, top, cell_w, 0.28), label)
        slide.shapes.add_equation(
            BBox.from_inches(left, top + 0.28, cell_w, cell_h - 0.35),
            latex=latex,
            size_pt=28,
            color="#111827",
            align="left",
        )

    inline = prs.slides.add_slide(prs.slide_layouts[6])
    box = inline.shapes.add_textbox(Inches(0.7), Inches(2.4), Inches(12), Inches(1.4))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    lead = p.add_run()
    lead.text = "Inline mix: compare "
    lead.font.size = Pt(24)
    lead.font.color.rgb = INK
    p.add_math(r"\frac{7}{10}", size_pt=24, color="#0B5CFF")
    mid = p.add_run()
    mid.text = " with "
    mid.font.size = Pt(24)
    mid.font.color.rgb = INK
    p.add_math(r"\frac{4}{10}", size_pt=24, color="#0B5CFF")
    tail = p.add_run()
    tail.text = " — denominators match."
    tail.font.size = Pt(24)
    tail.font.color.rgb = INK

    heading = inline.shapes.add_textbox(Inches(0.7), Inches(0.5), Inches(12), Inches(0.7))
    hp = heading.text_frame.paragraphs[0]
    hr = hp.add_run()
    hr.text = "Inline fractions in a sentence"
    hr.font.size = Pt(32)
    hr.font.bold = True
    hr.font.color.rgb = INK

    prs.save(str(out))
    return out


if __name__ == "__main__":
    path = main()
    print(path)
