"""Native PowerPoint equation (LaTeX → OMML) support."""

from __future__ import annotations

import io
from unittest.mock import patch

import pytest

from pptx2 import BBox, MathBackendUnavailable, Presentation
from pptx2.math import latex_to_omml, office_math_element, strip_math_delimiters
from pptx2.oxml.ns import qn
from pptx2.util import Inches


def _have_math_backend() -> bool:
    try:
        import latex2mathml  # noqa: F401
        import mathml2omml  # noqa: F401
    except ImportError:
        return False
    return True


needs_math = pytest.mark.skipif(
    not _have_math_backend(),
    reason="python-pptx2[math] extras (latex2mathml, mathml2omml) not installed",
)


@pytest.fixture
def slide():
    prs = Presentation()
    return prs.slides.add_slide(prs.slide_layouts[6])


class DescribeStripMathDelimiters:
    def it_unwraps_common_fences(self):
        assert strip_math_delimiters(r"$\frac{a}{b}$") == r"\frac{a}{b}"
        assert strip_math_delimiters(r"$$\frac{a}{b}$$") == r"\frac{a}{b}"
        assert strip_math_delimiters(r"\(\frac{a}{b}\)") == r"\frac{a}{b}"
        assert strip_math_delimiters(r"\[\frac{a}{b}\]") == r"\frac{a}{b}"
        assert (
            strip_math_delimiters(r"\begin{equation}E=mc^2\end{equation}") == "E=mc^2"
        )

    def it_leaves_a_bare_fragment_alone(self):
        assert strip_math_delimiters(r"\frac{a}{b}") == r"\frac{a}{b}"


class DescribeLatexToOmml:
    def it_rejects_empty_latex(self):
        with pytest.raises(ValueError, match="non-empty"):
            latex_to_omml("   ")

    def it_explains_a_missing_backend(self):
        with patch("pptx2.math._import_math_backends") as import_backends:
            import_backends.side_effect = MathBackendUnavailable(
                "pip install 'python-pptx2[math]'"
            )
            with pytest.raises(MathBackendUnavailable, match="python-pptx2\\[math\\]"):
                latex_to_omml(r"x")

    @needs_math
    def it_emits_an_omath_fragment(self):
        omml = latex_to_omml(r"\frac{a}{b}")
        assert omml.startswith("<m:oMath")
        assert "m:f" in omml or ":f>" in omml


@needs_math
class DescribeOfficeMathElement:
    def it_wraps_display_math_in_omathpara(self):
        marker = office_math_element(r"E=mc^2", display=True, align="center")
        assert marker.tag == qn("a14:m")
        assert marker.find(qn("m:oMathPara")) is not None
        jc = marker.find(".//%s" % qn("m:jc"))
        assert jc is not None
        assert jc.get(qn("m:val")) == "centerGroup"

    def it_keeps_inline_math_as_omath(self):
        marker = office_math_element(r"E=mc^2", display=False)
        assert marker.find(qn("m:oMathPara")) is None
        assert marker.find(qn("m:oMath")) is not None


@needs_math
class DescribeAddEquation:
    def it_writes_a14_math_into_the_shape(self, slide):
        shape = slide.shapes.add_equation(
            BBox.from_inches(1, 2, 8, 1.5),
            latex=r"\frac{a}{b}",
            size_pt=28,
            color="#0B5CFF",
        )
        xml = shape._element.xml
        assert "oMath" in xml
        assert 'sz="2800"' in xml
        assert 'val="0B5CFF"' in xml
        assert shape.name.startswith("Equation")

    def it_accepts_positional_lengths(self, slide):
        shape = slide.shapes.add_equation(
            Inches(1), Inches(2), Inches(6), Inches(1),
            latex=r"E=mc^2",
        )
        assert "oMath" in shape._element.xml

    def it_rejects_a_bad_bbox_call(self, slide):
        with pytest.raises(TypeError, match="BBox"):
            slide.shapes.add_equation(Inches(1), latex=r"x")


@needs_math
class DescribeAddMath:
    def it_mixes_runs_and_inline_math(self, slide):
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        paragraph = box.text_frame.paragraphs[0]
        paragraph.add_run().text = "Euler: "
        paragraph.add_math(r"e^{i\pi}+1=0")
        paragraph.add_run().text = "."
        xml = paragraph._element.xml
        assert "oMath" in xml
        assert paragraph.runs[0].text == "Euler: "
        assert paragraph.runs[1].text == "."

    def it_clears_math_with_the_paragraph(self, slide):
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        paragraph = box.text_frame.paragraphs[0]
        paragraph.add_math(r"x^2")
        assert paragraph._element.find(qn("a14:m")) is not None
        paragraph.clear()
        assert paragraph._element.find(qn("a14:m")) is None


@needs_math
class DescribeEquationRoundTrip:
    def it_survives_save_and_open(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_equation(
            BBox.from_inches(1, 1, 8, 1),
            latex=r"\sum_{i=1}^{n} i",
        )
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)
        xml = reopened.slides[0].shapes[-1]._element.xml
        assert "oMath" in xml
        assert qn("a14:m") in xml or "a14:m" in xml
