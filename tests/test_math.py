"""Native PowerPoint equation (LaTeX → OMML) support."""

from __future__ import annotations

import io
from unittest.mock import patch

import pytest

from pptx2 import BBox, MathBackendUnavailable, Presentation
from pptx2.math import latex_to_omml, office_math_element, strip_math_delimiters
from pptx2.mathml2omml import convert as mathml_to_omml
from pptx2.oxml.ns import qn
from pptx2.util import Inches


def _have_latex2mathml() -> bool:
    try:
        import latex2mathml  # noqa: F401
    except ImportError:
        return False
    return True


needs_math = pytest.mark.skipif(
    not _have_latex2mathml(),
    reason="python-pptx2[math] extra (latex2mathml) not installed",
)

FRAC_MML = (
    '<math xmlns="http://www.w3.org/1998/Math/MathML">'
    "<mfrac><mn>7</mn><mn>10</mn></mfrac></math>"
)
SQRT_MML = (
    '<math xmlns="http://www.w3.org/1998/Math/MathML">'
    "<msqrt><msup><mi>x</mi><mn>2</mn></msup></msqrt></math>"
)


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


class DescribeMathmlToOmml:
    def it_emits_a_spec_fraction(self):
        omml = mathml_to_omml(FRAC_MML)
        assert omml.startswith("<m:oMath")
        assert "<m:f>" in omml
        assert '<m:type m:val="bar"/>' in omml or 'm:val="bar"' in omml
        assert "<m:num>" in omml and "<m:den>" in omml
        assert "<m:t" in omml and ">7<" in omml and ">10<" in omml
        assert "<m:box>" not in omml

    def it_emits_a_hidden_degree_radical(self):
        omml = mathml_to_omml(SQRT_MML)
        assert "<m:rad>" in omml
        assert '<m:degHide m:val="on"/>' in omml or 'm:val="on"' in omml
        assert "<m:sSup>" in omml


class DescribeLatexToOmml:
    def it_rejects_empty_latex(self):
        with pytest.raises(ValueError, match="non-empty"):
            latex_to_omml("   ")

    def it_explains_a_missing_backend(self):
        with patch("pptx2.math._import_latex2mathml") as import_backend:
            import_backend.side_effect = MathBackendUnavailable(
                "pip install 'python-pptx2[math]'"
            )
            with pytest.raises(MathBackendUnavailable, match="python-pptx2\\[math\\]"):
                latex_to_omml(r"x")

    @needs_math
    def it_emits_an_omath_fraction(self):
        omml = latex_to_omml(r"\frac{a}{b}")
        assert omml.startswith("<m:oMath")
        assert "<m:f>" in omml
        assert "<m:fPr>" in omml
        assert "<m:num>" in omml and "<m:den>" in omml

    @needs_math
    def it_puts_a_summand_inside_nary_e(self):
        omml = latex_to_omml(r"\sum_{k=1}^{n} \frac{1}{k^2}")
        nary = omml[omml.index("<m:nary>") : omml.index("</m:nary>") + len("</m:nary>")]
        assert "<m:f>" in nary
        assert "<m:e/>" not in nary


@pytest.fixture
def slide():
    prs = Presentation()
    return prs.slides.add_slide(prs.slide_layouts[6])


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
        assert "<m:f>" in xml or ":f>" in xml
        assert 'sz="2800"' in xml
        assert 'val="0B5CFF"' in xml
        assert shape.name.startswith("Equation")

    def it_accepts_positional_lengths(self, slide):
        shape = slide.shapes.add_equation(
            Inches(1),
            Inches(2),
            Inches(6),
            Inches(1),
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
    def it_survives_save_and_open_with_fraction_structure(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_equation(
            BBox.from_inches(1, 1, 8, 1),
            latex=r"\frac{7}{10}",
        )
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)
        xml = reopened.slides[0].shapes[-1]._element.xml
        assert "<m:f>" in xml or ":f>" in xml
        assert "<m:num>" in xml and "<m:den>" in xml
        texts = [t.text for t in reopened.slides[0].shapes[-1]._element.iter(qn("m:t"))]
        assert "7" in texts and "10" in texts
        assert "710" not in texts

    def it_declares_math_host_namespaces_on_the_slide(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_equation(
            BBox.from_inches(1, 1, 8, 1),
            latex=r"\sqrt{2}",
        )
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        import zipfile

        with zipfile.ZipFile(buf) as zf:
            raw = zf.read("ppt/slides/slide1.xml")
        assert b'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"' in raw
        assert b'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"' in raw
        assert b'mc:Ignorable="' in raw
        assert b"a14" in raw
