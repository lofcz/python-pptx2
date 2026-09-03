"""Unit tests for ``pptx2.render.render_contact_sheet`` (LibreOffice mocked)."""

from __future__ import annotations

from pathlib import Path
from subprocess import CompletedProcess
from unittest.mock import patch

import pytest
from PIL import Image

from pptx2 import Presentation
from pptx2.render import ThumbnailRendererUnavailable, render_contact_sheet


@pytest.fixture
def five_slide_prs():
    prs = Presentation()
    for _ in range(5):
        prs.slides.add_slide(prs.slide_layouts[6])
    return prs


def _fake_soffice_run(num_slides: int, size=(320, 180)):
    def _runner(soffice_bin, deck_path, out_dir, timeout):
        for i in range(num_slides):
            Image.new("RGB", size, (200 + i * 10, 200, 200)).save(Path(out_dir) / f"slide{i}.png")
        return CompletedProcess(args=[], returncode=0, stdout=b"", stderr=b"")

    return _runner


class DescribeRenderContactSheet:
    def it_tiles_every_slide_into_one_png(self, tmp_path, five_slide_prs):
        out = tmp_path / "sheet.png"
        with (
            patch("pptx2.render.shutil.which", return_value="/usr/bin/soffice"),
            patch("pptx2.render._run_soffice", side_effect=_fake_soffice_run(5)),
        ):
            written = render_contact_sheet(five_slide_prs, out, cols=3, thumb_width=160, gap=10)
        assert written == out
        assert out.is_file()
        with Image.open(out) as sheet:
            # 3 columns x 2 rows of 160x90 thumbs with 10px gaps.
            assert sheet.size == (10 + 3 * 170, 10 + 2 * 100)

    def it_respects_a_slide_subset(self, tmp_path, five_slide_prs):
        out = tmp_path / "subset.png"
        with (
            patch("pptx2.render.shutil.which", return_value="/usr/bin/soffice"),
            patch("pptx2.render._run_soffice", side_effect=_fake_soffice_run(5)),
        ):
            render_contact_sheet(five_slide_prs, out, slides=[0, 4], cols=2, thumb_width=100, gap=0)
        with Image.open(out) as sheet:
            assert sheet.size == (200, 56)

    def it_is_reachable_from_the_presentation(self, tmp_path, five_slide_prs):
        out = tmp_path / "via-prs.png"
        with (
            patch("pptx2.render.shutil.which", return_value="/usr/bin/soffice"),
            patch("pptx2.render._run_soffice", side_effect=_fake_soffice_run(5)),
        ):
            five_slide_prs.render_contact_sheet(out, cols=5, thumb_width=64)
        assert out.is_file()

    def it_raises_unavailable_when_soffice_is_missing(self, tmp_path, five_slide_prs):
        with (
            patch("pptx2.render.shutil.which", return_value=None),
            patch.dict("os.environ", {}, clear=False),
        ):
            import os

            os.environ.pop("POWER_PPTX_SOFFICE", None)
            with pytest.raises(ThumbnailRendererUnavailable):
                render_contact_sheet(five_slide_prs, tmp_path / "x.png")

    def it_validates_arguments(self, tmp_path, five_slide_prs):
        with pytest.raises(ValueError):
            render_contact_sheet(five_slide_prs, tmp_path / "x.png", cols=0)
        with pytest.raises(ValueError):
            render_contact_sheet(five_slide_prs, tmp_path / "x.png", thumb_width=8)
