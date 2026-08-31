"""End-to-end tests for `os.PathLike` inputs at public entry points.

A `pathlib.Path` is accepted anywhere a file path (str) was accepted, with no
behavior change for str or file-like inputs.
"""

from __future__ import annotations

import io
import shutil
from pathlib import Path

import pytest

from pptx2 import Presentation
from pptx2.parts.image import Image

from .unitutil.file import absjoin, test_file_dir

test_image_path = absjoin(test_file_dir, "python-icon.jpeg")

#: Index of the default-template layout containing a picture placeholder.
PICTURE_LAYOUT_IDX = 8


class DescribePackageInputs:
    """`Presentation()` and `.save()` accept `pathlib.Path` as well as str/stream."""

    def it_can_open_a_presentation_from_a_path(self, pptx_file):
        prs = Presentation(pptx_file)

        assert prs.slide_width is not None

    def it_can_open_a_presentation_from_a_str_path_too(self, pptx_file):
        prs = Presentation(str(pptx_file))

        assert prs.slide_width is not None

    def it_can_open_a_presentation_from_a_stream_too(self, pptx_file):
        with open(pptx_file, "rb") as f:
            prs = Presentation(f)

        assert prs.slide_width is not None

    def it_can_save_a_presentation_to_a_path(self, tmp_path):
        pkg_path = tmp_path / "saved.pptx"
        prs = Presentation()

        prs.save(pkg_path)

        reopened = Presentation(pkg_path)
        assert len(reopened.slides) == len(prs.slides)

    def it_can_save_a_presentation_to_a_str_path_too(self, tmp_path):
        pkg_path = tmp_path / "saved.pptx"
        prs = Presentation()

        prs.save(str(pkg_path))

        reopened = Presentation(str(pkg_path))
        assert len(reopened.slides) == len(prs.slides)

    def it_can_save_a_presentation_to_a_stream_too(self, tmp_path):
        pkg_path = tmp_path / "saved.pptx"
        prs = Presentation()

        # -- "w+b" until the paper-pptx save-destination fix (upstream #23)
        # -- lands: the bootstrap hardening requires a stream it could rewind --
        with open(pkg_path, "w+b") as f:
            prs.save(f)

        reopened = Presentation(pkg_path)
        assert len(reopened.slides) == len(prs.slides)

    # --- fixture components -----------------------------------------

    @pytest.fixture
    def pptx_file(self, tmp_path) -> Path:
        """Return path of a small .pptx file written into `tmp_path`."""
        pkg_path = tmp_path / "input.pptx"
        stream = io.BytesIO()
        Presentation().save(stream)
        pkg_path.write_bytes(stream.getvalue())
        return pkg_path


class DescribeImageInputs:
    """`Image.from_file()` accepts `pathlib.Path` as well as str/stream."""

    def it_can_construct_an_image_from_a_path(self, image_file, image_bytes):
        image = Image.from_file(image_file)

        assert image.filename == image_file.name
        assert image.blob == image_bytes

    def it_can_construct_an_image_from_a_str_path_too(self, image_file, image_bytes):
        image = Image.from_file(str(image_file))

        assert image.filename == image_file.name
        assert image.blob == image_bytes

    def it_can_construct_an_image_from_a_stream_too(self, image_bytes):
        with open(test_image_path, "rb") as f:
            image = Image.from_file(f)

        assert image.filename is None
        assert image.blob == image_bytes


class DescribeAddPicture:
    """`shapes.add_picture()` accepts `pathlib.Path` as well as str/stream."""

    def it_can_add_a_picture_from_a_path(self, image_file, image_bytes):
        slide = _blank_slide()

        picture = slide.shapes.add_picture(image_file)

        assert picture.image.blob == image_bytes

    def it_can_add_a_picture_from_a_str_path_too(self, image_file, image_bytes):
        slide = _blank_slide()

        picture = slide.shapes.add_picture(str(image_file))

        assert picture.image.blob == image_bytes

    def it_can_add_a_picture_from_a_stream_too(self, image_bytes):
        slide = _blank_slide()
        with open(test_image_path, "rb") as f:
            picture = slide.shapes.add_picture(f)

        assert picture.image.blob == image_bytes


class DescribeInsertPicture:
    """Picture-placeholder `insert_picture()` accepts `pathlib.Path`."""

    def it_can_insert_a_picture_from_a_path(self, image_file, image_bytes):
        placeholder = _picture_placeholder()

        placeholder_picture = placeholder.insert_picture(image_file)

        assert placeholder_picture.image.blob == image_bytes

    def it_can_insert_a_picture_from_a_str_path_too(self, image_file, image_bytes):
        placeholder = _picture_placeholder()

        placeholder_picture = placeholder.insert_picture(str(image_file))

        assert placeholder_picture.image.blob == image_bytes


# --- helpers -------------------------------------------------------


def _blank_slide():
    """Return a new blank-content slide of a fresh default presentation."""
    prs = Presentation()
    return prs.slides.add_slide(prs.slide_layouts[6])


def _picture_placeholder():
    """Return the picture placeholder (idx 1) of a 'Picture with Caption' slide."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[PICTURE_LAYOUT_IDX])
    return slide.placeholders[1]


@pytest.fixture
def image_file(tmp_path) -> Path:
    """Return path of a copy of the test image renamed within `tmp_path`."""
    image_path = tmp_path / "a-picture.jpeg"
    shutil.copy(test_image_path, image_path)
    return image_path


@pytest.fixture
def image_bytes() -> bytes:
    with open(test_image_path, "rb") as f:
        return f.read()
