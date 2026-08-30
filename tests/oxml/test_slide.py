"""Unit-test suite for `pptx2.oxml.slide` module."""

from __future__ import annotations

from pptx2.oxml.ns import qn
from pptx2.oxml.slide import CT_NotesMaster, CT_NotesSlide, CT_SlideLayout

from ..unitutil.cxml import element, xml
from ..unitutil.file import snippet_text


class DescribeCT_NotesMaster(object):
    """Unit-test suite for `pptx2.oxml.slide.CT_NotesMaster` objects."""

    def it_can_create_a_default_notesMaster_element(self):
        notesMaster = CT_NotesMaster.new_default()
        assert notesMaster.xml == snippet_text("default-notesMaster")


class DescribeCT_NotesSlide(object):
    """Unit-test suite for `pptx2.oxml.slide.CT_NotesSlide` objects."""

    def it_can_create_a_new_notes_element(self):
        notes = CT_NotesSlide.new()
        assert notes.xml == snippet_text("default-notes")


class DescribeCT_SlideLayout(object):
    """Unit-test suite for `pptx2.oxml.slide.CT_SlideLayout` objects."""

    def it_can_create_a_new_sldLayout_element(self):
        sldLayout = CT_SlideLayout.new()

        assert sldLayout.tag == qn("p:sldLayout")
        assert sldLayout.cSld.spTree is not None
        assert sldLayout.xpath("p:cSld/p:spTree/p:nvGrpSpPr/p:cNvPr")[0].get("id") == "1"
        assert sldLayout.xpath("p:clrMapOvr/a:masterClrMapping")


class DescribeCT_SlideLayoutIdList(object):
    """Unit-test suite for `pptx2.oxml.slide.CT_SlideLayoutIdList` objects."""

    def it_can_add_a_sldLayoutId_with_an_allocatable_id(self):
        sldLayoutIdLst = element("p:sldLayoutIdLst/p:sldLayoutId{id=2147483649,r:id=rId1}")

        entry = sldLayoutIdLst.add_sldLayoutId("rId2")

        assert entry.rId == "rId2"
        assert entry.id == 2147483650
        assert sldLayoutIdLst.xml == xml(
            "p:sldLayoutIdLst/("
            "p:sldLayoutId{id=2147483649,r:id=rId1},"
            "p:sldLayoutId{id=2147483650,r:id=rId2})"
        )

    def it_allocates_the_spec_minimum_id_for_the_first_layout(self):
        sldLayoutIdLst = element("p:sldLayoutIdLst")

        assert sldLayoutIdLst.add_sldLayoutId("rId1").id == 2147483648

    def it_can_add_a_sldLayoutId_with_an_explicit_id(self):
        sldLayoutIdLst = element("p:sldLayoutIdLst/p:sldLayoutId{id=2147483649,r:id=rId1}")

        entry = sldLayoutIdLst.add_sldLayoutId("rId2", 2147483655)

        assert (entry.id, entry.rId) == (2147483655, "rId2")
