# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx2.parts.customprops` module."""

from __future__ import annotations

import io
from typing import Any

import pytest

from pptx2 import Presentation
from pptx2.opc.constants import CONTENT_TYPE as CT
from pptx2.opc.constants import RELATIONSHIP_TYPE as RT
from pptx2.oxml.customprops import CT_CustomProperties
from pptx2.parts.customprops import CustomProperties, CustomPropertiesPart

_CUSTOM_PROPS_XML = (
    b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
    b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
    b'custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/'
    b'2006/docPropsVTypes">'
    b'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Sponsor">'
    b"<vt:lpstr>Acme Corp</vt:lpstr></property>"
    b'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Quarter">'
    b"<vt:lpwstr>Q3</vt:lpwstr></property>"
    b'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Revision">'
    b"<vt:i4>7</vt:i4></property>"
    b'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="Score">'
    b"<vt:r8>3.25</vt:r8></property>"
    b'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="Confidential">'
    b"<vt:bool>true</vt:bool></property>"
    b'<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="Legacy">'
    b"<vt:date>2024-01-01T00:00:00Z</vt:date></property>"
    b"</Properties>\n"
)

_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"


class DescribeCustomPropertiesPart(object):
    """Unit-test suite for `pptx2.parts.customprops.CustomPropertiesPart` objects."""

    def it_can_construct_a_default_custom_props(self):
        custom_props = CustomPropertiesPart.default(None)  # type: ignore

        assert isinstance(custom_props, CustomPropertiesPart)
        assert custom_props.content_type is CT.OFC_CUSTOM_PROPERTIES
        assert custom_props.partname == "/docProps/custom.xml"
        assert isinstance(custom_props._element, CT_CustomProperties)
        assert len(custom_props._element) == 0

    @pytest.mark.parametrize(
        ("name", "expected_value", "expected_type"),
        [
            ("Sponsor", "Acme Corp", str),
            ("Quarter", "Q3", str),
            ("Revision", 7, int),
            ("Score", 3.25, float),
            ("Confidential", True, bool),
            # -- a VT type this library never writes still reads back (as str) --
            ("Legacy", "2024-01-01T00:00:00Z", str),
        ],
    )
    def it_knows_the_property_values(self, custom_props, name, expected_value, expected_type):
        value = custom_props[name]

        assert value == expected_value
        assert type(value) is expected_type

    def it_knows_the_property_names(self, custom_props):
        assert custom_props.keys() == [
            "Sponsor",
            "Quarter",
            "Revision",
            "Score",
            "Confidential",
            "Legacy",
        ]
        assert list(custom_props) == custom_props.keys()
        assert len(custom_props) == 6
        assert "Sponsor" in custom_props
        assert "Nope" not in custom_props

    def it_raises_keyerror_for_a_missing_property(self, custom_props):
        with pytest.raises(KeyError):
            custom_props["Nope"]

    @pytest.mark.parametrize(
        ("value", "vt_tag", "lexical"),
        [
            ("Acme Corp", "vt:lpstr", "Acme Corp"),
            (7, "vt:i4", "7"),
            (-42, "vt:i4", "-42"),
            (3.25, "vt:r8", "3.25"),
            (True, "vt:bool", "true"),
            (False, "vt:bool", "false"),
        ],
    )
    def it_can_set_each_value_type(self, value, vt_tag: str, lexical: str):
        custom_props = CustomPropertiesPart(None, None, None, CT_CustomProperties.new())  # type: ignore

        custom_props["Sponsor"] = value

        assert custom_props._element.xml == (
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
            'custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/'
            '2006/docPropsVTypes">\n'
            '  <property fmtid="%s" pid="2" name="Sponsor">\n'
            "    <%s>%s</%s>\n"
            "  </property>\n"
            "</Properties>\n" % (_FMTID, vt_tag, lexical, vt_tag)
        )
        assert custom_props["Sponsor"] == value
        assert type(custom_props["Sponsor"]) is type(value)

    def it_can_update_a_property_in_place(self, custom_props):
        custom_props["Revision"] = 8

        assert custom_props["Revision"] == 8
        assert len(custom_props) == 6
        # -- the update reuses the existing property element, pid included --
        prop = custom_props._element.property_by_name("Revision")
        assert prop.get("pid") == "4"

    def it_retypes_a_property_when_the_value_type_changes(self, custom_props):
        custom_props["Sponsor"] = 99

        assert custom_props["Sponsor"] == 99
        children = custom_props._element.property_by_name("Sponsor")
        assert (
            children[0].tag
            == "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}i4"
        )

    def it_can_delete_a_property(self, custom_props):
        del custom_props["Sponsor"]

        assert "Sponsor" not in custom_props
        assert custom_props.keys() == ["Quarter", "Revision", "Score", "Confidential", "Legacy"]
        with pytest.raises(KeyError):
            del custom_props["Sponsor"]

    def it_can_delete_the_last_property(self, custom_props):
        for name in list(custom_props):
            del custom_props[name]

        assert len(custom_props) == 0
        assert custom_props.keys() == []
        # -- the (now empty) element is still present and well-formed --
        assert len(custom_props._element) == 0
        assert custom_props._element.tag == CT_CustomProperties.new().tag

    def it_assignes_sequential_pids_starting_at_two(self):
        custom_props = CustomPropertiesPart(None, None, None, CT_CustomProperties.new())  # type: ignore

        custom_props["A"] = 1
        custom_props["B"] = 2

        pids = [prop.get("pid") for prop in custom_props._element.property_elms]
        assert pids == ["2", "3"]

    def it_does_not_reuse_a_pid_after_a_deletion(self):
        custom_props = CustomPropertiesPart(None, None, None, CT_CustomProperties.new())  # type: ignore

        custom_props["A"] = 1
        custom_props["B"] = 2
        del custom_props["A"]
        custom_props["C"] = 3

        pids = {prop.get("name"): prop.get("pid") for prop in custom_props._element.property_elms}
        assert pids == {"B": "3", "C": "4"}

    def it_writes_the_fmtid_on_each_property(self, custom_props):
        for prop in custom_props._element.property_elms:
            assert prop.get("fmtid") == _FMTID

    @pytest.mark.parametrize("value", [None, 1.5j, [1, 2], {"a": 1}, object()])
    def it_rejects_unsupported_value_types(self, value: Any):
        custom_props = CustomPropertiesPart(None, None, None, CT_CustomProperties.new())  # type: ignore

        with pytest.raises(TypeError):
            custom_props["N"] = value

    def it_rejects_int_values_that_do_not_fit_vt_i4(self):
        custom_props = CustomPropertiesPart(None, None, None, CT_CustomProperties.new())  # type: ignore

        with pytest.raises(ValueError):
            custom_props["N"] = 2**31
        with pytest.raises(ValueError):
            custom_props["N"] = -(2**31) - 1

    @pytest.mark.parametrize("name", ["", None, 42])
    def it_rejects_invalid_property_names(self, name: Any):
        custom_props = CustomPropertiesPart(None, None, None, CT_CustomProperties.new())  # type: ignore

        with pytest.raises((TypeError, ValueError)):
            custom_props[name] = "value"

    # -- fixtures ----------------------------------------------------

    @pytest.fixture
    def custom_props(self) -> CustomPropertiesPart:
        return CustomPropertiesPart.load(None, None, None, _CUSTOM_PROPS_XML)  # type: ignore


class DescribeCustomProperties(object):
    """Unit-test suite for the `CustomProperties` package-level mapping proxy."""

    def it_is_empty_for_a_fresh_package(self, fresh_properties):
        assert len(fresh_properties) == 0
        assert list(fresh_properties) == []
        assert fresh_properties.keys() == []
        assert "Sponsor" not in fresh_properties

    def it_raises_keyerror_for_a_missing_property(self, fresh_properties):
        with pytest.raises(KeyError):
            fresh_properties["Sponsor"]
        with pytest.raises(KeyError):
            del fresh_properties["Sponsor"]

    def it_does_not_create_a_part_on_read(self, prs, fresh_properties):
        "Sponsor" in fresh_properties
        list(fresh_properties)
        len(fresh_properties)

        partnames = [part.partname for part in prs.part.package.iter_parts()]
        assert "/docProps/custom.xml" not in partnames

    def it_creates_the_part_and_relationship_on_first_assignment(self, prs, fresh_properties):
        package = prs.part.package

        fresh_properties["Sponsor"] = "Acme Corp"

        custom_props_part = package.part_related_by(RT.CUSTOM_PROPERTIES)
        assert isinstance(custom_props_part, CustomPropertiesPart)
        assert custom_props_part.partname == "/docProps/custom.xml"
        assert custom_props_part.content_type == CT.OFC_CUSTOM_PROPERTIES

    def it_round_trips_values_and_types_through_save(self, prs, fresh_properties):
        fresh_properties["Sponsor"] = "Acme Corp"
        fresh_properties["Revision"] = 7
        fresh_properties["Score"] = 3.25
        fresh_properties["Confidential"] = True

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        reopened = Presentation(buf)
        custom_properties = reopened.custom_properties

        assert custom_properties.keys() == ["Sponsor", "Revision", "Score", "Confidential"]
        assert custom_properties["Sponsor"] == "Acme Corp"
        assert custom_properties["Revision"] == 7
        assert custom_properties["Score"] == 3.25
        assert custom_properties["Confidential"] is True
        assert type(custom_properties["Sponsor"]) is str
        assert type(custom_properties["Revision"]) is int
        assert type(custom_properties["Score"]) is float
        assert type(custom_properties["Confidential"]) is bool

    def it_is_a_stable_instance_per_package(self, prs):
        assert prs.custom_properties is prs.custom_properties
        assert isinstance(prs.custom_properties, CustomProperties)

    # -- fixtures ----------------------------------------------------

    @pytest.fixture
    def prs(self):
        return Presentation()

    @pytest.fixture
    def fresh_properties(self, prs) -> CustomProperties:
        return prs.custom_properties
