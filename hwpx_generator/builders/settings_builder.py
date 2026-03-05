"""Generate settings.xml for HWPX document."""

from __future__ import annotations

from typing import TYPE_CHECKING
from xml.etree.ElementTree import Element, SubElement, tostring

from ..utils import mm_to_hwp

if TYPE_CHECKING:
    from ..document import HwpxDocument

# Namespace constants
NS_HS = "http://www.hancom.co.kr/hwpml/2011/section"
NS_HC = "http://www.hancom.co.kr/hwpml/2011/core"


def build_settings_xml(doc: HwpxDocument) -> bytes:
    """Build settings.xml content."""
    ps = doc.page_settings

    root = Element("hs:settings")
    root.set("xmlns:hs", NS_HS)
    root.set("xmlns:hc", NS_HC)

    # <hs:beginNumber page="1" footnote="1" endnote="1"/>
    begin = SubElement(root, "hs:beginNumber")
    begin.set("page", "1")
    begin.set("footnote", "1")
    begin.set("endnote", "1")

    return _xml_declaration() + tostring(root, encoding="unicode").encode("utf-8")


def _xml_declaration() -> bytes:
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
