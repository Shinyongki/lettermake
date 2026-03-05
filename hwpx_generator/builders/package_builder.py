"""Package all XML files into a .hwpx ZIP archive."""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import TYPE_CHECKING

from .styles_builder import build_header_xml
from .xml_builder import build_section_xml

if TYPE_CHECKING:
    from ..document import HwpxDocument


def _make_zipinfo(filename: str, compress: int = zipfile.ZIP_DEFLATED) -> zipfile.ZipInfo:
    """Create ZipInfo matching Hancom's expected format."""
    info = zipfile.ZipInfo(filename)
    info.compress_type = compress
    info.create_system = 11          # NTFS (Windows)
    info.external_attr = 0x81800020  # match sample
    info.flag_bits = 4 if compress == zipfile.ZIP_DEFLATED else 0
    return info


def build_hwpx_package(doc: HwpxDocument, output_path: Path) -> None:
    """Build the .hwpx ZIP package and write to disk."""
    from ..elements.image import Image
    from ..elements.chart import Chart
    from ..elements.doc_title import DocTitle
    from .chart_builder import build_chart_xml

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Collect image blocks for BinData packaging
    image_blocks = [b for b in doc.blocks if isinstance(b, Image)]

    # Collect DocTitle blocks for decoration image packaging
    doc_title_blocks = [b for b in doc.blocks if isinstance(b, DocTitle)]

    # Collect chart blocks and assign IDs
    chart_blocks = [b for b in doc.blocks if isinstance(b, Chart)]
    for i, chart in enumerate(chart_blocks):
        chart.chart_id = i + 1

    # Build all_bin_items for content.hpf (images + doc_title decorations)
    all_bin_items = []
    for img in image_blocks:
        all_bin_items.append((f"image{img._bin_id}", img.bin_path, "image/png" if img.format == "png" else "image/jpeg"))
    for dt in doc_title_blocks:
        all_bin_items.append((f"image{dt._top_bin_id}", dt.top_bin_path, "image/png"))
        all_bin_items.append((f"image{dt._bottom_bin_id}", dt.bottom_bin_path, "image/png"))

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        # mimetype — MUST be first entry, stored uncompressed
        mi = _make_zipinfo("mimetype", zipfile.ZIP_STORED)
        mi.flag_bits = 0
        zf.writestr(mi, "application/hwp+zip")

        # version.xml
        zf.writestr(_make_zipinfo("version.xml", zipfile.ZIP_STORED), _version_xml())

        # Contents/
        zf.writestr(_make_zipinfo("Contents/header.xml"), build_header_xml(doc))
        zf.writestr(_make_zipinfo("Contents/section0.xml"), build_section_xml(doc))

        # Chart/ — embed chart XML files
        for chart in chart_blocks:
            chart_path = f"Chart/chart{chart.chart_id}.xml"
            zf.writestr(_make_zipinfo(chart_path), build_chart_xml(chart))

        # BinData/ — embed image files
        for img in image_blocks:
            img_path = Path(img.src)
            if img_path.exists():
                zf.writestr(_make_zipinfo(img.bin_path), img_path.read_bytes())
            elif hasattr(img, '_image_data') and img._image_data:
                zf.writestr(_make_zipinfo(img.bin_path), img._image_data)

        # BinData/ — embed DocTitle decoration images
        for dt in doc_title_blocks:
            if dt.top_image_path.exists():
                zf.writestr(_make_zipinfo(dt.top_bin_path), dt.top_image_path.read_bytes())
            if dt.bottom_image_path.exists():
                zf.writestr(_make_zipinfo(dt.bottom_bin_path), dt.bottom_image_path.read_bytes())

        # Preview/
        zf.writestr(_make_zipinfo("Preview/PrvText.txt"), "")

        # settings.xml
        zf.writestr(_make_zipinfo("settings.xml"), _settings_xml())

        # META-INF/
        zf.writestr(_make_zipinfo("META-INF/container.rdf"), _container_rdf())
        zf.writestr(_make_zipinfo("Contents/content.hpf"),
                     _content_hpf(image_blocks, all_bin_items))
        zf.writestr(_make_zipinfo("META-INF/container.xml"), _container_xml())
        zf.writestr(_make_zipinfo("META-INF/manifest.xml"), _manifest_xml())


def _version_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<hv:HWPVersion major="1" minor="1" micro="0" buildNumber="0"'
        b' xmlns:hv="urn:hancom:hwpml:version"/>\n'
    )


def _container_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        b'<ocf:container'
        b' xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container"'
        b' xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">'
        b'<ocf:rootfiles>'
        b'<ocf:rootfile full-path="Contents/content.hpf"'
        b' media-type="application/hwpml-package+xml"/>'
        b'<ocf:rootfile full-path="Preview/PrvText.txt"'
        b' media-type="text/plain"/>'
        b'<ocf:rootfile full-path="META-INF/container.rdf"'
        b' media-type="application/rdf+xml"/>'
        b'</ocf:rootfiles>'
        b'</ocf:container>'
    )


def _container_rdf() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        b'<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
        b'<rdf:Description rdf:about="">'
        b'<ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#"'
        b' rdf:resource="Contents/header.xml"/>'
        b'</rdf:Description>'
        b'<rdf:Description rdf:about="Contents/header.xml">'
        b'<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/>'
        b'</rdf:Description>'
        b'<rdf:Description rdf:about="">'
        b'<ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#"'
        b' rdf:resource="Contents/section0.xml"/>'
        b'</rdf:Description>'
        b'<rdf:Description rdf:about="Contents/section0.xml">'
        b'<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/>'
        b'</rdf:Description>'
        b'<rdf:Description rdf:about="">'
        b'<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/>'
        b'</rdf:Description>'
        b'</rdf:RDF>'
    )


def _manifest_xml() -> bytes:
    # Sample uses empty self-closing manifest
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        b'<odf:manifest'
        b' xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'
    )


def _settings_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        b'<ha:HWPApplicationSetting'
        b' xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"'
        b' xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
        b'<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>'
        b'</ha:HWPApplicationSetting>'
    )


_XMLNS_ALL = (
    ' xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"'
    ' xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"'
    ' xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph"'
    ' xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"'
    ' xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"'
    ' xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"'
    ' xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history"'
    ' xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page"'
    ' xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"'
    ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
    ' xmlns:opf="http://www.idpf.org/2007/opf/"'
    ' xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart"'
    ' xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"'
    ' xmlns:epub="http://www.idpf.org/2007/ops"'
    ' xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)


def _content_hpf(image_blocks=None, all_bin_items=None) -> bytes:
    """Build Contents/content.hpf with image items registered.

    Reference (from real HWPX): each image gets an opf:item entry with
    id="imageN", href="BinData/imageN.ext", media-type, isEmbeded="1".
    Note: "isEmbeded" is Hancom's actual spelling (not "isEmbedded").
    """
    # Build image opf:item entries from all_bin_items if provided
    img_items = ''
    if all_bin_items:
        for item_id, href, mime in all_bin_items:
            img_items += (
                f'<opf:item id="{item_id}" href="{href}"'
                f' media-type="{mime}" isEmbeded="1"/>'
            )
    else:
        for img in (image_blocks or []):
            mime = 'image/jpeg' if img.format == 'jpg' else f'image/{img.format}'
            img_items += (
                f'<opf:item id="image{img._bin_id}" href="{img.bin_path}"'
                f' media-type="{mime}" isEmbeded="1"/>'
            )

    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        f'<opf:package{_XMLNS_ALL}'
        f' version="" unique-identifier="" id="">'
        f'<opf:metadata>'
        f'<opf:title/>'
        f'<opf:language>ko</opf:language>'
        f'</opf:metadata>'
        f'<opf:manifest>'
        f'<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>'
        f'{img_items}'
        f'<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>'
        f'<opf:item id="settings" href="settings.xml" media-type="application/xml"/>'
        f'</opf:manifest>'
        f'<opf:spine>'
        f'<opf:itemref idref="header" linear="yes"/>'
        f'<opf:itemref idref="section0" linear="yes"/>'
        f'</opf:spine>'
        f'</opf:package>'
    ).encode('utf-8')
