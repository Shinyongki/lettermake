"""Phase 6 tests: image insertion with BinData packaging.

Tests are based on real HWPX file structure (reverse-engineered from
image1.hwpx and image2.hwpx reference files created in Hancom Hangul).

Key findings from reference files:
- Image tag is hp:pic (NOT hp:picture)
- Image reference uses <hc:img binaryItemIDRef="imageN"/> (NOT binDataIDRef)
- Image manifest is in content.hpf via <opf:item id="imageN" ... isEmbeded="1"/>
- header.xml has NO binData section
- manifest.xml stays empty
- Caption is inside hp:pic > hp:caption > hp:subList > hp:p
"""

import struct
import zipfile
import sys
import os
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, GovDocumentPreset, Image


# ─── Helper: create minimal valid image files in memory ─────────────


def _make_minimal_png(width=100, height=80):
    """Create a minimal valid PNG with custom width/height in IHDR."""
    import zlib

    sig = b'\x89PNG\r\n\x1a\n'

    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    ihdr = _png_chunk(b"IHDR", ihdr_data)

    raw_row = b'\x00' + b'\xff\x00\x00' * width
    raw_data = b''
    for _ in range(height):
        raw_data += raw_row
    compressed = zlib.compress(raw_data)
    idat = _png_chunk(b"IDAT", compressed)

    iend = _png_chunk(b"IEND", b"")

    return sig + ihdr + idat + iend


def _png_chunk(chunk_type, data):
    import zlib
    length = struct.pack(">I", len(data))
    crc = struct.pack(">I", zlib.crc32(chunk_type + data) & 0xFFFFFFFF)
    return length + chunk_type + data + crc


def _make_minimal_jpeg(width=120, height=90):
    """Create a minimal valid JPEG with SOF0 marker containing width/height."""
    buf = bytearray()

    buf += b'\xff\xd8'

    app0_data = b'JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00'
    buf += b'\xff\xe0'
    buf += struct.pack(">H", len(app0_data) + 2)
    buf += app0_data

    dqt_data = b'\x00' + bytes(range(1, 65))
    buf += b'\xff\xdb'
    buf += struct.pack(">H", len(dqt_data) + 2)
    buf += dqt_data

    sof_data = struct.pack(">BHH", 8, height, width)
    sof_data += b'\x01'
    sof_data += b'\x01\x11\x00'
    buf += b'\xff\xc0'
    buf += struct.pack(">H", len(sof_data) + 2)
    buf += sof_data

    dht_data = b'\x00'
    dht_data += b'\x00' * 16
    buf += b'\xff\xc4'
    buf += struct.pack(">H", len(dht_data) + 2)
    buf += dht_data

    sos_data = b'\x01\x01\x00\x00\x3f\x00'
    buf += b'\xff\xda'
    buf += struct.pack(">H", len(sos_data) + 2)
    buf += sos_data

    buf += b'\x00\xff\xd9'

    return bytes(buf)


# ─── Helper: read HWPX contents ────────────────────────────────────


def _read_hwpx(path):
    with zipfile.ZipFile(path, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        header = zf.read("Contents/header.xml").decode("utf-8")
        content_hpf = zf.read("Contents/content.hpf").decode("utf-8")
        names = zf.namelist()
    return section, header, content_hpf, names


# ─── Tests ──────────────────────────────────────────────────────────


def test_png_dimensions_reader():
    """Test that PNG dimensions are correctly read from IHDR."""
    data = _make_minimal_png(200, 150)
    img = Image(src="test.png", width_mm=80)
    img.resolve_dimensions_from_bytes(data, "png")

    assert img._pixel_width == 200, f"Expected pixel width 200, got {img._pixel_width}"
    assert img._pixel_height == 150, f"Expected pixel height 150, got {img._pixel_height}"
    assert img.height_mm is not None
    expected_h = 80.0 * 150 / 200  # = 60.0
    assert abs(img.height_mm - expected_h) < 0.01, \
        f"Expected height_mm ~{expected_h}, got {img.height_mm}"

    print("[PASS] test_png_dimensions_reader")


def test_jpeg_dimensions_reader():
    """Test that JPEG dimensions are correctly read from SOF0 marker."""
    data = _make_minimal_jpeg(320, 240)
    img = Image(src="test.jpg", width_mm=100)
    img.resolve_dimensions_from_bytes(data, "jpg")

    assert img._pixel_width == 320, f"Expected pixel width 320, got {img._pixel_width}"
    assert img._pixel_height == 240, f"Expected pixel height 240, got {img._pixel_height}"
    assert img.height_mm is not None
    expected_h = 100.0 * 240 / 320  # = 75.0
    assert abs(img.height_mm - expected_h) < 0.01, \
        f"Expected height_mm ~{expected_h}, got {img.height_mm}"

    print("[PASS] test_jpeg_dimensions_reader")


def test_image_format_detection():
    """Test format detection from file extension."""
    img_png = Image(src="photos/test.png")
    assert img_png.format == "png"

    img_jpg = Image(src="photos/test.jpg")
    assert img_jpg.format == "jpg"

    img_jpeg = Image(src="photos/test.JPEG")
    assert img_jpeg.format == "jpg"

    print("[PASS] test_image_format_detection")


def test_image_bin_path():
    """Test internal bin path generation."""
    img = Image(src="photos/test.png")
    img._bin_id = 3
    assert img.bin_filename == "image3.png"
    assert img.bin_path == "BinData/image3.png"

    img2 = Image(src="photos/test.jpg")
    img2._bin_id = 1
    assert img2.bin_filename == "image1.jpg"
    assert img2.bin_path == "BinData/image1.jpg"

    print("[PASS] test_image_bin_path")


def test_add_image_from_bytes_png():
    """Test adding a PNG image from bytes to a document and saving."""
    doc = HwpxDocument()
    png_data = _make_minimal_png(200, 150)

    doc.add_paragraph("Image test document", bold=True, align="CENTER")
    doc.add_image_from_bytes(png_data, "png", width_mm=80, align="center",
                             caption="Figure 1. Test PNG image")

    out = Path(__file__).parent / "output_p6_png.hwpx"
    doc.save(str(out))

    section, header, content_hpf, names = _read_hwpx(out)

    # Check hp:pic element in section XML (NOT hp:picture)
    assert "hp:pic " in section, "hp:pic element missing from section XML"
    assert 'binaryItemIDRef="image1"' in section, "binaryItemIDRef missing"
    assert "Figure 1. Test PNG image" in section, "Caption missing from section XML"
    # Caption should be inside hp:caption, not a separate paragraph
    assert "hp:caption" in section, "hp:caption element missing"

    # Check image item in content.hpf (NOT in header.xml)
    assert 'id="image1"' in content_hpf, "image1 item missing from content.hpf"
    assert 'href="BinData/image1.png"' in content_hpf, "BinData path missing from content.hpf"
    assert 'media-type="image/png"' in content_hpf, "PNG media type missing"
    assert 'isEmbeded="1"' in content_hpf, "isEmbeded flag missing"

    # Check BinData file in ZIP
    assert "BinData/image1.png" in names, "BinData/image1.png not in ZIP"

    # Verify image data in ZIP is correct
    with zipfile.ZipFile(out, "r") as zf:
        stored = zf.read("BinData/image1.png")
    assert stored == png_data, "Stored image data does not match original"

    print("[PASS] test_add_image_from_bytes_png")
    out.unlink()


def test_add_image_from_bytes_jpg():
    """Test adding a JPEG image from bytes."""
    doc = HwpxDocument()
    jpg_data = _make_minimal_jpeg(320, 240)

    doc.add_image_from_bytes(jpg_data, "jpg", width_mm=100, align="left")

    out = Path(__file__).parent / "output_p6_jpg.hwpx"
    doc.save(str(out))

    section, header, content_hpf, names = _read_hwpx(out)

    assert "hp:pic " in section
    assert 'media-type="image/jpeg"' in content_hpf
    assert "BinData/image1.jpg" in names
    assert 'horzAlign="LEFT"' in section

    print("[PASS] test_add_image_from_bytes_jpg")
    out.unlink()


def test_image_alignment():
    """Test that image alignment maps correctly to HWPX horzAlign."""
    doc = HwpxDocument()
    png_data = _make_minimal_png(100, 100)

    doc.add_image_from_bytes(png_data, "png", width_mm=50, align="left")
    doc.add_image_from_bytes(png_data, "png", width_mm=50, align="center")
    doc.add_image_from_bytes(png_data, "png", width_mm=50, align="right")

    out = Path(__file__).parent / "output_p6_align.hwpx"
    doc.save(str(out))

    section, _, _, _ = _read_hwpx(out)

    assert 'horzAlign="LEFT"' in section, "LEFT alignment missing"
    assert 'horzAlign="CENTER"' in section, "CENTER alignment missing"
    assert 'horzAlign="RIGHT"' in section, "RIGHT alignment missing"

    print("[PASS] test_image_alignment")
    out.unlink()


def test_image_auto_height():
    """Test auto-calculation of height from aspect ratio."""
    png_data = _make_minimal_png(400, 300)
    img = Image(src="test.png", width_mm=80)
    img.resolve_dimensions_from_bytes(png_data, "png")

    # 400:300 = 4:3, so height should be 80 * 300/400 = 60
    assert img.height_mm is not None
    assert abs(img.height_mm - 60.0) < 0.01

    print("[PASS] test_image_auto_height")


def test_image_explicit_height():
    """Test that explicit height_mm is preserved (not overwritten)."""
    png_data = _make_minimal_png(400, 300)
    img = Image(src="test.png", width_mm=80, height_mm=40)
    img.resolve_dimensions_from_bytes(png_data, "png")

    # Explicit height should be preserved
    assert img.height_mm == 40.0, f"Expected 40.0, got {img.height_mm}"

    print("[PASS] test_image_explicit_height")


def test_multiple_images():
    """Test multiple images get sequential bin IDs."""
    doc = HwpxDocument()
    png_data = _make_minimal_png(100, 100)
    jpg_data = _make_minimal_jpeg(200, 150)

    doc.add_image_from_bytes(png_data, "png", width_mm=50, caption="Image A")
    doc.add_image_from_bytes(jpg_data, "jpg", width_mm=60, caption="Image B")
    doc.add_image_from_bytes(png_data, "png", width_mm=70)

    out = Path(__file__).parent / "output_p6_multi.hwpx"
    doc.save(str(out))

    section, header, content_hpf, names = _read_hwpx(out)

    # Check all three items in content.hpf
    assert 'id="image1"' in content_hpf
    assert 'id="image2"' in content_hpf
    assert 'id="image3"' in content_hpf
    assert 'href="BinData/image1.png"' in content_hpf
    assert 'href="BinData/image2.jpg"' in content_hpf
    assert 'href="BinData/image3.png"' in content_hpf

    # Check all three files in ZIP
    assert "BinData/image1.png" in names
    assert "BinData/image2.jpg" in names
    assert "BinData/image3.png" in names

    # Check binaryItemIDRefs in section
    assert 'binaryItemIDRef="image1"' in section
    assert 'binaryItemIDRef="image2"' in section
    assert 'binaryItemIDRef="image3"' in section

    # Check captions
    assert "Image A" in section
    assert "Image B" in section

    print("[PASS] test_multiple_images")
    out.unlink()


def test_image_with_file_path():
    """Test add_image() with an actual file on disk."""
    png_data = _make_minimal_png(250, 200)
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
        f.write(png_data)
        tmp_path = f.name

    try:
        doc = HwpxDocument()
        doc.add_image(tmp_path, width_mm=60, align="right",
                      caption="Photo from file")

        out = Path(__file__).parent / "output_p6_file.hwpx"
        doc.save(str(out))

        section, header, content_hpf, names = _read_hwpx(out)

        assert "hp:pic " in section
        assert "Photo from file" in section
        assert 'horzAlign="RIGHT"' in section
        assert "BinData/image1.png" in names
        assert 'isEmbeded="1"' in content_hpf

        # Verify stored data matches
        with zipfile.ZipFile(out, "r") as zf:
            stored = zf.read("BinData/image1.png")
        assert stored == png_data

        print("[PASS] test_image_with_file_path")
        out.unlink()
    finally:
        os.unlink(tmp_path)


def test_image_imgdim():
    """Test that imgDim values = pixel dimensions × 75 (reference conversion)."""
    doc = HwpxDocument()
    png_data = _make_minimal_png(200, 150)
    doc.add_image_from_bytes(png_data, "png", width_mm=80)

    out = Path(__file__).parent / "output_p6_dim.hwpx"
    doc.save(str(out))

    section, _, _, _ = _read_hwpx(out)

    # imgDim should be pixels × 75 (7200/96 dpi)
    assert 'dimwidth="15000"' in section, \
        f"Expected dimwidth=15000 (200*75), got something else"
    assert 'dimheight="11250"' in section, \
        f"Expected dimheight=11250 (150*75), got something else"

    print("[PASS] test_image_imgdim")
    out.unlink()


def test_image_reference_structure():
    """Test that the generated XML matches the real HWPX reference structure."""
    doc = HwpxDocument()
    png_data = _make_minimal_png(100, 100)
    doc.add_image_from_bytes(png_data, "png", width_mm=50)

    out = Path(__file__).parent / "output_p6_struct.hwpx"
    doc.save(str(out))

    section, _, _, _ = _read_hwpx(out)

    # Check all required child elements of hp:pic (from reference)
    assert "hp:orgSz" in section, "hp:orgSz missing"
    assert "hp:curSz" in section, "hp:curSz missing"
    assert "hp:flip" in section, "hp:flip missing"
    assert "hp:rotationInfo" in section, "hp:rotationInfo missing"
    assert "hp:renderingInfo" in section, "hp:renderingInfo missing"
    assert "hc:img" in section, "hc:img missing"
    assert "hp:imgRect" in section, "hp:imgRect missing"
    assert "hp:imgClip" in section, "hp:imgClip missing"
    assert "hp:imgDim" in section, "hp:imgDim missing"
    assert "hp:effects" in section, "hp:effects missing"
    assert "hp:sz" in section, "hp:sz missing"
    assert "hp:pos" in section, "hp:pos missing"
    assert "hp:outMargin" in section, "hp:outMargin missing"
    # textWrap should be TOP_AND_BOTTOM (reference: image1.hwpx)
    assert 'textWrap="TOP_AND_BOTTOM"' in section, "textWrap should be TOP_AND_BOTTOM"
    assert 'numberingType="PICTURE"' in section, "numberingType should be PICTURE"

    print("[PASS] test_image_reference_structure")
    out.unlink()


def test_image_no_caption():
    """Test image without caption has no hp:caption element."""
    doc = HwpxDocument()
    png_data = _make_minimal_png(100, 100)

    doc.add_image_from_bytes(png_data, "png", width_mm=50)

    out = Path(__file__).parent / "output_p6_nocap.hwpx"
    doc.save(str(out))

    section, _, _, _ = _read_hwpx(out)

    assert "hp:pic " in section
    assert "hp:caption" not in section, "Caption should not be present"

    print("[PASS] test_image_no_caption")
    out.unlink()


def test_full_document_with_images():
    """Integration test: document with text, tables, and images."""
    doc = HwpxDocument()
    doc.apply_preset(GovDocumentPreset(
        font_body="맑은 고딕",
        size_body=12,
    ))

    doc.add_title("2026 Monitoring Report")
    doc.add_section_heading(1, "Overview")
    doc.add_paragraph("Below is a monitoring photo:")

    png_data = _make_minimal_png(640, 480)
    doc.add_image_from_bytes(png_data, "png", width_mm=120, align="center",
                             caption="Photo 1. Site monitoring")

    doc.add_paragraph("And another image:")
    jpg_data = _make_minimal_jpeg(800, 600)
    doc.add_image_from_bytes(jpg_data, "jpg", width_mm=100, align="left",
                             caption="Photo 2. Equipment check")

    table = doc.add_table(col_widths=[40, 130])
    table.add_header_row(["Item", "Details"])
    table.add_row(["Location", "Seoul"])

    out = Path(__file__).parent / "output_p6_full.hwpx"
    doc.save(str(out))

    section, header, content_hpf, names = _read_hwpx(out)

    # Text content
    assert "2026 Monitoring Report" in section
    assert "Overview" in section

    # Images
    assert "BinData/image1.png" in names
    assert "BinData/image2.jpg" in names
    assert "Photo 1. Site monitoring" in section
    assert "Photo 2. Equipment check" in section

    # content.hpf has image items
    assert 'id="image1"' in content_hpf
    assert 'id="image2"' in content_hpf

    # Table still works
    assert "hp:tbl" in section
    assert "Seoul" in section

    print("[PASS] test_full_document_with_images")
    out.unlink()


if __name__ == "__main__":
    test_png_dimensions_reader()
    test_jpeg_dimensions_reader()
    test_image_format_detection()
    test_image_bin_path()
    test_add_image_from_bytes_png()
    test_add_image_from_bytes_jpg()
    test_image_alignment()
    test_image_auto_height()
    test_image_explicit_height()
    test_multiple_images()
    test_image_with_file_path()
    test_image_imgdim()
    test_image_reference_structure()
    test_image_no_caption()
    test_full_document_with_images()
    print("\n=== All Phase 6 tests passed! ===")
