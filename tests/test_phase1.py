"""Phase 1 tests: basic HWPX structure generation."""

import zipfile
import os
import sys
from pathlib import Path

# Add parent to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, PageSettings


def test_basic_save():
    """Test that a basic document produces a valid ZIP with expected entries."""
    doc = HwpxDocument()
    out = Path(__file__).parent / "output_phase1.hwpx"
    doc.save(str(out))

    assert out.exists(), "Output file was not created"
    assert out.stat().st_size > 0, "Output file is empty"

    with zipfile.ZipFile(out, "r") as zf:
        names = zf.namelist()
        expected = [
            "mimetype",
            "version.xml",
            "META-INF/container.xml",
            "META-INF/manifest.xml",
            "settings.xml",
            "Contents/content.hpf",
            "Contents/header.xml",
            "Contents/section0.xml",
            "Preview/PrvText.txt",
        ]
        for entry in expected:
            assert entry in names, f"Missing entry: {entry}"

        # mimetype must be first and uncompressed
        info = zf.getinfo("mimetype")
        assert info.compress_type == zipfile.ZIP_STORED, "mimetype must be uncompressed"
        assert zf.read("mimetype") == b"application/hwp+zip"

        # Check XML is valid UTF-8
        for entry in expected:
            if entry.endswith(".xml") or entry.endswith(".hpf"):
                content = zf.read(entry)
                content.decode("utf-8")  # should not raise

    print("[PASS] test_basic_save")
    out.unlink()


def test_custom_page_settings():
    """Test custom page settings are reflected in section0.xml."""
    doc = HwpxDocument()
    doc.page_settings.margin_top = 10
    doc.page_settings.margin_bottom = 10
    doc.page_settings.margin_left = 20
    doc.page_settings.margin_right = 20

    out = Path(__file__).parent / "output_phase1_custom.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        # margin left=20mm -> 20*7200/25.4 = 5669
        assert 'left="5669"' in section, f"Expected left margin 5669 in section XML"
        # margin top=10mm -> 10*7200/25.4 = 2835
        assert 'top="2835"' in section, f"Expected top margin 2835 in section XML"

    print("[PASS] test_custom_page_settings")
    out.unlink()


def test_landscape_orientation():
    """Test landscape orientation."""
    doc = HwpxDocument()
    doc.page_settings.orientation = "landscape"

    out = Path(__file__).parent / "output_phase1_landscape.hwpx"
    doc.save(str(out))

    with zipfile.ZipFile(out, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        assert 'landscape="NARROWLY"' in section

    print("[PASS] test_landscape_orientation")
    out.unlink()


def test_page_settings_properties():
    """Test PageSettings computed properties."""
    ps = PageSettings()
    ps.margin_left = 20
    ps.margin_right = 20
    assert ps.content_width_mm == 170.0  # 210 - 20 - 20

    ps.orientation = "landscape"
    assert ps.content_width_mm == 257.0  # 297 - 20 - 20

    print("[PASS] test_page_settings_properties")


def test_unit_conversion():
    """Test mm/pt to HWP unit conversion."""
    from hwpx_generator.utils import mm_to_hwp, pt_to_hwp

    # 1mm = 7200/25.4 ≈ 283.46 -> round to 283
    assert mm_to_hwp(1) == 283
    # 10mm = 2835
    assert mm_to_hwp(10) == 2835
    # 20mm = 5669
    assert mm_to_hwp(20) == 5669
    # 210mm (A4 width) = 59528
    assert mm_to_hwp(210) == 59528

    # pt conversion: 10pt = 1000
    assert pt_to_hwp(10) == 1000
    assert pt_to_hwp(13) == 1300

    print("[PASS] test_unit_conversion")


if __name__ == "__main__":
    test_unit_conversion()
    test_page_settings_properties()
    test_basic_save()
    test_custom_page_settings()
    test_landscape_orientation()
    print("\n=== All Phase 1 tests passed! ===")
