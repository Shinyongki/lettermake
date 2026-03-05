"""Phase 8 tests: OOXML native chart generation."""

import zipfile
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from hwpx_generator import HwpxDocument, GovDocumentPreset, Chart, ChartData, ChartDataset


def _read_hwpx(path):
    with zipfile.ZipFile(path, "r") as zf:
        section = zf.read("Contents/section0.xml").decode("utf-8")
        names = zf.namelist()
        charts = {}
        for name in names:
            if name.startswith("Chart/"):
                charts[name] = zf.read(name).decode("utf-8")
        manifest = zf.read("META-INF/manifest.xml").decode("utf-8")
    return section, charts, manifest, names


def test_basic_bar_chart():
    """Test basic vertical bar chart."""
    doc = HwpxDocument()
    chart = doc.add_chart(
        chart_type="bar",
        width_mm=120,
        height_mm=70,
        title="기관별 현황",
        labels=["기관A", "기관B", "기관C"],
        datasets=[{"label": "완료", "values": [1, 1, 0], "color": "2E75B6"}],
    )

    out = Path(__file__).parent / "output_p8_bar.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section, charts, manifest, names = _read_hwpx(out)

    # Check section reference
    assert "hp:chart" in section, "Expected hp:chart reference in section"
    assert "Chart/chart1.xml" in section, "Expected chart file reference"

    # Check chart file exists
    assert "Chart/chart1.xml" in names, "Chart file not in ZIP"
    chart_xml = charts["Chart/chart1.xml"]

    # Check OOXML structure
    assert "c:chartSpace" in chart_xml
    assert "c:barChart" in chart_xml
    assert 'barDir val="col"' in chart_xml
    assert 'grouping val="clustered"' in chart_xml
    assert "기관별 현황" in chart_xml
    assert "기관A" in chart_xml
    assert "기관B" in chart_xml
    assert "완료" in chart_xml
    assert "2E75B6" in chart_xml

    # Note: manifest.xml is intentionally empty (matching real HWPX structure)
    # Chart references live in section XML's chartIDRef attribute

    print("[PASS] test_basic_bar_chart")
    out.unlink()


def test_horizontal_bar_chart():
    """Test horizontal bar chart."""
    doc = HwpxDocument()
    doc.add_chart(
        chart_type="barh",
        title="수평 바",
        labels=["X", "Y"],
        datasets=[{"label": "Series1", "values": [10, 20]}],
    )

    out = Path(__file__).parent / "output_p8_barh.hwpx"
    doc.save(str(out), auto_page_flow=False)

    _, charts, _, _ = _read_hwpx(out)
    chart_xml = charts["Chart/chart1.xml"]
    assert 'barDir val="bar"' in chart_xml, "Expected barDir=bar for horizontal"
    assert "수평 바" in chart_xml

    print("[PASS] test_horizontal_bar_chart")
    out.unlink()


def test_stacked_bar_chart():
    """Test stacked bar chart."""
    doc = HwpxDocument()
    doc.add_chart(
        chart_type="stacked_bar",
        labels=["A", "B"],
        datasets=[
            {"label": "S1", "values": [3, 5]},
            {"label": "S2", "values": [2, 4]},
        ],
    )

    out = Path(__file__).parent / "output_p8_stacked.hwpx"
    doc.save(str(out), auto_page_flow=False)

    _, charts, _, _ = _read_hwpx(out)
    chart_xml = charts["Chart/chart1.xml"]
    assert 'grouping val="stacked"' in chart_xml
    assert "S1" in chart_xml
    assert "S2" in chart_xml

    print("[PASS] test_stacked_bar_chart")
    out.unlink()


def test_line_chart():
    """Test line chart."""
    doc = HwpxDocument()
    doc.add_chart(
        chart_type="line",
        title="추이",
        labels=["1월", "2월", "3월"],
        datasets=[
            {"label": "A기관", "values": [10, 15, 12], "color": "FF0000"},
        ],
    )

    out = Path(__file__).parent / "output_p8_line.hwpx"
    doc.save(str(out), auto_page_flow=False)

    _, charts, _, _ = _read_hwpx(out)
    chart_xml = charts["Chart/chart1.xml"]
    assert "c:lineChart" in chart_xml
    assert 'grouping val="standard"' in chart_xml
    assert "추이" in chart_xml
    assert "1월" in chart_xml
    assert "FF0000" in chart_xml
    # Line chart has markers
    assert "c:marker" in chart_xml

    print("[PASS] test_line_chart")
    out.unlink()


def test_pie_chart():
    """Test pie chart."""
    doc = HwpxDocument()
    doc.add_chart(
        chart_type="pie",
        title="비율",
        labels=["완료", "진행중", "미착수"],
        datasets=[{"label": "현황", "values": [5, 3, 2]}],
    )

    out = Path(__file__).parent / "output_p8_pie.hwpx"
    doc.save(str(out), auto_page_flow=False)

    _, charts, _, _ = _read_hwpx(out)
    chart_xml = charts["Chart/chart1.xml"]
    assert "c:pieChart" in chart_xml
    assert "비율" in chart_xml
    assert "완료" in chart_xml
    # Pie charts should not have axes
    assert "c:catAx" not in chart_xml
    assert "c:valAx" not in chart_xml

    print("[PASS] test_pie_chart")
    out.unlink()


def test_multiple_datasets():
    """Test chart with multiple datasets."""
    doc = HwpxDocument()
    doc.add_chart(
        chart_type="bar",
        title="비교",
        labels=["A", "B", "C"],
        datasets=[
            {"label": "2025", "values": [1, 2, 3], "color": "2E75B6"},
            {"label": "2026", "values": [4, 5, 6], "color": "ED7D31"},
        ],
    )

    out = Path(__file__).parent / "output_p8_multi_ds.hwpx"
    doc.save(str(out), auto_page_flow=False)

    _, charts, _, _ = _read_hwpx(out)
    chart_xml = charts["Chart/chart1.xml"]
    assert chart_xml.count("<c:ser>") == 2, "Expected 2 series"
    assert "2025" in chart_xml
    assert "2026" in chart_xml
    assert "2E75B6" in chart_xml
    assert "ED7D31" in chart_xml

    print("[PASS] test_multiple_datasets")
    out.unlink()


def test_multiple_charts():
    """Test document with multiple charts."""
    doc = HwpxDocument()
    doc.add_chart(
        chart_type="bar",
        title="Chart 1",
        labels=["X"],
        datasets=[{"label": "S", "values": [1]}],
    )
    doc.add_chart(
        chart_type="pie",
        title="Chart 2",
        labels=["Y"],
        datasets=[{"label": "T", "values": [2]}],
    )

    out = Path(__file__).parent / "output_p8_multi_chart.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section, charts, manifest, names = _read_hwpx(out)

    assert "Chart/chart1.xml" in names
    assert "Chart/chart2.xml" in names
    assert "Chart/chart1.xml" in section
    assert "Chart/chart2.xml" in section
    assert "Chart 1" in charts["Chart/chart1.xml"]
    assert "Chart 2" in charts["Chart/chart2.xml"]
    assert "c:barChart" in charts["Chart/chart1.xml"]
    assert "c:pieChart" in charts["Chart/chart2.xml"]

    # Note: manifest.xml is empty; chart references are in section XML chartIDRef

    print("[PASS] test_multiple_charts")
    out.unlink()


def test_chart_no_title():
    """Test chart without a title."""
    doc = HwpxDocument()
    doc.add_chart(
        chart_type="bar",
        labels=["A", "B"],
        datasets=[{"label": "S", "values": [1, 2]}],
    )

    out = Path(__file__).parent / "output_p8_notitle.hwpx"
    doc.save(str(out), auto_page_flow=False)

    _, charts, _, _ = _read_hwpx(out)
    chart_xml = charts["Chart/chart1.xml"]
    assert "c:chartSpace" in chart_xml
    # No title element
    assert "<c:title>" not in chart_xml

    print("[PASS] test_chart_no_title")
    out.unlink()


def test_chart_with_text():
    """Integration test: chart mixed with text content."""
    doc = HwpxDocument()
    doc.add_paragraph("Before chart", bold=True)
    doc.add_chart(
        chart_type="bar",
        title="Test",
        labels=["A"],
        datasets=[{"label": "V", "values": [42]}],
    )
    doc.add_paragraph("After chart")

    out = Path(__file__).parent / "output_p8_integration.hwpx"
    doc.save(str(out), auto_page_flow=False)

    section, charts, _, _ = _read_hwpx(out)
    assert "Before chart" in section
    assert "After chart" in section
    assert "hp:chart" in section
    assert "42" in charts["Chart/chart1.xml"]

    print("[PASS] test_chart_with_text")
    out.unlink()


def test_chart_from_dict():
    """Test Chart.from_dict() class method."""
    chart = Chart.from_dict({
        "chart_type": "line",
        "width_mm": 100,
        "height_mm": 60,
        "title": "테스트",
        "data": {
            "labels": ["A", "B"],
            "datasets": [
                {"label": "S1", "values": [1, 2], "color": "FF0000"},
            ],
        },
    })

    assert chart.chart_type == "line"
    assert chart.width_mm == 100
    assert chart.title == "테스트"
    assert len(chart.data.labels) == 2
    assert len(chart.data.datasets) == 1
    assert chart.data.datasets[0].color == "FF0000"

    print("[PASS] test_chart_from_dict")


def test_default_colors():
    """Test that datasets without explicit colors get default colors."""
    doc = HwpxDocument()
    doc.add_chart(
        chart_type="bar",
        labels=["X"],
        datasets=[
            {"label": "No color", "values": [1]},
        ],
    )

    out = Path(__file__).parent / "output_p8_default_color.hwpx"
    doc.save(str(out), auto_page_flow=False)

    _, charts, _, _ = _read_hwpx(out)
    chart_xml = charts["Chart/chart1.xml"]
    # First default color is 2E75B6
    assert "2E75B6" in chart_xml

    print("[PASS] test_default_colors")
    out.unlink()


if __name__ == "__main__":
    test_basic_bar_chart()
    test_horizontal_bar_chart()
    test_stacked_bar_chart()
    test_line_chart()
    test_pie_chart()
    test_multiple_datasets()
    test_multiple_charts()
    test_chart_no_title()
    test_chart_with_text()
    test_chart_from_dict()
    test_default_colors()
    print("\n=== All Phase 8 tests passed! ===")
