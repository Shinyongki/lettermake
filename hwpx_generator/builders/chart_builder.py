"""Build OOXML chart XML (Chart/chartN.xml) for HWPX documents.

Generates DrawingML-compatible chart XML matching real Hancom HWPX structure
(reverse-engineered from shape.hwpx Chart/chart1.xml).

Key structural rules from reference:
- c:date1904 val="0" at top
- mc:AlternateContent for chart style
- Each c:ser has: c:tx with c:strRef+c:strCache, c:invertIfNegative,
  c:cat with c:strRef+c:strCache, c:val with c:numRef+c:numCache
- catAx has: majorTickMark, minorTickMark, tickLblPos, crosses, auto,
  lblAlgn, lblOffset, tickMarkSkip, noMultiLvlLbl
- valAx has: majorGridlines, numFmt, majorTickMark, minorTickMark,
  tickLblPos, crosses, crossBetween
- c:plotArea has c:spPr
- Legend position "r" (right)
- Global c:txPr with Korean font
- c:extLst with Hancom extension
"""

from __future__ import annotations

from typing import List, TYPE_CHECKING
from xml.sax.saxutils import escape

if TYPE_CHECKING:
    from ..elements.chart import Chart, ChartDataset

# OOXML namespaces
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Default color palette for datasets without explicit colors
DEFAULT_COLORS = [
    "2E75B6", "ED7D31", "A5A5A5", "FFC000", "5B9BD5",
    "70AD47", "264478", "9B57A0", "636363", "EB7E30",
]


def build_chart_xml(chart: Chart) -> bytes:
    """Build the full Chart/chartN.xml content matching real HWPX structure.

    Args:
        chart: The Chart element with data.

    Returns:
        UTF-8 encoded XML bytes for the chart file.
    """
    lines: List[str] = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    lines.append(
        f'<c:chartSpace xmlns:r="{NS_R}" xmlns:a="{NS_A}" xmlns:c="{NS_C}">'
    )

    # date1904 — matching reference
    lines.append('<c:date1904 val="0"/>')
    lines.append('<c:roundedCorners val="0"/>')

    # mc:AlternateContent for style — matching reference
    lines.append(
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    )
    lines.append(
        '<mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"'
        ' Requires="c14"><c14:style val="102"/></mc:Choice>'
    )
    lines.append('<mc:Fallback><c:style val="2"/></mc:Fallback>')
    lines.append('</mc:AlternateContent>')

    lines.append('<c:chart>')

    # Title
    if chart.title:
        lines.append('<c:title>')
        lines.append('<c:tx>')
        lines.append('<c:rich>')
        lines.append(
            '<a:bodyPr rot="0" vert="horz" wrap="none"'
            ' lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="1"/>'
        )
        lines.append('<a:lstStyle/>')
        lines.append('<a:p>')
        lines.append('<a:pPr algn="l"><a:defRPr b="0" i="0" u="none"/></a:pPr>')
        lines.append(f'<a:r><a:t>{escape(chart.title)}</a:t></a:r>')
        lines.append('</a:p>')
        lines.append('</c:rich>')
        lines.append('</c:tx>')
        lines.append('<c:layout/>')
        lines.append('<c:overlay val="0"/>')
        lines.append('</c:title>')

    lines.append('<c:autoTitleDeleted val="0"/>')

    # Plot area
    lines.append('<c:plotArea>')
    lines.append('<c:layout/>')

    # Chart type-specific content
    ct = chart.chart_type
    if ct == "bar":
        _append_bar_chart(lines, chart, bar_dir="col", grouping="clustered")
    elif ct == "barh":
        _append_bar_chart(lines, chart, bar_dir="bar", grouping="clustered")
    elif ct == "stacked_bar":
        _append_bar_chart(lines, chart, bar_dir="col", grouping="stacked")
    elif ct == "line":
        _append_line_chart(lines, chart)
    elif ct == "pie":
        _append_pie_chart(lines, chart)
    else:
        _append_bar_chart(lines, chart, bar_dir="col", grouping="clustered")

    # Axes (except pie) — matching reference structure
    if ct != "pie":
        # Category axis
        lines.append('<c:catAx>')
        lines.append('<c:axId val="197262204"/>')
        lines.append('<c:scaling><c:orientation val="minMax"/></c:scaling>')
        lines.append('<c:axPos val="b"/>')
        lines.append('<c:crossAx val="400635792"/>')
        lines.append('<c:delete val="0"/>')
        lines.append('<c:majorTickMark val="out"/>')
        lines.append('<c:minorTickMark val="none"/>')
        lines.append('<c:tickLblPos val="nextTo"/>')
        lines.append('<c:crosses val="autoZero"/>')
        lines.append('<c:auto val="1"/>')
        lines.append('<c:lblAlgn val="ctr"/>')
        lines.append('<c:lblOffset val="100"/>')
        lines.append('<c:tickMarkSkip val="1"/>')
        lines.append('<c:noMultiLvlLbl val="0"/>')
        lines.append('</c:catAx>')

        # Value axis
        lines.append('<c:valAx>')
        lines.append('<c:axId val="400635792"/>')
        lines.append('<c:scaling><c:orientation val="minMax"/></c:scaling>')
        lines.append('<c:axPos val="l"/>')
        lines.append('<c:crossAx val="197262204"/>')
        lines.append('<c:delete val="0"/>')
        lines.append('<c:majorGridlines/>')
        lines.append('<c:numFmt formatCode="General" sourceLinked="1"/>')
        lines.append('<c:majorTickMark val="out"/>')
        lines.append('<c:minorTickMark val="none"/>')
        lines.append('<c:tickLblPos val="nextTo"/>')
        lines.append('<c:crosses val="autoZero"/>')
        lines.append('<c:crossBetween val="between"/>')
        lines.append('</c:valAx>')

    # plotArea spPr — matching reference
    lines.append('<c:spPr>')
    lines.append('<a:noFill/>')
    lines.append(
        '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
        '<a:noFill/>'
        '<a:prstDash val="solid"/>'
        '<a:round/>'
        '<a:headEnd w="med" len="med"/>'
        '<a:tailEnd w="med" len="med"/>'
        '</a:ln>'
    )
    lines.append('</c:spPr>')

    lines.append('</c:plotArea>')

    # Legend — matching reference (position right)
    lines.append('<c:legend>')
    lines.append('<c:legendPos val="r"/>')
    lines.append('<c:layout/>')
    lines.append('<c:overlay val="0"/>')
    lines.append('</c:legend>')

    lines.append('<c:plotVisOnly val="0"/>')
    lines.append('<c:dispBlanksAs val="gap"/>')
    lines.append('</c:chart>')

    # Global text properties — matching reference
    lines.append('<c:txPr>')
    lines.append(
        '<a:bodyPr rot="0" vert="horz" wrap="none"'
        ' lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="1"/>'
    )
    lines.append('<a:p>')
    lines.append('<a:pPr algn="l">')
    lines.append(
        '<a:defRPr sz="1000" b="0" i="0" u="none">'
        '<a:latin typeface="함초롬돋움"/>'
        '<a:ea typeface="함초롬돋움"/>'
        '<a:cs typeface="함초롬돋움"/>'
        '<a:sym typeface="함초롬돋움"/>'
        '</a:defRPr>'
    )
    lines.append('</a:pPr>')
    lines.append('<a:endParaRPr/>')
    lines.append('</a:p>')
    lines.append('</c:txPr>')

    # Hancom extension — matching reference
    lines.append('<c:extLst>')
    lines.append('<c:ext uri="CC8EB2C9-7E31-499d-B8F2-F6CE61031016">')
    lines.append(
        '<ho:hncChartStyle xmlns:ho="http://schemas.haansoft.com/office/8.0"'
        ' layoutIndex="-1" colorIndex="0" styleIndex="0"/>'
    )
    lines.append('</c:ext>')
    lines.append('</c:extLst>')

    lines.append('</c:chartSpace>')

    return "\n".join(lines).encode("utf-8")


# ── Chart type builders ──────────────────────────────────────────

def _append_bar_chart(
    lines: List[str], chart: Chart, bar_dir: str, grouping: str,
) -> None:
    """Append c:barChart element matching reference structure."""
    lines.append('<c:barChart>')
    lines.append(f'<c:barDir val="{bar_dir}"/>')
    lines.append(f'<c:grouping val="{grouping}"/>')
    lines.append('<c:varyColors val="0"/>')

    for idx, ds in enumerate(chart.data.datasets):
        _append_series(lines, idx, ds, chart.data.labels)

    # gapWidth and overlap for stacked
    lines.append('<c:gapWidth val="150"/>')
    if grouping == "stacked":
        lines.append('<c:overlap val="100"/>')

    lines.append('<c:axId val="197262204"/>')
    lines.append('<c:axId val="400635792"/>')
    lines.append('</c:barChart>')


def _append_line_chart(lines: List[str], chart: Chart) -> None:
    """Append c:lineChart element."""
    lines.append('<c:lineChart>')
    lines.append('<c:grouping val="standard"/>')
    lines.append('<c:varyColors val="0"/>')

    for idx, ds in enumerate(chart.data.datasets):
        _append_series(lines, idx, ds, chart.data.labels, is_line=True)

    lines.append('<c:axId val="197262204"/>')
    lines.append('<c:axId val="400635792"/>')
    lines.append('</c:lineChart>')


def _append_pie_chart(lines: List[str], chart: Chart) -> None:
    """Append c:pieChart element."""
    lines.append('<c:pieChart>')
    lines.append('<c:varyColors val="1"/>')

    for idx, ds in enumerate(chart.data.datasets):
        _append_series(lines, idx, ds, chart.data.labels)

    lines.append('</c:pieChart>')


# ── Series builder ───────────────────────────────────────────────

def _append_series(
    lines: List[str],
    idx: int,
    ds: ChartDataset,
    labels: List[str],
    is_line: bool = False,
) -> None:
    """Append a c:ser element matching reference structure."""
    color = ds.color or DEFAULT_COLORS[idx % len(DEFAULT_COLORS)]

    lines.append('<c:ser>')
    lines.append(f'<c:idx val="{idx}"/>')
    lines.append(f'<c:order val="{idx}"/>')

    # Series label with strRef+strCache — matching reference
    if ds.label:
        lines.append('<c:tx>')
        lines.append('<c:strRef>')
        # Formula reference (Sheet1! style)
        col_letter = chr(ord('B') + idx)
        lines.append(f'<c:f>Sheet1!${col_letter}$1</c:f>')
        lines.append('<c:strCache>')
        lines.append('<c:ptCount val="1"/>')
        lines.append(f'<c:pt idx="0"><c:v>{escape(ds.label)}</c:v></c:pt>')
        lines.append('</c:strCache>')
        lines.append('</c:strRef>')
        lines.append('</c:tx>')

    # invertIfNegative — matching reference
    if not is_line:
        lines.append('<c:invertIfNegative val="0"/>')

    # Series color
    lines.append('<c:spPr>')
    if is_line:
        lines.append(f'<a:ln w="28575">')
        lines.append(f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>')
        lines.append(f'</a:ln>')
    else:
        lines.append(f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>')
    lines.append('</c:spPr>')

    # Line marker
    if is_line:
        lines.append('<c:marker>')
        lines.append('<c:symbol val="circle"/>')
        lines.append('<c:size val="5"/>')
        lines.append('</c:marker>')

    # Category (labels) with strRef+strCache — matching reference
    if labels:
        lines.append('<c:cat>')
        lines.append('<c:strRef>')
        lines.append(f'<c:f>Sheet1!$A$2:$A${len(labels) + 1}</c:f>')
        lines.append('<c:strCache>')
        lines.append(f'<c:ptCount val="{len(labels)}"/>')
        for i, lbl in enumerate(labels):
            lines.append(
                f'<c:pt idx="{i}"><c:v>{escape(lbl)}</c:v></c:pt>'
            )
        lines.append('</c:strCache>')
        lines.append('</c:strRef>')
        lines.append('</c:cat>')

    # Values with numRef+numCache — matching reference
    lines.append('<c:val>')
    lines.append('<c:numRef>')
    col_letter = chr(ord('B') + idx)
    lines.append(f'<c:f>Sheet1!${col_letter}$2:${col_letter}${len(ds.values) + 1}</c:f>')
    lines.append('<c:numCache>')
    lines.append('<c:formatCode>General</c:formatCode>')
    lines.append(f'<c:ptCount val="{len(ds.values)}"/>')
    for i, val in enumerate(ds.values):
        lines.append(
            f'<c:pt idx="{i}"><c:v>{val}</c:v></c:pt>'
        )
    lines.append('</c:numCache>')
    lines.append('</c:numRef>')
    lines.append('</c:val>')

    lines.append('</c:ser>')
