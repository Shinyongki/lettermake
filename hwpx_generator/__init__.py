"""hwpx_generator — Python library for generating HWPX documents."""

from .document import HwpxDocument, PageSettings
from .elements.paragraph import Paragraph, TextRun
from .elements.table import Table, TableRow, TableCell
from .elements.page_break import PageBreak
from .elements.image import Image
from .elements.diagram import Diagram, Node, Edge
from .elements.chart import Chart, ChartData, ChartDataset
from .elements.svg_element import SvgElement
from .elements.text_box import TextBox
from .elements.doc_title import DocTitle
from .elements.doc_purpose import DocPurpose
from .elements.section_block import SectionBlock
from .styles import StyleManager, Style, CharProps, ParaProps, BorderFillProps
from .presets.gov_document import GovDocumentPreset, BulletStyle

__all__ = [
    "HwpxDocument", "PageSettings",
    "Paragraph", "TextRun",
    "Table", "TableRow", "TableCell",
    "PageBreak",
    "Image",
    "Diagram", "Node", "Edge",
    "Chart", "ChartData", "ChartDataset",
    "SvgElement",
    "TextBox",
    "DocTitle", "DocPurpose", "SectionBlock",
    "StyleManager", "Style", "CharProps", "ParaProps", "BorderFillProps",
    "GovDocumentPreset", "BulletStyle",
]
__version__ = "0.1.0"
