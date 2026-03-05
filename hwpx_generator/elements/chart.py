"""Chart element — OOXML native chart for HWPX documents."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any


@dataclass
class ChartDataset:
    """A single data series for a chart."""

    label: str = ""
    values: List[float] = field(default_factory=list)
    color: Optional[str] = None  # hex, e.g. "2E75B6"


@dataclass
class ChartData:
    """Chart data containing labels and datasets."""

    labels: List[str] = field(default_factory=list)
    datasets: List[ChartDataset] = field(default_factory=list)


@dataclass
class Chart:
    """A chart block.

    Supported chart_type values:
        - bar: Vertical bar chart (c:barChart barDir=col, grouping=clustered)
        - barh: Horizontal bar chart (c:barChart barDir=bar, grouping=clustered)
        - line: Line chart (c:lineChart)
        - pie: Pie chart (c:pieChart)
        - stacked_bar: Stacked bar (c:barChart barDir=col, grouping=stacked)
    """

    chart_type: str = "bar"        # bar, barh, line, pie, stacked_bar
    width_mm: float = 120.0
    height_mm: float = 70.0
    title: str = ""
    data: ChartData = field(default_factory=ChartData)

    # Internal: assigned by the document during save
    chart_id: int = 0  # Index for Chart/chart{N}.xml

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> Chart:
        """Create a Chart from a JSON-like dictionary."""
        data_dict = d.get("data", {})
        datasets = []
        for ds in data_dict.get("datasets", []):
            datasets.append(ChartDataset(
                label=ds.get("label", ""),
                values=ds.get("values", []),
                color=ds.get("color"),
            ))
        chart_data = ChartData(
            labels=data_dict.get("labels", []),
            datasets=datasets,
        )
        return cls(
            chart_type=d.get("chart_type", "bar"),
            width_mm=d.get("width_mm", 120.0),
            height_mm=d.get("height_mm", 70.0),
            title=d.get("title", ""),
            data=chart_data,
        )
