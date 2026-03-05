"""Diagram element — native HWPX vector shape diagrams."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class Node:
    """A single node in a diagram."""

    id: str
    label: str = ""
    shape: str = "rect"  # rect, rounded_rect, ellipse, diamond
    fill_color: Optional[str] = None  # hex, None = use theme
    line_color: Optional[str] = None  # hex, None = use theme
    font_color: Optional[str] = None  # hex, None = "000000"
    font_size_pt: Optional[float] = None


@dataclass
class Edge:
    """An edge (connector) between two nodes."""

    from_id: str
    to_id: str
    head_style: str = "NORMAL"   # NORMAL (no arrow)
    tail_style: str = "ARROW"    # ARROW
    tail_sz: str = "MEDIUM_MEDIUM"
    line_color: Optional[str] = None  # hex, None = use theme


@dataclass
class Diagram:
    """A diagram block containing nodes and edges.

    Supported layouts:
        - step_flow: Horizontal or vertical step flow
        - free: Direct coordinate specification (nodes must have x, y, w, h)
    """

    layout: str = "step_flow"     # step_flow, free
    direction: str = "horizontal"  # horizontal, vertical
    width_mm: float = 160.0
    height_mm: float = 30.0
    theme_color: str = "1F4E79"   # hex
    nodes: List[Node] = field(default_factory=list)
    edges: List[Edge] = field(default_factory=list)

    def add_node(
        self,
        id: str,
        label: str = "",
        shape: str = "rect",
        **kwargs,
    ) -> Node:
        """Add a node to the diagram."""
        node = Node(id=id, label=label, shape=shape, **kwargs)
        self.nodes.append(node)
        return node

    def add_edge(
        self,
        from_id: str,
        to_id: str,
        **kwargs,
    ) -> Edge:
        """Add an edge between two nodes."""
        edge = Edge(from_id=from_id, to_id=to_id, **kwargs)
        self.edges.append(edge)
        return edge
