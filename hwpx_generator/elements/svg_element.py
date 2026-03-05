"""SVG element for HWPX documents — converts SVG to native HWPX shapes."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class SvgElement:
    """An SVG element to be converted to native HWPX vector shapes.

    Attributes:
        src: Path to the SVG file.
        svg_string: Raw SVG XML string (alternative to src).
        width_mm: Display width in millimeters.
        height_mm: Display height in mm (auto-calculated from viewBox if omitted).
        align: Horizontal alignment -- "left", "center", or "right".
        caption: Optional caption text below the SVG.
    """

    src: str = ""
    svg_string: str = ""
    width_mm: float = 160.0
    height_mm: Optional[float] = None
    align: str = "center"
    caption: Optional[str] = None

    def load_svg(self) -> str:
        """Load SVG content from file or return svg_string.

        Returns:
            The SVG XML string.
        """
        if self.svg_string:
            return self.svg_string

        if self.src:
            from pathlib import Path
            p = Path(self.src)
            if p.exists():
                return p.read_text(encoding="utf-8")
            raise FileNotFoundError(f"SVG file not found: {self.src}")

        raise ValueError("SvgElement requires either src or svg_string")
