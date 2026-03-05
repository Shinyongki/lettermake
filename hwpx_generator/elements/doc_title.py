"""DocTitle element — document title with top/bottom decoration lines."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional


@dataclass
class DocTitle:
    """Document title block: decoration image (top) + title text + decoration image (bottom).

    The decoration images are fixed assets embedded in the package.
    """

    text: str = ""

    # Internal: assigned during finalize
    _top_bin_id: int = 0
    _bottom_bin_id: int = 0
    _top_pixel_w: int = 786
    _top_pixel_h: int = 10
    _bottom_pixel_w: int = 787
    _bottom_pixel_h: int = 10

    @staticmethod
    def asset_dir() -> Path:
        """Return the path to the assets directory."""
        return Path(__file__).resolve().parent.parent / "assets"

    @property
    def top_image_path(self) -> Path:
        return self.asset_dir() / "doc_title_top.png"

    @property
    def bottom_image_path(self) -> Path:
        return self.asset_dir() / "doc_title_bottom.png"

    @property
    def top_bin_path(self) -> str:
        return f"BinData/image{self._top_bin_id}.png"

    @property
    def bottom_bin_path(self) -> str:
        return f"BinData/image{self._bottom_bin_id}.png"
