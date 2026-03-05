"""Image element for HWPX documents."""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class Image:
    """An image element (PNG or JPG) to be embedded in the document.

    Attributes:
        src: Path to the image file (relative or absolute).
        width_mm: Desired display width in millimeters.
        height_mm: Desired display height in mm; auto-calculated from aspect
                   ratio if not specified.
        align: Horizontal alignment — "left", "center", or "right".
        caption: Optional caption text displayed below the image.
        _bin_id: Internal — assigned during finalize (1-based).
        _pixel_width: Internal — original pixel width read from file.
        _pixel_height: Internal — original pixel height read from file.
    """

    src: str = ""
    width_mm: float = 80.0
    height_mm: Optional[float] = None
    align: str = "center"
    caption: Optional[str] = None

    # Assigned internally during packaging
    _bin_id: int = 0
    _pixel_width: int = 0
    _pixel_height: int = 0

    @property
    def format(self) -> str:
        """Return 'png' or 'jpg' based on file extension."""
        ext = Path(self.src).suffix.lower()
        if ext in (".jpg", ".jpeg"):
            return "jpg"
        return "png"

    @property
    def bin_filename(self) -> str:
        """Return the filename used inside BinData/ folder."""
        return f"image{self._bin_id}.{self.format}"

    @property
    def bin_path(self) -> str:
        """Return the full internal path inside the HWPX ZIP."""
        return f"BinData/{self.bin_filename}"

    def resolve_dimensions(self) -> None:
        """Read the actual pixel dimensions from the image file and compute
        height_mm if not explicitly set.

        Supports PNG (IHDR) and JPEG (SOF markers) without PIL.
        """
        path = Path(self.src)
        if not path.exists():
            raise FileNotFoundError(f"Image file not found: {self.src}")

        data = path.read_bytes()
        ext = path.suffix.lower()

        if ext == ".png":
            w, h = _read_png_dimensions(data)
        elif ext in (".jpg", ".jpeg"):
            w, h = _read_jpeg_dimensions(data)
        else:
            raise ValueError(f"Unsupported image format: {ext}")

        self._pixel_width = w
        self._pixel_height = h

        if self.height_mm is None and w > 0:
            self.height_mm = self.width_mm * h / w

    def resolve_dimensions_from_bytes(self, data: bytes, ext: str) -> None:
        """Read pixel dimensions from raw bytes (for testing without files)."""
        ext = ext.lower().lstrip(".")
        if ext == "png":
            w, h = _read_png_dimensions(data)
        elif ext in ("jpg", "jpeg"):
            w, h = _read_jpeg_dimensions(data)
        else:
            raise ValueError(f"Unsupported image format: {ext}")

        self._pixel_width = w
        self._pixel_height = h

        if self.height_mm is None and w > 0:
            self.height_mm = self.width_mm * h / w


def _read_png_dimensions(data: bytes) -> tuple:
    """Read width and height from PNG IHDR chunk.

    PNG layout: 8-byte signature, then IHDR chunk.
    IHDR data starts at offset 16: 4 bytes width + 4 bytes height (big-endian).
    """
    if len(data) < 24:
        raise ValueError("Invalid PNG: file too small")
    if data[:4] != b'\x89PNG':
        raise ValueError("Invalid PNG signature")

    width = struct.unpack(">I", data[16:20])[0]
    height = struct.unpack(">I", data[20:24])[0]
    return width, height


def _read_jpeg_dimensions(data: bytes) -> tuple:
    """Read width and height from JPEG SOF marker.

    Scans for SOF0 (0xFFC0) through SOF15 markers to find image dimensions.
    """
    if len(data) < 2 or data[:2] != b'\xff\xd8':
        raise ValueError("Invalid JPEG signature")

    offset = 2
    while offset < len(data) - 1:
        if data[offset] != 0xFF:
            offset += 1
            continue

        marker = data[offset + 1]

        # Skip filler bytes
        if marker == 0xFF:
            offset += 1
            continue

        # SOF markers: 0xC0-0xCF except 0xC4 (DHT), 0xC8 (reserved), 0xCC (DAC)
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                      0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            # SOF segment: length(2) + precision(1) + height(2) + width(2)
            if offset + 9 > len(data):
                raise ValueError("JPEG SOF marker truncated")
            height = struct.unpack(">H", data[offset + 5:offset + 7])[0]
            width = struct.unpack(">H", data[offset + 7:offset + 9])[0]
            return width, height

        # Skip this marker segment
        if offset + 3 >= len(data):
            break
        seg_len = struct.unpack(">H", data[offset + 2:offset + 4])[0]
        offset += 2 + seg_len

    raise ValueError("Could not find SOF marker in JPEG")
