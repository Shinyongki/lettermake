"""Government document style preset (공문서 표준 스타일 프리셋)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Any, Optional


@dataclass
class BulletStyle:
    """Style definition for a single bullet level."""

    prefix: str = ""
    bold: bool = False
    italic: bool = False
    color: str = "000000"
    size: float = 13.0            # font size in pt
    font_name: Optional[str] = None  # None = use body font
    indent_mm: float = 0.0       # left indent
    hanging_mm: float = 0.0      # hanging indent (first line outdented)
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    keep_with_next: bool = False
    align: str = "JUSTIFY"
    # Separate prefix styling (if set, prefix rendered as its own run)
    prefix_font_name: Optional[str] = None
    prefix_size: Optional[float] = None


@dataclass
class GovDocumentPreset:
    """Government document preset with sensible defaults.

    Usage::

        preset = GovDocumentPreset(font_body="휴먼명조", size_body=13)
        doc.apply_preset(preset)
    """

    # Fonts
    font_body: str = "휴먼명조"
    font_table: str = "맑은 고딕"
    font_title: Optional[str] = None  # None = use font_body

    # Sizes (pt)
    size_title: float = 15.0
    size_body: float = 13.0
    size_table: float = 10.0

    # Line spacing
    line_spacing: float = 160.0

    # Margins (mm)
    margin_top: float = 20.0
    margin_bottom: float = 20.0
    margin_left: float = 20.0
    margin_right: float = 20.0

    # Cell margins (mm)
    cell_margin_lr: float = 3.0
    cell_margin_tb: float = 2.0

    # Colors
    color_main: str = "1F4E79"
    color_sub: str = "2E75B6"
    color_accent: str = "C00000"

    # Bullet styles (initialized in __post_init__)
    bullet_styles: Dict[str, BulletStyle] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not self.bullet_styles:
            self.bullet_styles = self._default_bullet_styles()

    def _default_bullet_styles(self) -> Dict[str, BulletStyle]:
        return {
            "title": BulletStyle(
                prefix="",
                bold=True,
                color="000000",
                size=self.size_title,
                font_name=self.font_title or self.font_body,
                align="CENTER",
                space_after_pt=6.0,
            ),
            "h1": BulletStyle(
                prefix="■ {num}. ",
                bold=True,
                color=self.color_main,
                size=self.size_body,
                space_before_pt=6.0,
                space_after_pt=3.0,
                keep_with_next=True,
            ),
            "h2": BulletStyle(
                prefix="\uf06d ",
                bold=True,
                color="000000",
                size=self.size_body,
                indent_mm=4.0,
                space_before_pt=3.0,
                keep_with_next=True,
                prefix_font_name="휴먼명조",
                prefix_size=15.0,
            ),
            "sub1": BulletStyle(
                prefix="{num}. ",
                bold=False,
                color="000000",
                size=self.size_body,
                indent_mm=7.0,
                hanging_mm=4.0,
                keep_with_next=True,
            ),
            "li1": BulletStyle(
                prefix="\uf06d ",  # 동그라미 특수문자, 휴먼명조 15pt
                bold=False,
                color="000000",
                size=self.size_body,
                prefix_font_name="휴먼명조",
                prefix_size=15.0,
            ),
            "li2": BulletStyle(
                prefix="- ",
                bold=False,
                color="000000",
                size=self.size_body,
            ),
            "li3": BulletStyle(
                prefix="\u2219 ",  # ∙ U+2219 bullet dot
                bold=False,
                color="000000",
                size=self.size_body,
                font_name="휴먼명조",
            ),
            "li4": BulletStyle(
                prefix="{num}. ",
                bold=False,
                color="000000",
                size=self.size_body,
                font_name="휴먼명조",
            ),
            "li5": BulletStyle(
                prefix="{num}. ",
                bold=False,
                color="000000",
                size=self.size_body,
                font_name="휴먼명조",
            ),
            "note": BulletStyle(
                prefix="※ ",
                bold=True,
                color=self.color_accent,
                size=self.size_body - 2,
                indent_mm=4.0,
                space_before_pt=3.0,
            ),
        }
