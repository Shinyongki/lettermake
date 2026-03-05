"""PageBreak element — represents an explicit page break."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class PageBreak:
    """A page break marker.

    When rendered, this produces a paragraph with pageBreakBefore="true"
    in its paraPr, causing a page break in the HWPX output.
    """

    pass
