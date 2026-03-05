"""Generate Contents/header.xml — based on real HWPX file structure analysis.

Key differences from initial implementation:
- hh:refList instead of hh:mappingTable
- hh:fontfaces with hh:fontface lang= + hh:font children
- hh:charProperties with fontRef using lang attrs (hangul/latin/...)
- hh:paraProperties with hp:switch/case/default for margin
- hh:compatibleDocument required
- hh:beginNum required
- Minimum 17 default styles required
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List
from xml.sax.saxutils import escape

from ..utils import mm_to_hwp, pt_to_hwp
from ..styles import StyleManager, CharProps, ParaProps, BorderFillProps, Style, RESERVED_DOC_TITLE_PP_ID

if TYPE_CHECKING:
    from ..document import HwpxDocument

LANGS = ("HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER")

# Namespace URIs for hp:switch
NS_HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS_HWPUNITCHAR = "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"


def build_header_xml(doc: HwpxDocument) -> bytes:
    """Build Contents/header.xml with styles derived from document content."""
    sm = doc.style_manager
    L: List[str] = []

    L.append('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>')
    L.append(
        '<hh:head'
        ' xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"'
        f' xmlns:hp="{NS_HP}"'
        ' xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph"'
        ' xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"'
        ' xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"'
        ' xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"'
        ' xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history"'
        ' xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page"'
        ' xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"'
        ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
        ' xmlns:opf="http://www.idpf.org/2007/opf/"'
        ' xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart"'
        f' xmlns:hwpunitchar="{NS_HWPUNITCHAR}"'
        ' xmlns:epub="http://www.idpf.org/2007/ops"'
        ' xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
        ' version="1.5" secCnt="1">'
    )

    # --- beginNum ---
    L.append('<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>')

    # --- refList ---
    L.append('<hh:refList>')

    # ── fontfaces ──
    font_cnt = len(sm._face_names) or 1
    L.append(f'<hh:fontfaces itemCnt="7">')
    for lang in LANGS:
        L.append(f'<hh:fontface lang="{lang}" fontCnt="{font_cnt}">')
        for i, name in enumerate(sm._face_names):
            L.append(
                f'<hh:font id="{i}" face="{escape(name)}" type="TTF" isEmbedded="0">'
                f'<hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="4"'
                f' contrast="0" strokeVariation="1" armStyle="1"'
                f' letterform="1" midline="1" xHeight="1"/>'
                f'</hh:font>'
            )
        L.append(f'</hh:fontface>')
    L.append('</hh:fontfaces>')

    # ── borderFills ──
    # Always start with 2 default borderFills (matching sample structure)
    user_bf_cnt = len(sm._border_fills)
    bf_cnt = 2 + user_bf_cnt
    L.append(f'<hh:borderFills itemCnt="{bf_cnt}">')
    # Default borderFill id=1: no border, no fill
    L.append(
        '<hh:borderFill id="1" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
        '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
        '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
        '<hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:topBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
        '</hh:borderFill>'
    )
    # Default borderFill id=2: no border, transparent fill (referenced by charPr/paraPr)
    L.append(
        '<hh:borderFill id="2" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
        '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
        '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
        '<hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:topBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
        '<hc:fillBrush><hc:winBrush faceColor="none" hatchColor="#999999" alpha="0"/></hc:fillBrush>'
        '</hh:borderFill>'
    )
    # User-defined borderFills (id starts from 3)
    for i, bf in enumerate(sm._border_fills):
        bf_id = i + 3
        color = f"#{bf.border_color}" if not bf.border_color.startswith("#") else bf.border_color
        L.append(f'<hh:borderFill id="{bf_id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">')
        L.append('<hh:slash type="NONE" Crooked="0" isCounter="0"/>')
        L.append('<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>')

        # Per-side borders (extended) or uniform borders
        side_map = {
            "left": bf.left_border, "right": bf.right_border,
            "top": bf.top_border, "bottom": bf.bottom_border,
        }
        for side_name in ("left", "right", "top", "bottom"):
            override = side_map[side_name]
            if override:
                s_type, s_width, s_color = override
                L.append(
                    f'<hh:{side_name}Border type="{s_type}"'
                    f' width="{s_width}" color="{s_color}"/>'
                )
            else:
                L.append(
                    f'<hh:{side_name}Border type="{bf.border_type}"'
                    f' width="{bf.border_width}" color="{color}"/>'
                )

        L.append('<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>')

        # Fill: custom XML takes priority, then simple bg_color
        if bf.fill_xml:
            L.append(bf.fill_xml)
        elif bf.bg_color:
            bg = f"#{bf.bg_color}" if not bf.bg_color.startswith("#") else bf.bg_color
            L.append(
                f'<hc:fillBrush>'
                f'<hc:winBrush faceColor="{bg}" hatchColor="#FFFFFF" alpha="0"/>'
                f'</hc:fillBrush>'
            )
        L.append('</hh:borderFill>')
    L.append('</hh:borderFills>')

    # ── charProperties ──
    cp_cnt = len(sm._char_prs) or 1
    L.append(f'<hh:charProperties itemCnt="{cp_cnt}">')
    for i, cp in enumerate(sm._char_prs):
        height = pt_to_hwp(cp.font_size_pt)
        color = f"#{cp.color}" if not cp.color.startswith("#") else cp.color
        face_id = sm.get_face_id(cp.font_name)

        # Build fontRef with per-lang font ID
        fref_attrs = " ".join(f'{lang.lower()}="{face_id}"' for lang in LANGS)
        ratio_attrs = " ".join(f'{lang.lower()}="100"' for lang in LANGS)
        spacing_attrs = " ".join(f'{lang.lower()}="{cp.char_spacing}"' for lang in LANGS)

        # Underline
        ul_type = "BOTTOM" if cp.underline else "NONE"
        # Strikeout
        strike_shape = "SLASH" if cp.strike else "NONE"

        L.append(
            f'<hh:charPr id="{i}" height="{height}" textColor="{color}"'
            f' shadeColor="none" useFontSpace="0" useKerning="0"'
            f' symMark="NONE" borderFillIDRef="2">'
        )
        L.append(f'<hh:fontRef {fref_attrs}/>')
        L.append(f'<hh:ratio {ratio_attrs}/>')
        L.append(f'<hh:spacing {spacing_attrs}/>')
        L.append(f'<hh:relSz {ratio_attrs}/>')
        L.append(f'<hh:offset {spacing_attrs}/>')
        if cp.bold:
            L.append(f'<hh:bold/>')
        if cp.italic:
            L.append(f'<hh:italic/>')
        L.append(f'<hh:underline type="{ul_type}" shape="SOLID" color="#000000"/>')
        L.append(f'<hh:strikeout shape="{strike_shape}" color="#000000"/>')
        L.append('<hh:outline type="NONE"/>')
        L.append('<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>')
        L.append('</hh:charPr>')
    L.append('</hh:charProperties>')

    # ── tabProperties ──
    L.append('<hh:tabProperties itemCnt="1">')
    L.append('<hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>')
    L.append('</hh:tabProperties>')

    # ── numberings ──
    L.append('<hh:numberings itemCnt="1">')
    L.append('<hh:numbering id="1" start="0">')
    L.append('<hh:paraHead start="1" level="1" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" charPrIDRef="4294967295" checkable="0">^1.</hh:paraHead>')
    L.append('<hh:paraHead start="1" level="2" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="HANGUL_SYLLABLE" charPrIDRef="4294967295" checkable="0">^2.</hh:paraHead>')
    L.append('<hh:paraHead start="1" level="3" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" charPrIDRef="4294967295" checkable="0">^3)</hh:paraHead>')
    L.append('<hh:paraHead start="1" level="4" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="HANGUL_SYLLABLE" charPrIDRef="4294967295" checkable="0">^4)</hh:paraHead>')
    L.append('<hh:paraHead start="1" level="5" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" charPrIDRef="4294967295" checkable="0">(^5)</hh:paraHead>')
    L.append('<hh:paraHead start="1" level="6" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="HANGUL_SYLLABLE" charPrIDRef="4294967295" checkable="0">(^6)</hh:paraHead>')
    L.append('<hh:paraHead start="1" level="7" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="CIRCLED_DIGIT" charPrIDRef="4294967295" checkable="1">^7</hh:paraHead>')
    L.append('<hh:paraHead start="1" level="8" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="CIRCLED_HANGUL_SYLLABLE" charPrIDRef="4294967295" checkable="1">^8</hh:paraHead>')
    L.append('<hh:paraHead start="1" level="9" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="HANGUL_JAMO" charPrIDRef="4294967295" checkable="0"/>')
    L.append('<hh:paraHead start="1" level="10" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="ROMAN_SMALL" charPrIDRef="4294967295" checkable="1"/>')
    L.append('</hh:numbering>')
    L.append('</hh:numberings>')

    # ── paraProperties ──
    # IDs must be contiguous 0..N for Hangul to resolve paraPrIDRef correctly.
    # Dynamic entries fill 0..(N-1) skipping RESERVED_DOC_TITLE_PP_ID,
    # then padding entries fill gaps up to RESERVED_DOC_TITLE_PP_ID.
    _used_ids = set(sm.get_para_pr_id(pp) for pp in sm._para_prs)
    _pad_ids = [i for i in range(RESERVED_DOC_TITLE_PP_ID) if i not in _used_ids]
    pp_cnt = len(sm._para_prs) + len(_pad_ids) + 1  # dynamic + padding + reserved
    L.append(f'<hh:paraProperties itemCnt="{pp_cnt}">')

    def _emit_para_pr(L, pp_id, pp):
        indent = mm_to_hwp(pp.indent_first_mm)
        left = mm_to_hwp(pp.indent_left_mm)
        right = mm_to_hwp(pp.indent_right_mm)
        kwn = "1" if pp.keep_with_next else "0"
        wo = "1" if pp.widow_orphan else "0"
        pbb = "1" if pp.page_break_before else "0"

        if pp.line_spacing_type == "PERCENT":
            ls_value = str(int(pp.line_spacing_value))
        else:
            ls_value = str(mm_to_hwp(pp.line_spacing_value))

        sp_before = int(pp.space_before_pt * 100)
        sp_after = int(pp.space_after_pt * 100)

        L.append(
            f'<hh:paraPr id="{pp_id}" tabPrIDRef="0" condense="0"'
            f' fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" checked="0">'
        )
        L.append(f'<hh:align horizontal="{pp.align}" vertical="BASELINE"/>')
        L.append('<hh:heading type="NONE" idRef="0" level="0"/>')
        L.append(
            f'<hh:breakSetting breakLatinWord="KEEP_WORD"'
            f' breakNonLatinWord="KEEP_WORD" widowOrphan="{wo}"'
            f' keepWithNext="{kwn}" keepLines="0"'
            f' pageBreakBefore="{pbb}" lineWrap="BREAK"/>'
        )
        L.append('<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>')

        margin_xml = (
            f'<hh:margin>'
            f'<hc:intent value="{indent}" unit="HWPUNIT"/>'
            f'<hc:left value="{left}" unit="HWPUNIT"/>'
            f'<hc:right value="{right}" unit="HWPUNIT"/>'
            f'<hc:prev value="{sp_before}" unit="HWPUNIT"/>'
            f'<hc:next value="{sp_after}" unit="HWPUNIT"/>'
            f'</hh:margin>'
            f'<hh:lineSpacing type="{pp.line_spacing_type}" value="{ls_value}" unit="HWPUNIT"/>'
        )

        L.append('<hp:switch>')
        L.append(f'<hp:case hp:required-namespace="{NS_HWPUNITCHAR}">')
        L.append(margin_xml)
        L.append('</hp:case>')
        L.append('<hp:default>')
        L.append(margin_xml)
        L.append('</hp:default>')
        L.append('</hp:switch>')

        L.append(
            '<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0"'
            ' offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
        )
        L.append('</hh:paraPr>')

    # Emit dynamic paraPrs (IDs skip RESERVED_DOC_TITLE_PP_ID)
    for pp in sm._para_prs:
        pp_id = sm.get_para_pr_id(pp)
        _emit_para_pr(L, pp_id, pp)

    # Pad missing IDs (0..9) with default ParaProps to keep IDs contiguous
    _default_pp = ParaProps()  # 160% JUSTIFY
    for pad_id in sorted(_pad_ids):
        _emit_para_pr(L, pad_id, _default_pp)

    # Reserved paraPr id=10: DocTitle (CENTER, 120% line spacing) — always present
    _emit_para_pr(L, RESERVED_DOC_TITLE_PP_ID,
                  ParaProps(align="CENTER", line_spacing_value=120.0))

    L.append('</hh:paraProperties>')

    # ── styles ──
    # Build required default styles list
    _emit_styles(L, sm)

    L.append('</hh:refList>')

    # --- compatibleDocument ---
    L.append(
        '<hh:compatibleDocument targetProgram="HWP201X">'
        '<hh:layoutCompatibility/>'
        '</hh:compatibleDocument>'
    )

    # --- docOption ---
    L.append('<hh:docOption><hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/></hh:docOption>')

    # --- trackchangeConfig ---
    L.append('<hh:trackchageConfig flags="56"/>')

    L.append('</hh:head>')

    return "\n".join(L).encode("utf-8")


def _emit_styles(L: List[str], sm: StyleManager) -> None:
    """Emit hh:styles with required Korean default styles."""
    # Required default styles (minimum set for HWP compatibility)
    DEFAULT_STYLES = [
        (0, "PARA", "바탕글", "Normal"),
        (1, "PARA", "본문", "Body"),
        (2, "PARA", "개요 1", "Outline 1"),
        (3, "PARA", "개요 2", "Outline 2"),
        (4, "PARA", "개요 3", "Outline 3"),
        (5, "PARA", "개요 4", "Outline 4"),
        (6, "PARA", "개요 5", "Outline 5"),
        (7, "PARA", "개요 6", "Outline 6"),
        (8, "PARA", "개요 7", "Outline 7"),
        (9, "PARA", "개요 8", "Outline 8"),
        (10, "PARA", "개요 9", "Outline 9"),
        (11, "PARA", "개요 10", "Outline 10"),
        (12, "CHAR", "쪽 번호", "Page Number"),
        (13, "PARA", "머리말", "Header"),
        (14, "PARA", "각주", "Footnote"),
        (15, "PARA", "미주", "Endnote"),
        (16, "PARA", "메모", "Memo"),
        (17, "PARA", "차례 제목", "TOC Heading"),
        (18, "PARA", "차례 1", "TOC 1"),
        (19, "PARA", "차례 2", "TOC 2"),
        (20, "PARA", "차례 3", "TOC 3"),
        (21, "PARA", "캡션", "Caption"),
    ]

    # Count: defaults + any extra user styles
    user_extra = sum(1 for s in sm._style_list if s.name not in [ds[2] for ds in DEFAULT_STYLES])
    style_cnt = len(DEFAULT_STYLES) + user_extra
    L.append(f'<hh:styles itemCnt="{style_cnt}">')

    # Emit defaults
    emitted_ids = set()
    for sid, stype, name, eng_name in DEFAULT_STYLES:
        # Check if there's a user-registered style that overrides this ID
        cp_ref = 0
        pp_ref = 0
        if name in sm._style_id_map:
            style_obj = sm._styles.get(name)
            if style_obj:
                cp_ref = sm.get_char_pr_id(style_obj.char_props)
                pp_ref = sm.get_para_pr_id(style_obj.para_props)

        next_ref = sid  # self-referencing for next style
        L.append(
            f'<hh:style id="{sid}" type="{stype}" name="{escape(name)}"'
            f' engName="{eng_name}"'
            f' paraPrIDRef="{pp_ref}" charPrIDRef="{cp_ref}"'
            f' nextStyleIDRef="{next_ref}" langID="1042" lockForm="0"/>'
        )
        emitted_ids.add(sid)

    # Emit any additional user styles beyond the defaults
    next_id = len(DEFAULT_STYLES)
    for style in sm._style_list:
        if style.name in [ds[2] for ds in DEFAULT_STYLES]:
            continue  # Already emitted
        cp_id = sm.get_char_pr_id(style.char_props)
        pp_id = sm.get_para_pr_id(style.para_props)
        L.append(
            f'<hh:style id="{next_id}" type="PARA" name="{escape(style.name)}"'
            f' engName="" paraPrIDRef="{pp_id}" charPrIDRef="{cp_id}"'
            f' nextStyleIDRef="{next_id}" langID="1042" lockForm="0"/>'
        )
        next_id += 1

    L.append('</hh:styles>')
