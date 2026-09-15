"""Примитивы автоисправления: прямое форматирование прогонов и абзацев, стили, колонтитулы."""

from __future__ import annotations

from typing import Iterable

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Twips
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from . import rules as R
from .ooxml import get_or_add_ppr, insert_in_order, PPR_SEQ

_ALIGN = {
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}
_THEME_FONT_ATTRS = ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme")
_CHAR_IND_ATTRS = (
    "start", "end", "firstLineChars", "hangingChars", "leftChars", "rightChars", "startChars", "endChars",
)


# --- прогоны ---------------------------------------------------------------------------------


def _fonts_element(r):
    return r.get_or_add_rPr().get_or_add_rFonts()


def set_font_name(runs: Iterable, name: str = R.FONT_NAME) -> None:
    for r in runs:
        rf = _fonts_element(r)
        for attr in ("ascii", "hAnsi", "cs"):
            rf.set(qn(f"w:{attr}"), name)
        for attr in _THEME_FONT_ATTRS:
            rf.attrib.pop(qn(f"w:{attr}"), None)
        rf.attrib.pop(qn("w:hint"), None)


def set_font_size(runs: Iterable, size_pt: float = R.FONT_SIZE_PT) -> None:
    for r in runs:
        Run(r, None).font.size = Pt(size_pt)
        szcs = r.rPr.find(qn("w:szCs"))
        if szcs is not None:
            szcs.set(qn("w:val"), str(int(size_pt * 2)))


def set_black(runs: Iterable) -> None:
    for r in runs:
        Run(r, None).font.color.rgb = RGBColor(0, 0, 0)
        color = r.rPr.find(qn("w:color"))
        if color is not None:
            for attr in ("themeColor", "themeTint", "themeShade"):
                color.attrib.pop(qn(f"w:{attr}"), None)


def remove_underline(runs: Iterable) -> None:
    for r in runs:
        Run(r, None).font.underline = False


def set_bold(runs: Iterable, on: bool) -> None:
    for r in runs:
        font = Run(r, None).font
        font.bold = on
        if r.rPr.find(qn("w:bCs")) is not None:
            font.complex_script_bold = on


def caps_off(runs: Iterable) -> None:
    for r in runs:
        rpr = r.get_or_add_rPr()
        for tag in ("w:caps", "w:smallCaps"):
            el = rpr.find(qn(tag))
            if el is not None:
                rpr.remove(el)
        Run(r, None).font.all_caps = False


def normalize_runs(runs: Iterable, size_pt: float | None = R.FONT_SIZE_PT) -> None:
    runs = list(runs)
    set_font_name(runs)
    if size_pt is not None:
        set_font_size(runs, size_pt)
    set_black(runs)
    remove_underline(runs)


# --- абзацы ----------------------------------------------------------------------------------


def _pf(p):
    return Paragraph(p, None).paragraph_format


def set_alignment(p, key: str) -> None:
    _pf(p).alignment = _ALIGN[key]


def set_indent(p, first: int = 0, left: int = 0, right: int = 0) -> None:
    pf = _pf(p)
    pf.left_indent = Twips(left)
    pf.right_indent = Twips(right)
    pf.first_line_indent = Twips(first)
    ind = p.pPr.find(qn("w:ind"))
    if ind is not None:
        for attr in _CHAR_IND_ATTRS:
            ind.attrib.pop(qn(f"w:{attr}"), None)


def set_spacing_zero(p) -> None:
    pf = _pf(p)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    sp = p.pPr.find(qn("w:spacing"))
    if sp is not None:
        for attr in ("beforeAutospacing", "afterAutospacing", "beforeLines", "afterLines"):
            sp.attrib.pop(qn(f"w:{attr}"), None)


def set_line_spacing(p, value: float) -> None:
    _pf(p).line_spacing = value


def set_style(p, style_id: str) -> None:
    ppr = get_or_add_ppr(p)
    ps = ppr.find(qn("w:pStyle"))
    if ps is None:
        ps = OxmlElement("w:pStyle")
        insert_in_order(ppr, ps, PPR_SEQ)
    ps.set(qn("w:val"), style_id)


# --- стили -----------------------------------------------------------------------------------


def _style_ppr_child(style, tag: str, attrs: dict[str, str]) -> None:
    ppr = style.element.get_or_add_pPr()
    el = ppr.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        insert_in_order(ppr, el, PPR_SEQ)
    for k, v in attrs.items():
        el.set(qn(k), v)


def _normalize_style_font(style, bold: bool | None = None) -> None:
    font = style.font
    font.name = R.FONT_NAME
    rpr = style.element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    rf.set(qn("w:cs"), R.FONT_NAME)
    for attr in _THEME_FONT_ATTRS:
        rf.attrib.pop(qn(f"w:{attr}"), None)
    font.size = Pt(R.FONT_SIZE_PT)
    font.color.rgb = RGBColor(0, 0, 0)
    color = rpr.find(qn("w:color"))
    if color is not None:
        for attr in ("themeColor", "themeTint", "themeShade"):
            color.attrib.pop(qn(f"w:{attr}"), None)
    if bold is not None:
        font.bold = bold


def ensure_heading_style(document, level: int) -> str:
    """Возвращает styleId стиля «Заголовок N», при необходимости создавая его."""
    name = f"Heading {level}"
    try:
        style = document.styles[name]
    except KeyError:
        style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        try:
            style.base_style = document.styles["Normal"]
        except KeyError:
            pass
        style.quick_style = True
        _style_ppr_child(style, "w:keepNext", {})
        _style_ppr_child(style, "w:outlineLvl", {"w:val": str(level - 1)})
        _normalize_style_font(style, bold=True)
    return style.style_id


def normalize_styles(document) -> None:
    """Заголовки и оглавление: Times New Roman 14, чёрный — чтобы автосодержание собиралось правильно."""
    for level in (1, 2, 3):
        try:
            _normalize_style_font(document.styles[f"Heading {level}"], bold=True)
        except KeyError:
            pass
    for level in (1, 2, 3):
        try:
            style = document.styles[f"TOC {level}"]
        except KeyError:
            continue
        _normalize_style_font(style)
        style.paragraph_format.first_line_indent = Twips(0)


def request_fields_update(document) -> None:
    """Просит Word обновить поля (содержание) при открытии файла."""
    settings = document.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        settings.append(el)
    el.set(qn("w:val"), "true")


# --- колонтитулы ------------------------------------------------------------------------------


def _field_run(kind: str, payload: str | None = None):
    r = OxmlElement("w:r")
    if kind in ("begin", "separate", "end"):
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), kind)
        r.append(fc)
    elif kind == "instr":
        it = OxmlElement("w:instrText")
        it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        it.text = payload
        r.append(it)
    else:
        t = OxmlElement("w:t")
        t.text = payload
        r.append(t)
    normalize_runs([r])
    return r


def add_page_numbers(document) -> None:
    section = document.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    paragraph = footer.paragraphs[0]
    p = paragraph._p
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)
    set_alignment(p, "center")
    set_indent(p, 0, 0, 0)
    set_spacing_zero(p)
    for kind, payload in (("begin", None), ("instr", " PAGE "), ("separate", None), ("text", "3"), ("end", None)):
        p.append(_field_run(kind, payload))
    section.different_first_page_header_footer = True
