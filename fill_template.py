# -*- coding: utf-8 -*-
"""
fill_template.py
================
Preenche templates .docx de plano alimentar veterinario (Apsirtus) e retorna
os bytes do documento pronto para conversao via LibreOffice.
"""

from __future__ import annotations

import io
from typing import List, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


class _C:
    TEAL_DARK   = RGBColor(0x1A, 0x52, 0x76)
    TEAL_MED    = RGBColor(0x2E, 0x86, 0xC1)
    TEAL_LIGHT  = RGBColor(0xEB, 0xF5, 0xFB)
    TABLE_HDR   = "1A5276"
    TABLE_ALT   = "EBF5FB"
    TABLE_TOTAL = "D6EAF8"
    WHITE       = "FFFFFF"
    TEXT_DARK   = RGBColor(0x1C, 0x28, 0x33)


_TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
    "tblStyleRowBandSize", "tblStyleColBandSize",
    "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar",
    "tblLook", "tblCaption", "tblPrChange",
]

_TCPR_ORDER = [
    "cnfStyle", "tcW", "gridSpan", "vMerge",
    "tcBorders", "shd", "noWrap", "tcMar",
    "textDirection", "tcFitText", "vAlign",
    "hideMark", "headers", "cellIns", "cellDel", "trPr",
]


def _insert_ordered(parent, new_el, order: list) -> None:
    ns_tag = new_el.tag
    local  = ns_tag.split("}")[1] if "}" in ns_tag else ns_tag
    for old in parent.findall(ns_tag):
        parent.remove(old)
    new_pos = order.index(local) if local in order else len(order)
    insert_before = None
    for child in list(parent):
        child_local = child.tag.split("}")[1] if "}" in child.tag else child.tag
        child_pos = order.index(child_local) if child_local in order else len(order)
        if child_pos > new_pos:
            insert_before = child
            break
    if insert_before is not None:
        insert_before.addprevious(new_el)
    else:
        parent.append(new_el)


def _set_cell_bg(cell, hex_color: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    _insert_ordered(tcPr, shd, _TCPR_ORDER)


def _set_cell_width(cell, width_dxa: int) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"),    str(width_dxa))
    tcW.set(qn("w:type"), "dxa")
    _insert_ordered(tcPr, tcW, _TCPR_ORDER)


def _set_cell_margins(cell, top=80, bottom=80, left=140, right=140) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("left", left),
                      ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    _insert_ordered(tcPr, tcMar, _TCPR_ORDER)


def _set_cell_borders(cell, color: str = "CCCCCC") -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    borders_el = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "single")
        b.set(qn("w:sz"),    "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), color)
        borders_el.append(b)
    _insert_ordered(tcPr, borders_el, _TCPR_ORDER)


def _cell_valign(cell, val: str = "center") -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), val)
    _insert_ordered(tcPr, vAlign, _TCPR_ORDER)


def _set_row_height(row, twips: int, rule: str = "atLeast") -> None:
    trPr = row._tr.get_or_add_trPr()
    trH = OxmlElement("w:trHeight")
    trH.set(qn("w:val"),   str(twips))
    trH.set(qn("w:hRule"), rule)
    trPr.append(trH)


def _para_space(para, before=0, after=0) -> None:
    pPr = para._element.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:before"), str(before))
    spacing.set(qn("w:after"),  str(after))


def _set_table_width(table, width_dxa: int) -> None:
    tbl   = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"),    str(width_dxa))
    tblW.set(qn("w:type"), "dxa")
    _insert_ordered(tblPr, tblW, _TBLPR_ORDER)


def _remove_table_borders(table) -> None:
    tbl   = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        return
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "none")
        b.set(qn("w:sz"),    "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        borders.append(b)
    _insert_ordered(tblPr, borders, _TBLPR_ORDER)


def _replace_in_paragraph(para, placeholder: str, value: str) -> None:
    for run in para.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, value)
            return
    full = "".join(r.text for r in para.runs)
    if placeholder in full:
        new_text = full.replace(placeholder, value)
        if para.runs:
            para.runs[0].text = new_text
            for run in para.runs[1:]:
                run.text = ""


def _replace_all_placeholders(doc: Document, replacements: dict) -> None:
    def process_para(p):
        for key, val in replacements.items():
            ph = "{{" + key + "}}"
            if ph in p.text:
                _replace_in_paragraph(p, ph, str(val))

    for p in doc.paragraphs:
        process_para(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    process_para(p)

    # Escaneia rodape de secao (footer fixo com {{data}})
    for section in doc.sections:
        try:
            footer = section.footer
            for p in footer.paragraphs:
                process_para(p)
            for table in footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            process_para(p)
        except Exception:
            pass


def _build_ingredients_table(
    doc: Document,
    ingredients: List[Tuple[str, str]],
    total: str,
) -> object:
    CONTENT_W = 9638
    COL_INGR  = int(CONTENT_W * 0.745)
    COL_QTD   = CONTENT_W - COL_INGR

    table = doc.add_table(rows=0, cols=2)
    _set_table_width(table, CONTENT_W)
    _remove_table_borders(table)

    # Cabecalho
    hdr_row = table.add_row()
    _set_row_height(hdr_row, 640)
    for idx, (cell, header_text, col_w) in enumerate(zip(
        hdr_row.cells,
        ["Ingredientes", "Qtd (g)"],
        [COL_INGR, COL_QTD],
    )):
        cell.text = ""
        _set_cell_bg(cell, _C.TABLE_HDR)
        _set_cell_borders(cell, _C.TABLE_HDR)
        _set_cell_width(cell, col_w)
        _set_cell_margins(cell, top=0, bottom=0, left=160, right=160)
        _cell_valign(cell, "center")
        p = cell.paragraphs[0]
        p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if idx == 1
                       else WD_ALIGN_PARAGRAPH.LEFT)
        _para_space(p, before=0, after=0)
        run = p.add_run(header_text)
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Linhas de ingredientes
    for i, (ingredient, qty) in enumerate(ingredients):
        bg = _C.WHITE
        row = table.add_row()
        _set_row_height(row, 700)

        c0 = row.cells[0]
        c0.text = ""
        _set_cell_bg(c0, bg)
        _set_cell_borders(c0, "D5E8F8")
        _set_cell_width(c0, COL_INGR)
        _set_cell_margins(c0, top=0, bottom=0, left=160, right=160)
        _cell_valign(c0, "center")
        p0 = c0.paragraphs[0]
        _para_space(p0, before=0, after=0)
        r0 = p0.add_run(str(ingredient))
        r0.font.name = "Arial"
        r0.font.size = Pt(11)
        r0.font.color.rgb = _C.TEXT_DARK

        c1 = row.cells[1]
        c1.text = ""
        _set_cell_bg(c1, bg)
        _set_cell_borders(c1, "D5E8F8")
        _set_cell_width(c1, COL_QTD)
        _set_cell_margins(c1, top=0, bottom=0, left=160, right=160)
        _cell_valign(c1, "center")
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _para_space(p1, before=0, after=0)
        r1 = p1.add_run(str(qty))
        r1.font.name = "Arial"
        r1.font.size = Pt(11)
        r1.font.color.rgb = _C.TEXT_DARK

    # Linha TOTAL
    tot_row = table.add_row()
    _set_row_height(tot_row, 720)

    ct0 = tot_row.cells[0]
    ct0.text = ""
    _set_cell_bg(ct0, _C.TABLE_TOTAL)
    _set_cell_borders(ct0, _C.TABLE_HDR)
    _set_cell_width(ct0, COL_INGR)
    _set_cell_margins(ct0, top=0, bottom=0, left=160, right=160)
    _cell_valign(ct0, "center")
    pt0 = ct0.paragraphs[0]
    _para_space(pt0, before=0, after=0)
    rt0 = pt0.add_run("TOTAL")
    rt0.bold = True
    rt0.font.name = "Arial"
    rt0.font.size = Pt(11)
    rt0.font.color.rgb = _C.TEAL_DARK

    ct1 = tot_row.cells[1]
    ct1.text = ""
    _set_cell_bg(ct1, _C.TABLE_TOTAL)
    _set_cell_borders(ct1, _C.TABLE_HDR)
    _set_cell_width(ct1, COL_QTD)
    _set_cell_margins(ct1, top=0, bottom=0, left=160, right=160)
    _cell_valign(ct1, "center")
    pt1 = ct1.paragraphs[0]
    pt1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _para_space(pt1, before=0, after=0)
    rt1 = pt1.add_run(str(total))
    rt1.bold = True
    rt1.font.name = "Arial"
    rt1.font.size = Pt(11)
    rt1.font.color.rgb = _C.TEAL_DARK

    return table._tbl


def fill_template(
    docx_path: str,
    data: dict,
    ingredients: List[Tuple[str, str]],
    total: str,
    opcao_num: str,
) -> bytes:
    doc = Document(docx_path)

    replacements = dict(data)
    replacements["opcao_num"] = str(opcao_num)

    _replace_all_placeholders(doc, replacements)

    sample_table = None
    for t in doc.tables:
        try:
            if t.rows[0].cells[0].text.strip() == "Ingredientes":
                sample_table = t
                break
        except (IndexError, AttributeError):
            continue

    if sample_table is None:
        raise ValueError(
            "Tabela de ingredientes nao encontrada no template. "
            "Certifique-se de usar um template gerado por create_templates.py."
        )

    tbl_el = _build_ingredients_table(doc, ingredients, total)

    sample_tbl_el = sample_table._tbl
    sample_tbl_el.addprevious(tbl_el)
    sample_tbl_el.getparent().remove(sample_tbl_el)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()
