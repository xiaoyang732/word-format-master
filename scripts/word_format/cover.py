# -*- coding: utf-8 -*-
"""
高校学术学位论文/毕业设计通用精美封面模板模块
通用规范、排版美观合理，仅展示必要学术信息（标题、作者、学号、学院、专业、指导教师、日期等）。
无绑定特定高校硬编码，排版疏密得体，适用于演示与初版通用规范。
"""

from __future__ import annotations
from typing import List, Tuple
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn


def apply_east_asia_font(run, font_name: str = "宋体", ascii_font: str = "Times New Roman"):
    """为 run 节点精确指定中西文字体分离"""
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is not None:
        rPr.remove(rFonts)
    rPr.append(parse_xml(
        f'<w:rFonts {nsdecls("w")} w:eastAsia="{font_name}" w:ascii="{ascii_font}" w:hAnsi="{ascii_font}"/>'
    ))


def add_standard_thesis_cover(
    doc: Document,
    *,
    university_name: str | None = None,
    degree_type: str = "毕 业 设 计（论 文）",
    title_cn: str = "基于检索增强生成（RAG）的领域知识问答系统设计与实现",
    title_en: str | None = "Design and Implementation of Domain Knowledge Question Answering System Based on Retrieval-Augmented Generation",
    fields: List[Tuple[str, str]] | None = None,
    date_str: str = "二〇二六年五月",
    student_id: str = "2022010101",
    author_name: str = "张    三",
    college: str = "计算机科学与技术学院",
    major: str = "计算机科学与技术",
    advisor: str = "李四  教授",
    use_section_break: bool = True,
) -> None:
    """
    构建通用、规范、美观大气的学术毕业设计（论文）封面。
    不绑定任何特定学校，展示必要核心信息，排版疏密适宜。
    """
    if doc.sections:
        doc.sections[0].different_first_page_header_footer = True

    if fields is None:
        fields = [
            ("学生姓名：", author_name),
            ("学　　号：", student_id),
            ("所在学院：", college),
            ("专业名称：", major),
            ("指导教师：", advisor),
        ]

    # 1. 顶部留白与学校名称（如有传入则显示，默认优雅留白）
    p_top = doc.add_paragraph()
    p_top.paragraph_format.first_line_indent = Pt(0)
    p_top.paragraph_format.space_before = Pt(40 if university_name else 60)
    p_top.paragraph_format.space_after = Pt(0)

    if university_name:
        p_univ = doc.add_paragraph()
        p_univ.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_univ.paragraph_format.space_before = Pt(0)
        p_univ.paragraph_format.space_after = Pt(12)
        p_univ.paragraph_format.first_line_indent = Pt(0)
        r_univ = p_univ.add_run(university_name)
        r_univ.font.size = Pt(26)
        r_univ.font.bold = True
        apply_east_asia_font(r_univ, "华文中宋", "Times New Roman")

    # 2. 论文类别醒目标题（如：毕 业 设 计（论 文））
    p_type = doc.add_paragraph()
    p_type.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_type.paragraph_format.space_before = Pt(0)
    p_type.paragraph_format.space_after = Pt(45)
    p_type.paragraph_format.first_line_indent = Pt(0)
    r_type = p_type.add_run(degree_type)
    r_type.font.size = Pt(24)
    r_type.font.bold = True
    apply_east_asia_font(r_type, "黑体", "Times New Roman")

    # 3. 论文题目区域（主标题与可选英文副标题）
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.line_spacing = 1.3
    p_title.paragraph_format.space_before = Pt(10)
    p_title.paragraph_format.space_after = Pt(14 if title_en else 40)
    p_title.paragraph_format.first_line_indent = Pt(0)
    # 清理边框
    pPr = p_title._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is not None:
        pPr.remove(pBdr)
    r_title = p_title.add_run(title_cn)
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0, 0, 0)
    apply_east_asia_font(r_title, "黑体", "Times New Roman")

    if title_en and title_en.strip():
        p_etitle = doc.add_paragraph()
        p_etitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_etitle.paragraph_format.line_spacing = 1.2
        p_etitle.paragraph_format.space_before = Pt(0)
        p_etitle.paragraph_format.space_after = Pt(45)
        p_etitle.paragraph_format.first_line_indent = Pt(0)
        pPr_e = p_etitle._p.get_or_add_pPr()
        pBdr_e = pPr_e.find(qn("w:pBdr"))
        if pBdr_e is not None:
            pPr_e.remove(pBdr_e)
        r_etitle = p_etitle.add_run(title_en.strip())
        r_etitle.font.size = Pt(12)
        r_etitle.font.bold = True
        r_etitle.font.color.rgb = RGBColor(0, 0, 0)
        apply_east_asia_font(r_etitle, "Times New Roman", "Times New Roman")

    # 4. 学生及指导教师填报网格（双列下划线居中表格）
    grid = doc.add_table(rows=len(fields), cols=2)
    grid.alignment = WD_TABLE_ALIGNMENT.CENTER
    grid._tbl.tblPr.set(qn("w:tblLayout"), "fixed")

    # 清空外部与内部网格线
    tblPr_g = grid._tbl.tblPr
    tblBorders_g = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="none"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="none"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr_g.append(tblBorders_g)

    grid.columns[0].width = Cm(4.2)
    grid.columns[1].width = Cm(8.5)

    for r_idx, (label, val) in enumerate(fields):
        row = grid.rows[r_idx]
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]

        # 标签列：靠右对齐、加粗、小三号（14-15pt）宋体
        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_lbl.paragraph_format.line_spacing = Pt(28)
        p_lbl.paragraph_format.space_before = Pt(4)
        p_lbl.paragraph_format.space_after = Pt(4)
        p_lbl.paragraph_format.first_line_indent = Pt(0)
        r_lbl = p_lbl.add_run(label)
        r_lbl.font.size = Pt(14)
        r_lbl.font.bold = True
        apply_east_asia_font(r_lbl, "宋体", "Times New Roman")

        # 填空列：居中对齐、小三号宋体、底边单实线下划线
        p_val = cell_val.paragraphs[0]
        p_val.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_val.paragraph_format.line_spacing = Pt(28)
        p_val.paragraph_format.space_before = Pt(4)
        p_val.paragraph_format.space_after = Pt(4)
        p_val.paragraph_format.first_line_indent = Pt(0)
        r_val = p_val.add_run(val)
        r_val.font.size = Pt(14)
        apply_east_asia_font(r_val, "宋体", "Times New Roman")

        # 设置单元格仅底部单实线下划线（0.75磅）
        tcPr = cell_val._tc.get_or_add_tcPr()
        tcBorders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/>'
            f'<w:left w:val="none"/>'
            f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'<w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(tcBorders)

    # 5. 完成日期（底部居中）
    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_date.paragraph_format.space_before = Pt(48)
    p_date.paragraph_format.space_after = Pt(0)
    p_date.paragraph_format.first_line_indent = Pt(0)
    r_date = p_date.add_run(date_str)
    r_date.font.size = Pt(14)
    r_date.font.bold = True
    apply_east_asia_font(r_date, "宋体", "Times New Roman")

    # 6. 封面结束分节或分页
    if use_section_break:
        # 插入下一页分节符，使正文/摘要与封面彻底分节隔离
        sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
        # 默认解除与封面页眉页脚的继承关联
        sec.header.is_linked_to_previous = False
        sec.footer.is_linked_to_previous = False
    else:
        doc.add_page_break()
