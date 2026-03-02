#!/usr/bin/env python3
# ===== styles_loader.py =====
"""样式加载模块"""
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx.oxml.ns import qn
from utils import rgb

def load_styles(doc: Document, cfg: dict):
    """
    将 JSON 样式配置写入空文档。
    兼容已存在的样式：存在则更新，不存在则创建。
    """
    # ---------------- 字符样式 ----------------
    for cs in cfg.get("character_styles", []):
        style_name = cs["name"]
        try:
            style = doc.styles[style_name]
        except KeyError:
            style = doc.styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)

        font_cfg = cs["font"]
        style.font.bold = font_cfg.get("bold", False)
        if "color" in font_cfg:
            style.font.color.rgb = rgb(font_cfg["color"])

    # ---------------- 段落样式 ----------------
    for ps in cfg.get("paragraph_styles", []):
        style_name = ps["name"]
        try:
            style = doc.styles[style_name]
        except KeyError:
            style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)

        font_cfg = ps["font"]
        style.font.name = font_cfg["name"]
        style.font.size = Pt(font_cfg["size"])
        style.font.bold = font_cfg.get("bold", False)

        # 中文字体
        if hasattr(style, "_element"):
            rPr = style._element.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            rFonts.set(qn("w:eastAsia"), font_cfg["name"])

        # 段落格式
        pf_cfg = ps.get("paragraph_format", {})
        pf = style.paragraph_format
        pf.space_before = Pt(pf_cfg.get("space_before", 0))
        pf.space_after = Pt(pf_cfg.get("space_after", 0))
        pf.line_spacing = pf_cfg.get("line_spacing", 1.25)

        if "left_indent" in pf_cfg:
            pf.left_indent = Pt(pf_cfg["left_indent"])
        if "first_line_indent" in pf_cfg:
            pf.first_line_indent = Pt(pf_cfg["first_line_indent"])

        # 大纲级别
        if "outline_lvl" in ps:
            pPr = style._element.get_or_add_pPr()
            pPr.get_or_add_outlineLvl().val = ps["outline_lvl"]