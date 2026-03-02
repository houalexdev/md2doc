#!/usr/bin/env python3
# ===== utils.py =====
"""工具函数"""
from docx.shared import RGBColor

def rgb(hex_string: str):
    """将十六进制颜色转换为 python-docx 可用的 RGBColor。"""
    return RGBColor.from_string(hex_string)
