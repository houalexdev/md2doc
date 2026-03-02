#!/usr/bin/env python3
# ===== md_parser.py =====
"""Markdown 解析模块"""
import re

def parse_md(md: str):
    """
    极简 Markdown 解析器。
    支持：标题、加粗、表格、列表项、普通段落、代码块。
    """
    lines = md.strip().splitlines()
    n = len(lines)
    i = 0

    in_code = False
    code_lang = ""
    code_lines = []

    while i < n:
        line = lines[i]

        # ---------- 代码块开始/结束 ----------
        m = re.match(r"^```(?:(\w+))?", line)
        if m:
            if not in_code:
                in_code = True
                code_lang = m.group(1) or ""
                code_lines = []
            else:
                in_code = False
                yield "codeblock", code_lang, "\n".join(code_lines)
            i += 1
            continue

        if in_code:
            code_lines.append(line)
            i += 1
            continue
        
        # ---------- 标题（# ~ #####） ----------
        m = re.match(r"^(#{1,5})\s*(.+)", line)
        if m:
            yield "heading", len(m.group(1)), m.group(2)
            i += 1
            continue

        # ---------- 表格 ----------
        if "|" in line:
            rows = []
            while i < n and "|" in lines[i]:
                cells = [c.strip() for c in lines[i].split("|")[1:-1]]
                if not all(re.match(r"^[-:]+$", c) for c in cells if c):
                    rows.append(cells)
                i += 1

            if rows:
                yield "table", rows
            continue

        # ---------- 空行 ----------
        if line.strip() == "":
            i += 1
            continue

        # ---------- 列表项 ----------
        m = re.match(r"^(\s*)([-*+])\s+(.*)", line)
        if m:
            indent = len(m.group(1))
            level = min(indent // 4, 2)
            yield "list_item", level, m.group(3)
            i += 1
            continue

        # ---------- 普通段落 ----------
        yield "para", line
        i += 1