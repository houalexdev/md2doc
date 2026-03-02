#!/usr/bin/env python3
# ===== numbering.py =====
"""编号列表处理模块"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def create_numbering_definition(doc, cfg):
    """
    根据 styles.json 中的 numbering_styles 配置创建编号定义。
    返回 numId 用于后续段落引用。
    """
    numbering_part = doc.part.numbering_part
    if numbering_part is None:
        from docx.oxml.numbering import CT_Numbering
        numbering_part = doc.part._numbering_part
        if numbering_part is None:
            numbering_element = CT_Numbering()
            doc.part._numbering_part = numbering_element
            numbering_part = doc.part.numbering_part
    
    numbering = numbering_part.element
    
    # 查找可用的 abstractNumId 和 numId
    abstract_num_id = 1
    num_id = 1
    
    for abstractNum in numbering.findall(qn('w:abstractNum')):
        aid = int(abstractNum.get(qn('w:abstractNumId')))
        if aid >= abstract_num_id:
            abstract_num_id = aid + 1
    
    for num in numbering.findall(qn('w:num')):
        nid = int(num.get(qn('w:numId')))
        if nid >= num_id:
            num_id = nid + 1
    
    # 从配置中读取列表样式
    numbering_styles = cfg.get("numbering_styles", [])
    if not numbering_styles:
        # 默认配置
        levels_cfg = [
            {"level": 0, "start": 1, "format": "decimal", "text": "(%1)", 
             "alignment": "left", "indent": {"left": 720, "hanging": 360}},
            {"level": 1, "start": 1, "format": "decimalEnclosedCircle", "text": "%2",
             "alignment": "left", "indent": {"left": 1080, "hanging": 360}},
            {"level": 2, "start": 1, "format": "decimal", "text": "%3)",
             "alignment": "left", "indent": {"left": 1440, "hanging": 360}},
        ]
    else:
        levels_cfg = numbering_styles[0].get("levels", [])
    
    # 创建 abstractNum
    abstractNum = OxmlElement('w:abstractNum')
    abstractNum.set(qn('w:abstractNumId'), str(abstract_num_id))
    
    multiLevelType = OxmlElement('w:multiLevelType')
    multiLevelType.set(qn('w:val'), 'multilevel')
    abstractNum.append(multiLevelType)
    
    # 定义各级
    for level_cfg in levels_cfg:
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(level_cfg["level"]))
        
        start = OxmlElement('w:start')
        start.set(qn('w:val'), str(level_cfg["start"]))
        lvl.append(start)
        
        numFmt = OxmlElement('w:numFmt')
        numFmt.set(qn('w:val'), level_cfg["format"])
        lvl.append(numFmt)
        
        lvlText = OxmlElement('w:lvlText')
        lvlText.set(qn('w:val'), level_cfg["text"])
        lvl.append(lvlText)
        
        lvlJc = OxmlElement('w:lvlJc')
        lvlJc.set(qn('w:val'), level_cfg["alignment"])
        lvl.append(lvlJc)
        
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), str(level_cfg["indent"]["left"]))
        ind.set(qn('w:hanging'), str(level_cfg["indent"]["hanging"]))
        pPr.append(ind)
        lvl.append(pPr)
        
        # 字体设置（如果有）
        if "font" in level_cfg:
            rPr = OxmlElement('w:rPr')
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), level_cfg["font"]["name"])
            rFonts.set(qn('w:eastAsia'), level_cfg["font"]["name"])
            rFonts.set(qn('w:hAnsi'), level_cfg["font"]["name"])
            rPr.append(rFonts)
            
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), str(int(level_cfg["font"]["size"] * 2)))
            rPr.append(sz)
            
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), str(int(level_cfg["font"]["size"] * 2)))
            rPr.append(szCs)
            
            lvl.append(rPr)
        
        abstractNum.append(lvl)
    
    numbering.append(abstractNum)
    
    # 创建 num 引用
    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(num_id))
    
    abstractNumId_ref = OxmlElement('w:abstractNumId')
    abstractNumId_ref.set(qn('w:val'), str(abstract_num_id))
    num.append(abstractNumId_ref)
    
    numbering.append(num)
    
    return num_id


def apply_list_numbering(paragraph, level, num_id):
    """
    将段落应用到指定的编号列表。
    level: 0, 1, 2 (对应 ilvl)
    num_id: 编号定义的 ID
    """
    pPr = paragraph._element.get_or_add_pPr()
    
    numPr = OxmlElement('w:numPr')
    
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(level))
    numPr.append(ilvl)
    
    numId = OxmlElement('w:numId')
    numId.set(qn('w:val'), str(num_id))
    numPr.append(numId)
    
    pPr.append(numPr)