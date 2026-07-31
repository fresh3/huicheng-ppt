#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""汇成医美教育 PPT 生成器 — 模板编辑 + 内容生成
优化版：唯一元素ID、输入校验、布局注册表、健壮错误处理
"""

import json, sys, os, re, shutil, zipfile, subprocess, tempfile, argparse
from lxml import etree

A  = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P  = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

# ── 品牌色 ────────────────────────────────────────────────
BRAND_GREEN  = '46A53B'
DARK_GREEN   = '213A25'
GOLD         = 'FBB03B'
WHITE        = 'FFFFFF'

def emu(inches):
    return int(inches * 914400)

# ── 布局尺寸 ─────────────────────────────────────────────
SLIDE_W  = emu(13.33)   # 宽屏
SLIDE_H  = emu(7.50)
TITLE_Y  = emu(1.5)
LINE_Y   = emu(2.5)
BODY_Y   = emu(2.8)
BODY_X   = emu(0.85)
BODY_W   = SLIDE_W - emu(0.85) * 2
BODY_H   = emu(6.2) - BODY_Y
BAR_Y    = emu(6.57)
BAR_H    = emu(0.52)

# ── 布局字段校验 ──────────────────────────────────────────
LAYOUT_FIELDS = {
    'bullets':      ('bullets',),
    'stat':         ('number',),
    'two_column':   ('left_title', 'left_items', 'right_title', 'right_items'),
    'title_content':('content',),
    'quote':        ('quote',),
    'image_text':   ('side_title', 'side_items'),
    'timeline':     ('events',),
    'metrics':      ('metrics',),
    'process':      ('steps',),
    'comparison':   ('left_header', 'right_header', 'comparison_items'),
    'intro':        (),
    'table':        ('headers', 'rows'),
    'team':         ('members',),
    'case':         ('case_name', 'case_result', 'case_detail'),
    'funnel':       ('stages',),
    'takeaway':     ('takeaways',),
}

# ── 全局元素ID计数器 ─────────────────────────────────────
_ELEM_CTR = 0
def _next_id():
    """返回全局唯一元素ID，确保同一slide中各元素不冲突"""
    global _ELEM_CTR
    _ELEM_CTR += 1
    return str(100 + _ELEM_CTR)

# ── XML 工具 ─────────────────────────────────────────────
def find(el, xpath):
    return el.find(xpath, namespaces={'a': A, 'p': P, 'r': R})

def save(rt, path):
    rt.getroottree().write(path, xml_declaration=True, encoding='UTF-8', standalone=True)



# ── 字号提取 ─────────────────────────────────────────────
def _get_sz(slide_xml):
    """从 slide XML 中找到第一个带字号的 run，返回 sz 值"""
    for rs in slide_xml.xpath('//a:r/a:rPr[@sz]', namespaces={'a': A}):
        try:
            return int(rs.get('sz'))
        except (ValueError, TypeError):
            pass
    return None

# ── 文本框 ───────────────────────────────────────────────
def make_textbox(x, y, cx, cy, name, paragraphs, font_name=None, font_sz=None):
    """
    paragraphs: [{ text, bold, color, sz, bullet, font }]
    font_name:  整体字体（微软雅黑/Arial），None 则用模板默认
    font_sz:    整体字号（百分之一pt），None 则用模板默认
    """
    sp = etree.Element(f'{{{P}}}sp')

    # nvSpPr
    nv = etree.SubElement(sp, f'{{{P}}}nvSpPr')
    cNvPr = etree.SubElement(nv, f'{{{P}}}cNvPr')
    cNvPr.set('id', _next_id())
    cNvPr.set('name', name)
    etree.SubElement(nv, f'{{{P}}}cNvSpPr').set('txBox', '1')
    etree.SubElement(nv, f'{{{P}}}nvPr')

    # spPr
    spPr = etree.SubElement(sp, f'{{{P}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A}}}xfrm')
    off  = etree.SubElement(xfrm, f'{{{A}}}off'); off.set('x', str(x)); off.set('y', str(y))
    ext  = etree.SubElement(xfrm, f'{{{A}}}ext'); ext.set('cx', str(cx)); ext.set('cy', str(cy))
    prstGeom = etree.SubElement(spPr, f'{{{A}}}prstGeom'); prstGeom.set('prst', 'rect')
    etree.SubElement(prstGeom, f'{{{A}}}avLst')
    etree.SubElement(spPr, f'{{{A}}}noFill')
    ln = etree.SubElement(spPr, f'{{{A}}}ln'); ln.set('w', '9525')
    etree.SubElement(ln, f'{{{A}}}noFill')

    # txBody
    txBody = etree.SubElement(sp, f'{{{P}}}txBody')
    bodyPr = etree.SubElement(txBody, f'{{{A}}}bodyPr')
    bodyPr.set('wrap', 'square'); bodyPr.set('rtlCol', '0'); bodyPr.set('anchor', 't')
    etree.SubElement(txBody, f'{{{A}}}lstStyle')

    for para in paragraphs:
        p = etree.SubElement(txBody, f'{{{A}}}p')
        pPr = etree.SubElement(p, f'{{{A}}}pPr'); pPr.set('lvl', '0')
        if para.get('bullet'):
            buNone = etree.SubElement(pPr, f'{{{A}}}buNone')
            pPr.remove(buNone)
            buClr  = etree.SubElement(pPr, f'{{{A}}}buClr')
            buClrS = etree.SubElement(buClr, f'{{{A}}}srgbClr'); buClrS.set('val', GOLD)
            buChar = etree.SubElement(pPr, f'{{{A}}}buChar'); buChar.set('char', '●')
            buSzPct = etree.SubElement(pPr, f'{{{A}}}buSzPct'); buSzPct.set('val', '80000')
            buSpc   = etree.SubElement(pPr, f'{{{A}}}buSpcPct'); buSpc.set('val', '60000')

        r   = etree.SubElement(p, f'{{{A}}}r')
        rPr = etree.SubElement(r, f'{{{A}}}rPr')
        rPr.set('lang', 'zh-CN'); rPr.set('dirty', '0')
        if para.get('bold'):
            rPr.set('b', '1')
        if para.get('color'):
            sf = etree.SubElement(rPr, f'{{{A}}}solidFill')
            cl = etree.SubElement(sf, f'{{{A}}}srgbClr'); cl.set('val', para['color'])
        if font_sz or para.get('sz'):
            rPr.set('sz', str(font_sz or para['sz']))
        if font_name or para.get('font'):
            ea = etree.SubElement(rPr, f'{{{A}}}ea'); ea.set('typeface', font_name or para['font'])
        t = etree.SubElement(r, f'{{{A}}}t'); t.text = para.get('text', '')

    return sp

# ── 线条 ─────────────────────────────────────────────────
def make_line(x, y, cx, color=GOLD, w=19050):
    cxn = etree.Element(f'{{{P}}}cxnSp')
    nv = etree.SubElement(cxn, f'{{{P}}}nvCxnSpPr')
    c = etree.SubElement(nv, f'{{{P}}}cNvPr'); c.set('id', _next_id()); c.set('name', 'Line')
    etree.SubElement(nv, f'{{{P}}}cNvCxnSpPr'); etree.SubElement(nv, f'{{{P}}}nvPr')

    spPr = etree.SubElement(cxn, f'{{{P}}}spPr')
    xf = etree.SubElement(spPr, f'{{{A}}}xfrm')
    o = etree.SubElement(xf, f'{{{A}}}off'); o.set('x', str(x)); o.set('y', str(y))
    e = etree.SubElement(xf, f'{{{A}}}ext'); e.set('cx', str(cx)); e.set('cy', '10160')
    pg = etree.SubElement(spPr, f'{{{A}}}prstGeom'); pg.set('prst', 'line')
    etree.SubElement(pg, f'{{{A}}}avLst')

    ln = etree.SubElement(spPr, f'{{{A}}}ln'); ln.set('w', str(w))
    sf = etree.SubElement(ln, f'{{{A}}}solidFill')
    cl = etree.SubElement(sf, f'{{{A}}}srgbClr'); cl.set('val', color)

    style = etree.SubElement(cxn, f'{{{P}}}style')
    lnRef   = etree.SubElement(style, f'{{{A}}}lnRef');   lnRef.set('idx', '2')
    etree.SubElement(lnRef, f'{{{A}}}schemeClr').set('val', 'accent1')
    fillRef = etree.SubElement(style, f'{{{A}}}fillRef'); fillRef.set('idx', '0')
    etree.SubElement(fillRef, f'{{{A}}}srgbClr').set('val', WHITE)
    effRef  = etree.SubElement(style, f'{{{A}}}effectRef'); effRef.set('idx', '0')
    etree.SubElement(effRef, f'{{{A}}}srgbClr').set('val', WHITE)
    fontRef = etree.SubElement(style, f'{{{A}}}fontRef'); fontRef.set('idx', 'minor')
    etree.SubElement(fontRef, f'{{{A}}}schemeClr').set('val', 'tx1')
    return cxn

# ── 图片 ─────────────────────────────────────────────────
def make_pic(rid, x, y, cx, cy, name):
    pic = etree.Element(f'{{{P}}}pic')

    nvPicPr = etree.SubElement(pic, f'{{{P}}}nvPicPr')
    cNvPr   = etree.SubElement(nvPicPr, f'{{{P}}}cNvPr')
    cNvPr.set('id', _next_id()); cNvPr.set('name', name)
    cNvPicPr = etree.SubElement(nvPicPr, f'{{{P}}}cNvPicPr')
    etree.SubElement(cNvPicPr, f'{{{A}}}picLocks').set('noChangeAspect', '1')
    etree.SubElement(nvPicPr, f'{{{P}}}nvPr')

    blipFill = etree.SubElement(pic, f'{{{P}}}blipFill')
    blip = etree.SubElement(blipFill, f'{{{A}}}blip'); blip.set(f'{{{R}}}embed', rid)
    stretch = etree.SubElement(blipFill, f'{{{A}}}stretch')
    etree.SubElement(stretch, f'{{{A}}}fillRect')

    spPr = etree.SubElement(pic, f'{{{P}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A}}}xfrm')
    off  = etree.SubElement(xfrm, f'{{{A}}}off'); off.set('x', str(x)); off.set('y', str(y))
    ext  = etree.SubElement(xfrm, f'{{{A}}}ext'); ext.set('cx', str(cx)); ext.set('cy', str(cy))
    prstGeom = etree.SubElement(spPr, f'{{{A}}}prstGeom'); prstGeom.set('prst', 'rect')
    etree.SubElement(prstGeom, f'{{{A}}}avLst')

    return pic

# ── 背景图 rId 提取 ──────────────────────────────────────
def _get_bg_rid(slide_xml):
    """找到背景图 pic 的 rId（模板 slide4 里第一张图）"""
    pics = slide_xml.xpath('//p:pic', namespaces={'p': P, 'a': A, 'r': R})
    for pic in pics:
        blip = find(pic, './/a:blip')
        if blip is not None:
            return blip.get(f'{{{R}}}embed')
    return None

# ── 构建内容页 ───────────────────────────────────────────
def create_content(path, rels_path, data, media_dir):
    """
    path       — 新克隆的 slideN.xml 路径
    rels_path  — 对应的 slideN.xml.rels 路径
    data       — 单个 content 数据（含 title/layout/各种字段）
    media_dir  — skill 的 media/ 目录
    """
    slide_xml = etree.parse(path).getroot()

    # 1) 找到并移除所有现有元素（pic/sp/cxnSp/group）
    bg_rid = _get_bg_rid(slide_xml)
    spTree = find(slide_xml, './/p:cSld/p:spTree')
    keep = []
    for child in list(spTree):
        tag = etree.QName(child.tag).localname
        if tag in ('grpSpPr', 'nvGrpSpPr'):
            keep.append(child)
    for child in list(spTree):
        if child not in keep:
            spTree.remove(child)

    # 2) 重新添加全屏背景图
    bg_pic = make_pic(bg_rid, 0, 0, SLIDE_W, SLIDE_H, 'BG')
    spTree.append(bg_pic)

    # 3) 底部品牌栏
    bar_pic = make_pic('rId5', 0, BAR_Y, SLIDE_W, BAR_H, 'BottomBar')
    spTree.append(bar_pic)

    # 4) 标题
    paras = [{'text': data['title'], 'bold': True, 'color': WHITE, 'sz': 3600}]
    title_box = make_textbox(BODY_X, TITLE_Y, BODY_W, emu(0.9), 'Title', paras)
    spTree.append(title_box)

    # 5) 金色分割线
    gold_line = make_line(BODY_X, LINE_Y, emu(2.5))
    spTree.append(gold_line)

    # 6) 加载 rels 并补充缺失图片（底部栏、logo 等）
    rels_tree = etree.parse(rels_path)
    rr = rels_tree.getroot()
    existing_rids = {r.get('Id') for r in rr.findall('Relationship')}

    # 底部品牌栏
    bar_rid = 'rId5'
    if bar_rid not in existing_rids:
        nr = etree.SubElement(rr, 'Relationship')
        nr.set('Id', bar_rid)
        nr.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        nr.set('Target', f'../media/image11.png')
        nr.set('TargetMode', 'Internal')
        existing_rids.add(bar_rid)

    # 布局内容元素
    layout = data.get('layout', 'bullets')
    y_pos = int(BODY_Y)
    line_spacing = emu(0.4)
    item_sz = 1800
    sub_sz  = 1400
    unit_sz = 2400

    # 布局内辅助函数
    def textbox(x, y, cx, cy, name, paragraphs, font_name=None, font_sz=None):
        sp = make_textbox(x, y, cx, cy, name, paragraphs, font_name, font_sz)
        spTree.append(sp)
        return y + cy

    def line(x, y, cx, color=GOLD, w=19050):
        ln = make_line(x, y, cx, color, w)
        spTree.append(ln)
        return y + emu(0.15)

    def pic(rid, x, y, cx, cy, name):
        pc = make_pic(rid, x, y, cx, cy, name)
        spTree.append(pc)
        return y + cy

    def build_bullets(items):
        paras = [{'text': t, 'bold': False, 'color': WHITE, 'sz': item_sz, 'bullet': True} for t in items]
        return textbox(BODY_X, y_pos, BODY_W, BODY_H, 'Body', paras)

    # ── 布局注册表 ─────────────────────────────────────
    # 每个布局接收 (data, y_pos) 返回下一个 y 位置
    # 辅助函数通过闭包访问 spTree / BODY_X / BODY_W / BODY_H 等

    LEFT_X  = BODY_X
    RIGHT_X = BODY_X + emu(6.0)
    HALF_W  = emu(5.5)

    def _layout_bullets(data, y):
        return build_bullets(data.get('bullets', ['（暂无内容）']))

    def _layout_title_content(data, y):
        paras = [{'text': data.get('content', ''), 'bold': False, 'color': WHITE, 'sz': item_sz}]
        return textbox(BODY_X, y, BODY_W, BODY_H, 'Content', paras)

    def _layout_stat(data, y):
        number = data.get('number', '0')
        unit   = data.get('unit', '')
        desc   = data.get('desc', '')
        ny = textbox(BODY_X, y, emu(5.0), emu(1.8), 'Number',
                     [{'text': number, 'bold': True, 'color': GOLD, 'sz': 7200, 'font': 'Arial'}])
        uy = textbox(BODY_X, ny, emu(3.0), emu(0.6), 'Unit',
                     [{'text': unit, 'bold': False, 'color': WHITE, 'sz': unit_sz}]) if unit else ny
        if desc:
            textbox(BODY_X, uy + emu(0.1), BODY_W, emu(0.6), 'Desc',
                    [{'text': desc, 'bold': False, 'color': WHITE, 'sz': sub_sz}])
        return uy + emu(0.6) if desc else uy

    def _layout_two_column(data, y):
        lt = data.get('left_title', '')
        rt = data.get('right_title', '')
        textbox(LEFT_X, y, HALF_W, emu(0.5), 'LTitle',
                [{'text': lt, 'bold': True, 'color': GOLD, 'sz': 2200}])
        textbox(RIGHT_X, y, HALF_W, emu(0.5), 'RTitle',
                [{'text': rt, 'bold': True, 'color': GOLD, 'sz': 2200}])
        ly = textbox(LEFT_X, y + emu(0.6), HALF_W, emu(2.8), 'LBody',
                     [{'text': i, 'bold': False, 'color': WHITE, 'sz': item_sz, 'bullet': True}
                      for i in data.get('left_items', [])])
        textbox(RIGHT_X, y + emu(0.6), HALF_W, emu(2.8), 'RBody',
                [{'text': i, 'bold': False, 'color': WHITE, 'sz': item_sz, 'bullet': True}
                 for i in data.get('right_items', [])])
        return max(ly, y + emu(3.4))

    def _layout_quote(data, y):
        q = data.get('quote', '')
        src = data.get('source', '')
        qy = textbox(BODY_X, y, BODY_W, emu(2.5), 'Quote',
                     [{'text': f'"{q}"', 'bold': False, 'color': WHITE, 'sz': 2800}])
        if src:
            textbox(BODY_X, qy + emu(0.1), BODY_W, emu(0.5), 'Source',
                    [{'text': f'— {src}', 'bold': False, 'color': GOLD, 'sz': 1400}])
        return qy + emu(0.6) if src else qy

    def _layout_image_text(data, y):
        # 左侧图片占位（实际图片需用户提供，这里用文本代替）
        textbox(LEFT_X, y, HALF_W, emu(2.8), 'ImagePlaceholder',
                [{'text': '[图片区域]', 'bold': False, 'color': WHITE, 'sz': 1800}])
        st = data.get('side_title', '')
        if st:
            textbox(RIGHT_X, y, HALF_W, emu(0.5), 'SideTitle',
                    [{'text': st, 'bold': True, 'color': GOLD, 'sz': 2200}])
        textbox(RIGHT_X, y + emu(0.6), HALF_W, emu(2.4), 'SideBody',
                [{'text': i, 'bold': False, 'color': WHITE, 'sz': item_sz, 'bullet': True}
                 for i in data.get('side_items', [])])
        return y + emu(3.0)

    def _layout_timeline(data, y):
        events = data.get('events', [])
        for ev in events:
            yr = ev.get('year', '')
            desc = ev.get('event', '')
            textbox(LEFT_X, y, emu(1.5), emu(0.5), 'Year',
                    [{'text': yr, 'bold': True, 'color': GOLD, 'sz': 2800, 'font': 'Arial'}])
            textbox(LEFT_X + emu(1.8), y, BODY_W - emu(1.8), emu(0.8), 'Event',
                    [{'text': desc, 'bold': False, 'color': WHITE, 'sz': item_sz}])
            y += emu(0.9)
        return y

    def _layout_metrics(data, y):
        metrics = data.get('metrics', [])
        card_w = emu(2.8)
        gap = emu(0.3)
        for i, m in enumerate(metrics):
            cx = LEFT_X + i * (card_w + gap)
            label = m.get('label', '')
            value = m.get('value', '')
            unit  = m.get('unit', '')
            sub   = m.get('sub', '')
            textbox(cx, y, card_w, emu(0.4), f'MLabel{i}',
                    [{'text': label, 'bold': False, 'color': GOLD, 'sz': 1600}])
            textbox(cx, y + emu(0.45), card_w, emu(1.0), f'MVal{i}',
                    [{'text': value, 'bold': True, 'color': WHITE, 'sz': 4800, 'font': 'Arial'}])
            if unit:
                textbox(cx, y + emu(1.5), card_w, emu(0.4), f'MUnit{i}',
                        [{'text': unit, 'bold': False, 'color': GOLD, 'sz': unit_sz}])
            if sub:
                textbox(cx, y + emu(1.95), card_w, emu(0.35), f'MSub{i}',
                        [{'text': sub, 'bold': False, 'color': WHITE, 'sz': 1200}])
        return y + emu(2.4)

    def _layout_process(data, y):
        steps = data.get('steps', [])
        n = len(steps)
        if n == 0:
            return y
        gap = emu(0.3)
        col_w = (BODY_W - (n - 1) * gap) // max(n, 1)
        for i, step in enumerate(steps):
            cx = LEFT_X + i * (col_w + gap)
            textbox(cx, y, col_w, emu(0.7), f'StepNum{i}',
                    [{'text': str(i + 1), 'bold': True, 'color': GOLD, 'sz': 3600, 'font': 'Arial'}])
            st = step.get('title', f'步骤{i+1}')
            textbox(cx, y + emu(0.8), col_w, emu(0.5), f'StepTitle{i}',
                    [{'text': st, 'bold': True, 'color': WHITE, 'sz': 1800}])
            desc = step.get('desc', '')
            if desc:
                textbox(cx, y + emu(1.4), col_w, emu(1.5), f'StepDesc{i}',
                        [{'text': desc, 'bold': False, 'color': WHITE, 'sz': sub_sz}])
        return y + emu(3.0)

    def _layout_comparison(data, y):
        lh = data.get('left_header', '')
        rh = data.get('right_header', '')
        items = data.get('comparison_items', [])
        textbox(LEFT_X, y, HALF_W, emu(0.5), 'CmpLH',
                [{'text': lh, 'bold': True, 'color': GOLD, 'sz': 2200}])
        textbox(RIGHT_X, y, HALF_W, emu(0.5), 'CmpRH',
                [{'text': rh, 'bold': True, 'color': GOLD, 'sz': 2200}])
        cy = y + emu(0.6)
        for item in items:
            textbox(LEFT_X, cy, HALF_W, emu(0.7), 'CmpL',
                    [{'text': item.get('left', ''), 'bold': False, 'color': WHITE, 'sz': 1500}])
            textbox(RIGHT_X, cy, HALF_W, emu(0.7), 'CmpR',
                    [{'text': item.get('right', ''), 'bold': False, 'color': WHITE, 'sz': 1500}])
            cy += emu(0.8)
        return cy

    def _layout_intro(data, y):
        kicker = data.get('kicker', '')
        lead   = data.get('lead', '')
        if kicker:
            textbox(BODY_X, y, BODY_W, emu(0.5), 'Kicker',
                    [{'text': kicker, 'bold': True, 'color': BRAND_GREEN, 'sz': 1600}])
            y += emu(0.6)
        if lead:
            textbox(BODY_X, y, BODY_W, emu(2.0), 'Lead',
                    [{'text': lead, 'bold': False, 'color': WHITE, 'sz': 1800}])
            y += emu(2.2)
        return y

    def _layout_table(data, y):
        headers = data.get('headers', [])
        rows    = data.get('rows', [])
        # 简单表格用文本模拟，列宽均分
        ncol = max(len(headers), 1)
        col_w = BODY_W // ncol
        # 表头
        for j, h in enumerate(headers):
            textbox(BODY_X + j * col_w, y, col_w, emu(0.45), f'TH{j}',
                    [{'text': h, 'bold': True, 'color': GOLD, 'sz': 1600}])
        ry = y + emu(0.55)
        for i, row in enumerate(rows):
            for j, cell in enumerate(row):
                textbox(BODY_X + j * col_w, ry, col_w, emu(0.4), f'TR{i}C{j}',
                        [{'text': str(cell), 'bold': False, 'color': WHITE, 'sz': 1400}])
            ry += emu(0.45)
        return ry

    def _layout_team(data, y):
        members = data.get('members', [])
        n = len(members)
        if n == 0:
            return y
        card_w = emu(3.5)
        gap = emu(0.4)
        total_w = n * card_w + (n - 1) * gap
        start_x = (SLIDE_W - total_w) // 2
        for i, m in enumerate(members):
            cx = start_x + i * (card_w + gap)
            name = m.get('name', '')
            role = m.get('role', '')
            desc = m.get('desc', '')
            textbox(cx, y, card_w, emu(0.5), f'TName{i}',
                    [{'text': name, 'bold': True, 'color': WHITE, 'sz': 2200}])
            textbox(cx, y + emu(0.55), card_w, emu(0.35), f'TRole{i}',
                    [{'text': role, 'bold': True, 'color': GOLD, 'sz': 1400}])
            if desc:
                textbox(cx, y + emu(0.95), card_w, emu(1.5), f'TDesc{i}',
                        [{'text': desc, 'bold': False, 'color': WHITE, 'sz': 1200}])
        return y + emu(2.5)

    def _layout_case(data, y):
        name   = data.get('case_name', '')
        result = data.get('case_result', '')
        detail = data.get('case_detail', '')
        textbox(BODY_X, y, BODY_W, emu(0.6), 'CaseName',
                [{'text': name, 'bold': True, 'color': GOLD, 'sz': 2800}])
        ry = y + emu(0.7)
        if result:
            textbox(BODY_X, ry, BODY_W, emu(0.5), 'CaseResult',
                    [{'text': result, 'bold': True, 'color': WHITE, 'sz': 2000}])
            ry += emu(0.6)
        if detail:
            textbox(BODY_X, ry, BODY_W, emu(2.0), 'CaseDetail',
                    [{'text': detail, 'bold': False, 'color': WHITE, 'sz': item_sz}])
            ry += emu(2.2)
        return ry

    def _layout_funnel(data, y):
        stages = data.get('stages', [])
        if not stages:
            return y
        n = len(stages)
        stage_h = emu(0.7)
        gap = emu(0.1)
        max_w = BODY_W
        for i, stage in enumerate(stages):
            w = max_w * (n - i) // n
            cx = (SLIDE_W - w) // 2
            lbl = stage.get('label', '')
            val = stage.get('value', '')
            textbox(cx, y, w, stage_h, f'Funnel{i}',
                    [{'text': f'{lbl}  {val}', 'bold': True, 'color': WHITE, 'sz': 1600}])
            y += stage_h + gap
        return y

    def _layout_takeaway(data, y):
        takeaways = data.get('takeaways', ['（暂无内容）'])
        paras = [{'text': t, 'bold': False, 'color': WHITE, 'sz': 2000, 'bullet': True} for t in takeaways]
        return textbox(BODY_X, y, BODY_W, BODY_H, 'Takeaway', paras)

    # ── 布局分发 ────────────────────────────────────────
    LAYOUTS = {
        'bullets':       _layout_bullets,
        'title_content': _layout_title_content,
        'stat':          _layout_stat,
        'two_column':    _layout_two_column,
        'quote':         _layout_quote,
        'image_text':    _layout_image_text,
        'timeline':      _layout_timeline,
        'metrics':       _layout_metrics,
        'process':       _layout_process,
        'comparison':    _layout_comparison,
        'intro':         _layout_intro,
        'table':         _layout_table,
        'team':          _layout_team,
        'case':          _layout_case,
        'funnel':        _layout_funnel,
        'takeaway':      _layout_takeaway,
    }

    layout_fn = LAYOUTS.get(layout)
    if layout_fn:
        layout_fn(data, y_pos)
    else:
        print(f'⚠ 未知布局 "{layout}"，使用默认 bullets 布局', file=sys.stderr)
        _layout_bullets(data, y_pos)

    # 保存 rels（所有 rId 添加完毕后）
    save(rels_tree.getroot(), rels_path)
    save(slide_xml, path)

# ── 克隆幻灯片 ───────────────────────────────────────────
def dup_slide(unpacked, src, add_slide_py):
    """复制 slide，返回新 slide 文件名（如 slide6.xml）"""
    res = subprocess.run(
        ['python3', add_slide_py, unpacked, src],
        capture_output=True, text=True, encoding='utf-8'
    )
    if res.returncode != 0:
        raise RuntimeError(f'add_slide.py 失败: {res.stderr}')
    m = re.search(r'(ppt/slides/slide\d+\.xml)', res.stdout)
    if not m:
        raise RuntimeError(f'无法解析 add_slide.py 输出:\n{res.stdout}')
    return os.path.basename(m.group(1))

# ── 编辑封面 ─────────────────────────────────────────────
def edit_cover(path, plan):
    slide = etree.parse(path).getroot()
    for p in slide.xpath('//a:p', namespaces={'a': A}):
        runs = p.findall('a:r', namespaces={'a': A})
        text = ''.join(r.find('a:t', namespaces={'a': A}).text or '' for r in runs)
        if 'PPT' in text:
            replace_para_text(p, plan.get('title', ''))
        elif 'HCYM EDUCATION' in text:
            replace_para_text(p, plan.get('subtitle', ''))
        elif 'REPORT' in text.upper() or 'ANNUAL' in text.upper():
            replace_para_text(p, plan.get('subtitle2', ''))
        elif '主讲人' in text:
            replace_para_text(p, plan.get('presenter', ''))
    save(slide, path)

def replace_para_text(para, new_text):
    """替换段落文本，保留首 run 的格式"""
    runs = para.findall('a:r', namespaces={'a': A})
    if not runs:
        return
    rPr = find(runs[0], 'a:rPr')
    rPr_copy = None
    if rPr is not None:
        rPr_copy = etree.fromstring(etree.tostring(rPr))
    for r in runs:
        para.remove(r)
    r   = etree.SubElement(para, f'{{{A}}}r')
    if rPr_copy is not None:
        r.insert(0, rPr_copy)
    t   = etree.SubElement(r, f'{{{A}}}t')
    t.text = new_text

# ── 编辑目录 ─────────────────────────────────────────────
def edit_toc(path, items):
    slide = etree.parse(path).getroot()
    for i, item_text in enumerate(items[:5], 1):
        for p in slide.xpath('//a:p', namespaces={'a': A}):
            runs = p.findall('a:r', namespaces={'a': A})
            text = ''.join(r.find('a:t', namespaces={'a': A}).text or '' for r in runs)
            if text == f'目录标题{i}':
                replace_para_text(p, item_text)
                break
    save(slide, path)

# ── 编辑章节页 ───────────────────────────────────────────
def edit_section(path, data):
    slide = etree.parse(path).getroot()
    for p in slide.xpath('//a:p', namespaces={'a': A}):
        runs = p.findall('a:r', namespaces={'a': A})
        text = ''.join(r.find('a:t', namespaces={'a': A}).text or '' for r in runs)
        if 'PART' in text or 'Part' in text:
            part = data.get('part', '01')
            replace_para_text(p, f'PART {part}')
        elif text.strip().startswith('20'):  # 年份
            replace_para_text(p, data.get('year', ''))
        elif len(text.strip()) >= 2 and 'PART' not in text and not text.strip().startswith('20'):
            replace_para_text(p, data.get('title', ''))
    save(slide, path)

# ── 主生成流程 ───────────────────────────────────────────
def generate(plan, tpl_path, pptx_scripts, out_path, skill_root):
    tpl_path    = os.path.abspath(tpl_path)
    pptx_scripts = os.path.abspath(pptx_scripts)
    out_path    = os.path.abspath(out_path)
    skill_root  = os.path.abspath(skill_root)

    add_slide_py = os.path.join(pptx_scripts, 'add_slide.py')
    clean_py     = os.path.join(pptx_scripts, 'clean.py')

    # 校验依赖
    for f, desc in [(add_slide_py, 'add_slide.py'), (clean_py, 'clean.py'),
                    (tpl_path, 'template.pptx')]:
        if not os.path.exists(f):
            print(f'✗ 找不到 {desc}: {f}', file=sys.stderr); sys.exit(1)

    work = tempfile.mkdtemp(prefix='hc_ppt_')
    print(f'[1/6] 解压模板...')
    with zipfile.ZipFile(tpl_path, 'r') as zf:
        zf.extractall(work)

    slides_dir = os.path.join(work, 'ppt', 'slides')
    rels_dir   = os.path.join(work, 'ppt', 'slides', '_rels')
    media_dir  = os.path.join(skill_root, 'media')

    # 拷贝品牌素材
    if os.path.isdir(media_dir):
        dst_media = os.path.join(work, 'ppt', 'media')
        os.makedirs(dst_media, exist_ok=True)
        for f in os.listdir(media_dir):
            shutil.copy2(os.path.join(media_dir, f), os.path.join(dst_media, f))

    # ── 按 slide 分配 ──────────────────────────────────
    cover_slides   = [s for s in plan['slides'] if s['type'] == 'cover']
    toc_slides     = [s for s in plan['slides'] if s['type'] == 'toc']
    section_slides = [s for s in plan['slides'] if s['type'] == 'section']
    content_slides = [s for s in plan['slides'] if s['type'] == 'content']
    end_slides     = [s for s in plan['slides'] if s['type'] == 'end']

    print(f'[2/6] 克隆幻灯片...')
    slide_map = {}
    for i, s in enumerate(cover_slides):
        slide_map[f'cover_{i}'] = f'slide{i+1}.xml'

    # 克隆内容页（从 slide4.xml）
    content_files = []
    for i, _ in enumerate(content_slides):
        new = dup_slide(work, 'slide4.xml', add_slide_py)
        content_files.append(new)
        slide_map[f'content_{i}'] = new

    # 结束页
    end_files = []
    for i, _ in enumerate(end_slides):
        new = dup_slide(work, 'slide5.xml', add_slide_py)
        end_files.append(new)
        slide_map[f'end_{i}'] = new

    print(f'[3/6] 编辑页面内容...')

    # 封面
    for i, s in enumerate(cover_slides):
        edit_cover(os.path.join(slides_dir, slide_map[f'cover_{i}']), s)

    # 目录
    for i, s in enumerate(toc_slides):
        edit_toc(os.path.join(slides_dir, f'slide{i+2}.xml'), s.get('items', []))

    # 章节
    sec_idx = 0
    for s in section_slides:
        # 章节页用 slide3 的克隆
        new = dup_slide(work, 'slide3.xml', add_slide_py)
        edit_section(os.path.join(slides_dir, new), s)
        slide_map[f'section_{sec_idx}'] = new
        sec_idx += 1

    # 内容页
    for i, s in enumerate(content_slides):
        sfile = content_files[i]
        rfile = sfile.replace('.xml', '.xml.rels')
        create_content(
            os.path.join(slides_dir, sfile),
            os.path.join(rels_dir, rfile),
            s, media_dir
        )

    print(f'[4/6] 重排序幻灯片...')
    pres_path = os.path.join(work, 'ppt', 'presentation.xml')
    pres = etree.parse(pres_path).getroot()
    sldIdLst = find(pres, './/p:sldIdLst')

    # 构建目标顺序
    ordered = []
    for s in plan['slides']:
        t = s['type']
        if t == 'cover':
            ordered.append(slide_map['cover_0'])
        elif t == 'toc':
            ordered.append('slide2.xml')
        elif t == 'section':
            pass  # 已单独处理
        elif t == 'content':
            pass  # 已单独处理
        elif t == 'end':
            pass  # 已单独处理

    # 按 plan 顺序重新构建
    full_order = []
    sec_i = cnt_i = end_i = 0
    for s in plan['slides']:
        t = s['type']
        if t == 'cover':   full_order.append(slide_map.get('cover_0', 'slide1.xml'))
        elif t == 'toc':   full_order.append('slide2.xml')
        elif t == 'section':
            full_order.append(slide_map.get(f'section_{sec_i}', 'slide3.xml')); sec_i += 1
        elif t == 'content':
            full_order.append(slide_map.get(f'content_{cnt_i}')); cnt_i += 1
        elif t == 'end':
            full_order.append(slide_map.get(f'end_{end_i}', 'slide5.xml')); end_i += 1

    # 重建 sldIdLst
    old_ids = list(sldIdLst)
    id_map = {os.path.basename(e.get(f'{{{R}}}id')): e for e in old_ids}
    for e in old_ids:
        sldIdLst.remove(e)

    base_id = 256
    for fname in full_order:
        if fname and fname in id_map:
            el = id_map[fname]
            el.set('id', str(base_id))
            sldIdLst.append(el)
            base_id += 1

    save(pres, pres_path)

    print(f'[5/6] 清理无用文件...')
    subprocess.run(['python3', clean_py, work], capture_output=True)

    print(f'[6/6] 打包 PPTX...')
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(work):
            for f in files:
                fp = os.path.join(root, f)
                zf.write(fp, os.path.relpath(fp, work))

    shutil.rmtree(work, ignore_errors=True)
    print(f'✓ 完成！→ {out_path}')
    print(f'  共 {len(full_order)} 页')

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='汇成 PPT 生成器')
    parser.add_argument('plan', help='JSON 计划文件路径')
    parser.add_argument('--pptx-scripts', required=True, help='pptx skill 脚本目录')
    parser.add_argument('--output', '-o', default='huicheng_output.pptx')
    parser.add_argument('--skill-root', required=True, help='skill 根目录（含 templates/ 和 media/）')
    args = parser.parse_args()

    with open(args.plan, encoding='utf-8') as f:
        plan = json.load(f)

    tpl = os.path.join(args.skill_root, 'templates', 'template.pptx')
    generate(plan, tpl, args.pptx_scripts, args.output, args.skill_root)
