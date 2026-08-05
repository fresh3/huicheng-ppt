#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""汇成医美教育 PPT 生成器 — 模板编辑 + 内容生成
优化版：唯一元素ID、输入校验、布局注册表、健壮错误处理
"""

import json, sys, os, re, shutil, zipfile, subprocess, tempfile, argparse, math
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
BAR_Y    = 6007100   # 模板 slide4 精确值
BAR_H    = 698500   # 模板 slide4 精确值

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
    'intro':        ('kicker', 'lead'),
    'table':        ('headers', 'rows'),
    'team':         ('members',),
    'case':         ('case_name', 'case_result', 'case_detail'),
    'funnel':       ('stages',),
    'takeaway':     ('takeaways',),
    # ── dashiai-ppt 风格新增布局 ──
    'comparison_cards': ('left_title', 'left_items', 'right_title', 'right_items'),
    'process_circles':  ('steps',),
    'tag_list':         ('tags',),
    'highlight_box':    ('big_number', 'big_label'),
    'dual_stat':        ('left_value', 'right_value'),
    'card_grid':        ('cards',),
    'before_after':     ('before_title', 'before_items', 'after_title', 'after_items'),
    # ── 10 种 dashi-ppt 语义布局 ──
    'swot':             ('quadrants',),
    'quadrant':         ('items',),
    'checklist':        ('items',),
    'scorecard':        ('items',),
    'stair':            ('steps',),
    'flywheel':         ('items',),
    'statement':        ('text',),
    'journey':          ('stages',),
    'pricing':          ('plans',),
    'faq':              ('items',),
    # ── 本轮新增 9 种 ──
    'pie_chart':        ('slices',),
    'bar_chart':        ('bars',),
    'dashboard':        ('cards',),
    'hero_banner':      ('subtitle',),
    'numbered_list':    ('items',),
    'matrix_2x2':       ('quadrants',),
    'chart_placeholder':('chart_type',),
    'kpi_card':         ('cards',),
    'feature_grid':     ('items',),
    # ── 本轮补齐 15 种 ──
    'radar_chart':      ('axes',),
    'pyramid':          ('levels',),
    'roadmap':          ('phases',),
    'venn':             ('groups',),
    'ranking':          ('items',),
    'waterfall':        ('values',),
    'heatmap':          ('rows',),
    'gantt':            ('tasks',),
    'cycle':            ('steps',),
    'big_number':       ('number',),
    'gallery':          ('items',),
    'layers':           ('layers',),
    'bento':            ('cells',),
    'gauge':            ('value',),
    'testimonial':      ('quotes',),
    # ── 再次补齐 10 种 ──
    'treemap':          ('items',),
    'scatter':          ('points',),
    'stacked_bar':      ('categories',),
    'profile':          ('name',),
    'spotlight':        ('big_stat',),
    'risk':             ('risks',),
    'swimlane':         ('lanes',),
    'overview':         ('summary', 'key_points'),
    'principles':       ('items',),
    'org_chart':        ('nodes',),
    # ── 补齐 dashi-ppt 图表 11 种 ──
    'bump':             ('series',),
    'dumbbell':         ('items',),
    'lollipop':         ('items',),
    'waffle':           ('items',),
    'radial_bar':       ('items',),
    'diverging':        ('items',),
    'tornado':          ('items',),
    'honeycomb':        ('items',),
    'slope':            ('items',),
    'pictogram':        ('items',),
    'sunburst':         ('inner',),
    # ── 补齐 dashi-ppt 剩余 15 种 ──
    'mekko':            ('items',),
    'grouped':          ('categories',),
    'trend':            ('series',),
    'chain':            ('stages',),
    'calendar':         ('weeks',),
    'orbit':            ('center',),
    'triptych':         ('panels',),
    'meter':            ('items',),
    'pareto':           ('items',),
    'delta':            ('items',),
    'milestones':       ('milestones',),
    'spectrum':         ('items',),
    'logowall':         ('items',),
    'masonry':          ('items',),
    'ladder':           ('stages',),
    'mindmap':          ('center', 'branches'),
    'network':          ('nodes', 'edges', 'center'),
    'mosaic':           ('items',),
    'sticker_bubble':   ('big_number', 'big_label', 'satellites'),
    'bubbletl':         ('items',),
    'icicle':           ('items',),
    'candles':          ('items',),
    'hypecycle':        ('items', 'phases'),
    'typeriver':        ('words', 'lead'),
    'ribbon':           ('items',),
    'vinyl':            ('title', 'tracks'),
    'polar_rose':       ('items',),
    'histogram':        ('bins', 'x_label', 'y_label'),
    'quotewall':        ('quotes',),
    'metro':            ('lines',),
    'balance':          ('left', 'right'),
    'fiveforces':       ('center', 'forces'),
    'glossary':         ('items',),
    'album':            ('title', 'tracks'),
    'bracket':          ('groups',),
    'horizon':          ('views',),
    'stack':            ('layers',),
    'gate':             ('layers',),
    'triad':            ('items', 'center'),
    'loop':             ('steps', 'center'),
    'ecosystem':        ('center', 'nodes'),
    'chronicle':        ('events',),
    'manifesto':        ('text', 'sub'),
    'comparetable':     ('headers', 'rows'),
    'dotfield':         ('items',),
    'bullet':           ('items',),
    'combo':            ('categories', 'bars', 'line'),
    'stream':           ('series', 'labels'),
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
    etree.SubElement(bodyPr, f'{{{A}}}normAutofit')   # 文字超出时自动缩小
    etree.SubElement(txBody, f'{{{A}}}lstStyle')

    for para in (paragraphs if paragraphs else [{'text': ''}]):
        p = etree.SubElement(txBody, f'{{{A}}}p')
        pPr = etree.SubElement(p, f'{{{A}}}pPr'); pPr.set('lvl', '0')
        if para.get('bullet'):
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

def _diag_line(x1, y1, x2, y2, color, w=19050):
    """Draw a diagonal line from (x1,y1) to (x2,y2) using a thin rotated rect."""
    import math
    dx = x2 - x1
    dy = y2 - y1
    length = int(math.sqrt(dx * dx + dy * dy))
    if length < 1:
        return None
    angle = math.atan2(dy, dx)
    angle_deg = int(math.degrees(angle) * 60000)
    h = w  # line thickness = stroke width
    # Correct offset so visual start = (x1, y1)
    off_x = int(x1 - (length * (1 - math.cos(angle)) + h * math.sin(angle)) / 2)
    off_y = int(y1 - (length * math.sin(angle) - h * (1 - math.cos(angle))) / 2)
    sp = etree.Element(f'{{{P}}}sp')
    nv = etree.SubElement(sp, f'{{{P}}}nvSpPr')
    c = etree.SubElement(nv, f'{{{P}}}cNvPr'); c.set('id', _next_id()); c.set('name', 'DiagLine')
    etree.SubElement(nv, f'{{{P}}}cNvSpPr'); etree.SubElement(nv, f'{{{P}}}nvPr')
    spPr = etree.SubElement(sp, f'{{{P}}}spPr')
    xf = etree.SubElement(spPr, f'{{{A}}}xfrm')
    xf.set('rot', str(angle_deg))
    o = etree.SubElement(xf, f'{{{A}}}off'); o.set('x', str(off_x)); o.set('y', str(off_y))
    e = etree.SubElement(xf, f'{{{A}}}ext'); e.set('cx', str(length)); e.set('cy', str(h))
    pg = etree.SubElement(spPr, f'{{{A}}}prstGeom'); pg.set('prst', 'rect')
    etree.SubElement(pg, f'{{{A}}}avLst')
    sf = etree.SubElement(spPr, f'{{{A}}}solidFill')
    cl = etree.SubElement(sf, f'{{{A}}}srgbClr'); cl.set('val', color)
    ln = etree.SubElement(spPr, f'{{{A}}}ln')
    etree.SubElement(ln, f'{{{A}}}noFill')
    return sp

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

# ── 圆角矩形（带填充+文本） ────────────────────────────
def make_rect(x, y, cx, cy, fill_color, name, paragraphs,
              radius=30000, line_color=None, line_w=0, anchor='ctr'):
    """
    圆角矩形形状，支持填充色、文本、圆角
    paragraphs: [{ text, bold, color, sz, font }]
    """
    sp = etree.Element(f'{{{P}}}sp')

    nv = etree.SubElement(sp, f'{{{P}}}nvSpPr')
    cNvPr = etree.SubElement(nv, f'{{{P}}}cNvPr')
    cNvPr.set('id', _next_id()); cNvPr.set('name', name)
    etree.SubElement(nv, f'{{{P}}}cNvSpPr')
    etree.SubElement(nv, f'{{{P}}}nvPr')

    spPr = etree.SubElement(sp, f'{{{P}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A}}}xfrm')
    off  = etree.SubElement(xfrm, f'{{{A}}}off'); off.set('x', str(x)); off.set('y', str(y))
    ext  = etree.SubElement(xfrm, f'{{{A}}}ext'); ext.set('cx', str(cx)); ext.set('cy', str(cy))
    prstGeom = etree.SubElement(spPr, f'{{{A}}}prstGeom')
    prstGeom.set('prst', 'roundRect')
    avLst = etree.SubElement(prstGeom, f'{{{A}}}avLst')
    gd = etree.SubElement(avLst, f'{{{A}}}gd'); gd.set('name', 'adj'); gd.set('fmla', f'val {radius}')

    fill = etree.SubElement(spPr, f'{{{A}}}solidFill')
    cl = etree.SubElement(fill, f'{{{A}}}srgbClr'); cl.set('val', fill_color)

    ln = etree.SubElement(spPr, f'{{{A}}}ln')
    ln.set('w', str(line_w) if line_w else '0')
    if line_color and line_w:
        lsf = etree.SubElement(ln, f'{{{A}}}solidFill')
        lcl = etree.SubElement(lsf, f'{{{A}}}srgbClr'); lcl.set('val', line_color)
    else:
        etree.SubElement(ln, f'{{{A}}}noFill')

    txBody = etree.SubElement(sp, f'{{{P}}}txBody')
    bodyPr = etree.SubElement(txBody, f'{{{A}}}bodyPr')
    bodyPr.set('wrap', 'square'); bodyPr.set('rtlCol', '0'); bodyPr.set('anchor', anchor)
    bodyPr.set('lIns', '91440'); bodyPr.set('rIns', '91440')
    bodyPr.set('tIns', '45720'); bodyPr.set('bIns', '45720')
    etree.SubElement(bodyPr, f'{{{A}}}normAutofit')   # 文字超出时自动缩小
    etree.SubElement(txBody, f'{{{A}}}lstStyle')

    for para in (paragraphs if paragraphs else [{'text': ''}]):
        p = etree.SubElement(txBody, f'{{{A}}}p')
        pPr = etree.SubElement(p, f'{{{A}}}pPr'); pPr.set('lvl', '0')
        algn = para.get('align')
        if algn:
            pPr.set('algn', algn)
        r = etree.SubElement(p, f'{{{A}}}r')
        rPr = etree.SubElement(r, f'{{{A}}}rPr')
        rPr.set('lang', 'zh-CN'); rPr.set('dirty', '0')
        if para.get('bold'):
            rPr.set('b', '1')
        if para.get('color'):
            sf = etree.SubElement(rPr, f'{{{A}}}solidFill')
            pcl = etree.SubElement(sf, f'{{{A}}}srgbClr'); pcl.set('val', para['color'])
        if para.get('sz'):
            rPr.set('sz', str(para['sz']))
        if para.get('font'):
            ea = etree.SubElement(rPr, f'{{{A}}}ea'); ea.set('typeface', para['font'])
        t = etree.SubElement(r, f'{{{A}}}t'); t.text = para.get('text', '')

    return sp

# ── 圆形（带填充+文本） ─────────────────────────────────
def make_circle(cx_pos, cy_pos, size, fill_color, name, paragraphs):
    """
    圆形形状，位置 (cx_pos, cy_pos) 为左上角
    paragraphs: [{ text, bold, color, sz, font }]
    """
    sp = etree.Element(f'{{{P}}}sp')

    nv = etree.SubElement(sp, f'{{{P}}}nvSpPr')
    cNvPr = etree.SubElement(nv, f'{{{P}}}cNvPr')
    cNvPr.set('id', _next_id()); cNvPr.set('name', name)
    etree.SubElement(nv, f'{{{P}}}cNvSpPr')
    etree.SubElement(nv, f'{{{P}}}nvPr')

    spPr = etree.SubElement(sp, f'{{{P}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A}}}xfrm')
    off  = etree.SubElement(xfrm, f'{{{A}}}off'); off.set('x', str(cx_pos)); off.set('y', str(cy_pos))
    ext  = etree.SubElement(xfrm, f'{{{A}}}ext'); ext.set('cx', str(size)); ext.set('cy', str(size))
    prstGeom = etree.SubElement(spPr, f'{{{A}}}prstGeom'); prstGeom.set('prst', 'ellipse')
    etree.SubElement(prstGeom, f'{{{A}}}avLst')

    fill = etree.SubElement(spPr, f'{{{A}}}solidFill')
    cl = etree.SubElement(fill, f'{{{A}}}srgbClr'); cl.set('val', fill_color)
    ln_el = etree.SubElement(spPr, f'{{{A}}}ln')
    etree.SubElement(ln_el, f'{{{A}}}noFill')

    txBody = etree.SubElement(sp, f'{{{P}}}txBody')
    bodyPr = etree.SubElement(txBody, f'{{{A}}}bodyPr')
    bodyPr.set('wrap', 'square'); bodyPr.set('rtlCol', '0'); bodyPr.set('anchor', 'ctr')
    bodyPr.set('lIns', '0'); bodyPr.set('rIns', '0')
    bodyPr.set('tIns', '0'); bodyPr.set('bIns', '0')
    etree.SubElement(txBody, f'{{{A}}}lstStyle')

    for para in (paragraphs if paragraphs else [{'text': ''}]):
        p = etree.SubElement(txBody, f'{{{A}}}p')
        pPr = etree.SubElement(p, f'{{{A}}}pPr'); pPr.set('lvl', '0')
        pPr.set('algn', 'ctr')
        r = etree.SubElement(p, f'{{{A}}}r')
        rPr = etree.SubElement(r, f'{{{A}}}rPr')
        rPr.set('lang', 'zh-CN'); rPr.set('dirty', '0')
        if para.get('bold'):
            rPr.set('b', '1')
        if para.get('color'):
            sf = etree.SubElement(rPr, f'{{{A}}}solidFill')
            pcl = etree.SubElement(sf, f'{{{A}}}srgbClr'); pcl.set('val', para['color'])
        if para.get('sz'):
            rPr.set('sz', str(para['sz']))
        if para.get('font'):
            ea = etree.SubElement(rPr, f'{{{A}}}ea'); ea.set('typeface', para['font'])
        t = etree.SubElement(r, f'{{{A}}}t'); t.text = para.get('text', '')

    return sp

# ── 通用形状（饼图扇形等）────────────────────────────────
def make_shape(x, y, cx, cy, prst, fill_color, name, paragraphs,
               adj=None, anchor='ctr'):
    """通用预设形状，prst: 'pie'/'donut'/'blockArc' 等"""
    sp = etree.Element(f'{{{P}}}sp')
    nv = etree.SubElement(sp, f'{{{P}}}nvSpPr')
    cNvPr = etree.SubElement(nv, f'{{{P}}}cNvPr')
    cNvPr.set('id', _next_id()); cNvPr.set('name', name)
    etree.SubElement(nv, f'{{{P}}}cNvSpPr')
    etree.SubElement(nv, f'{{{P}}}nvPr')

    spPr = etree.SubElement(sp, f'{{{P}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A}}}xfrm')
    off  = etree.SubElement(xfrm, f'{{{A}}}off'); off.set('x', str(x)); off.set('y', str(y))
    ext  = etree.SubElement(xfrm, f'{{{A}}}ext')
    ext.set('cx', str(cx)); ext.set('cy', str(cy))
    pg = etree.SubElement(spPr, f'{{{A}}}prstGeom'); pg.set('prst', prst)
    avLst = etree.SubElement(pg, f'{{{A}}}avLst')
    if adj is not None:
        gd = etree.SubElement(avLst, f'{{{A}}}gd'); gd.set('name', 'adj'); gd.set('fmla', f'val {adj}')

    fill = etree.SubElement(spPr, f'{{{A}}}solidFill')
    cl = etree.SubElement(fill, f'{{{A}}}srgbClr'); cl.set('val', fill_color)
    ln_el = etree.SubElement(spPr, f'{{{A}}}ln'); ln_el.set('w', '0')
    etree.SubElement(ln_el, f'{{{A}}}noFill')

    txBody = etree.SubElement(sp, f'{{{P}}}txBody')
    bodyPr = etree.SubElement(txBody, f'{{{A}}}bodyPr')
    bodyPr.set('wrap', 'square'); bodyPr.set('rtlCol', '0'); bodyPr.set('anchor', anchor)
    bodyPr.set('lIns', '45720'); bodyPr.set('rIns', '45720')
    bodyPr.set('tIns', '22860'); bodyPr.set('bIns', '22860')
    etree.SubElement(bodyPr, f'{{{A}}}normAutofit')
    etree.SubElement(txBody, f'{{{A}}}lstStyle')

    for para in (paragraphs if paragraphs else [{'text': ''}]):
        p = etree.SubElement(txBody, f'{{{A}}}p')
        pPr = etree.SubElement(p, f'{{{A}}}pPr'); pPr.set('lvl', '0')
        algn = para.get('align')
        if algn:
            pPr.set('algn', algn)
        r = etree.SubElement(p, f'{{{A}}}r')
        rPr = etree.SubElement(r, f'{{{A}}}rPr')
        rPr.set('lang', 'zh-CN'); rPr.set('dirty', '0')
        if para.get('bold'):
            rPr.set('b', '1')
        if para.get('color'):
            sf = etree.SubElement(rPr, f'{{{A}}}solidFill')
            pcl = etree.SubElement(sf, f'{{{A}}}srgbClr'); pcl.set('val', para['color'])
        if para.get('sz'):
            rPr.set('sz', str(para['sz']))
        if para.get('font'):
            ea = etree.SubElement(rPr, f'{{{A}}}ea'); ea.set('typeface', para['font'])
        t = etree.SubElement(r, f'{{{A}}}t'); t.text = para.get('text', '')

    return sp

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
    bar_pic = make_pic('rId2', 0, BAR_Y, SLIDE_W, BAR_H, 'BottomBar')  # rId2 已有 image11.png
    spTree.append(bar_pic)

    # 4) 标题
    paras = [{'text': data['title'], 'bold': True, 'color': BRAND_GREEN, 'sz': 3600}]
    title_box = make_textbox(BODY_X, TITLE_Y, BODY_W, emu(0.9), 'Title', paras)
    spTree.append(title_box)

    # 加载 rels 并补充缺失图片（底部栏、logo 等）
    rels_tree = etree.parse(rels_path)
    rr = rels_tree.getroot()
    existing_rids = {r.get('Id') for r in rr.findall('Relationship')}

    # 底部品牌栏
    # rId2 已在 slide4 模板中指向 image11.png，无需重复添加

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
        paras = [{'text': t, 'bold': False, 'color': '333333', 'sz': item_sz, 'bullet': True} for t in items]
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
        paras = [{'text': data.get('content', ''), 'bold': False, 'color': '333333', 'sz': item_sz}]
        return textbox(BODY_X, y, BODY_W, BODY_H, 'Content', paras)

    def _layout_stat(data, y):
        number = data.get('number', '0')
        unit   = data.get('unit', '')
        desc   = data.get('desc', '')
        ny = textbox(BODY_X, y, emu(5.0), emu(1.8), 'Number',
                     [{'text': number, 'bold': True, 'color': BRAND_GREEN, 'sz': 7200, 'font': 'Arial'}])
        uy = textbox(BODY_X, ny, emu(3.0), emu(0.6), 'Unit',
                     [{'text': unit, 'bold': False, 'color': 'FBB03B', 'sz': unit_sz}]) if unit else ny
        if desc:
            textbox(BODY_X, uy + emu(0.1), BODY_W, emu(0.6), 'Desc',
                    [{'text': desc, 'bold': False, 'color': '555555', 'sz': sub_sz}])
        return uy + emu(0.6) if desc else uy

    def _layout_two_column(data, y):
        lt = data.get('left_title', '')
        rt = data.get('right_title', '')
        textbox(LEFT_X, y, HALF_W, emu(0.5), 'LTitle',
                [{'text': lt, 'bold': True, 'color': GOLD, 'sz': 2200}])
        textbox(RIGHT_X, y, HALF_W, emu(0.5), 'RTitle',
                [{'text': rt, 'bold': True, 'color': GOLD, 'sz': 2200}])
        ly = textbox(LEFT_X, y + emu(0.6), HALF_W, emu(2.8), 'LBody',
                     [{'text': i, 'bold': False, 'color': '333333', 'sz': item_sz, 'bullet': True}
                      for i in data.get('left_items', [])])
        textbox(RIGHT_X, y + emu(0.6), HALF_W, emu(2.8), 'RBody',
                [{'text': i, 'bold': False, 'color': '333333', 'sz': item_sz, 'bullet': True}
                 for i in data.get('right_items', [])])
        return max(ly, y + emu(3.4))

    def _layout_quote(data, y):
        q = data.get('quote', '')
        src = data.get('source', '')
        qy = textbox(BODY_X, y, BODY_W, emu(2.5), 'Quote',
                     [{'text': f'"{q}"', 'bold': False, 'color': '333333', 'sz': 2800}])
        if src:
            textbox(BODY_X, qy + emu(0.1), BODY_W, emu(0.5), 'Source',
                    [{'text': f'— {src}', 'bold': False, 'color': GOLD, 'sz': 1400}])
        return qy + emu(0.6) if src else qy

    def _layout_image_text(data, y):
        # 左侧图片占位（实际图片需用户提供，这里用文本代替）
        textbox(LEFT_X, y, HALF_W, emu(2.8), 'ImagePlaceholder',
                [{'text': '[图片区域]', 'bold': False, 'color': BRAND_GREEN, 'sz': 1800}])
        st = data.get('side_title', '')
        if st:
            textbox(RIGHT_X, y, HALF_W, emu(0.5), 'SideTitle',
                    [{'text': st, 'bold': True, 'color': GOLD, 'sz': 2200}])
        textbox(RIGHT_X, y + emu(0.6), HALF_W, emu(2.4), 'SideBody',
                [{'text': i, 'bold': False, 'color': '333333', 'sz': item_sz, 'bullet': True}
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
                    [{'text': desc, 'bold': False, 'color': '333333', 'sz': item_sz}])
            y += emu(0.9)
        return y

    def _layout_metrics(data, y):
        metrics = data.get('metrics', [])
        if not metrics:
            return y
        n = min(len(metrics), 5)
        gap = emu(0.2)
        card_w = (BODY_W - (n - 1) * gap) // n
        for i, m in enumerate(metrics[:n]):
            cx = LEFT_X + i * (card_w + gap)
            label = m.get('label', '')
            value = m.get('value', '')
            unit  = m.get('unit', '')
            sub   = m.get('sub', '')
            textbox(cx, y, card_w, emu(0.4), f'MLabel{i}',
                    [{'text': label, 'bold': False, 'color': GOLD, 'sz': 1600}])
            textbox(cx, y + emu(0.45), card_w, emu(1.0), f'MVal{i}',
                    [{'text': value, 'bold': True, 'color': BRAND_GREEN, 'sz': 4800, 'font': 'Arial'}])
            if unit:
                textbox(cx, y + emu(1.5), card_w, emu(0.4), f'MUnit{i}',
                        [{'text': unit, 'bold': False, 'color': GOLD, 'sz': 2400}])
            if sub:
                textbox(cx, y + emu(1.95), card_w, emu(0.35), f'MSub{i}',
                        [{'text': sub, 'bold': False, 'color': '555555', 'sz': 1200}])
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
                    [{'text': st, 'bold': True, 'color': '333333', 'sz': 1800}])
            desc = step.get('desc', '')
            if desc:
                textbox(cx, y + emu(1.4), col_w, emu(1.5), f'StepDesc{i}',
                        [{'text': desc, 'bold': False, 'color': '555555', 'sz': sub_sz}])
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
                    [{'text': item.get('left', ''), 'bold': False, 'color': '333333', 'sz': 1500}])
            textbox(RIGHT_X, cy, HALF_W, emu(0.7), 'CmpR',
                    [{'text': item.get('right', ''), 'bold': False, 'color': '333333', 'sz': 1500}])
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
                    [{'text': lead, 'bold': False, 'color': '333333', 'sz': 1800}])
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
                        [{'text': str(cell), 'bold': False, 'color': '555555', 'sz': 1400}])
            ry += emu(0.45)
        return ry

    def _layout_team(data, y):
        members = data.get('members', [])
        n = min(len(members), 4)
        if n == 0:
            return y
        gap = emu(0.2)
        card_w = (BODY_W - (n - 1) * gap) // n
        for i, m in enumerate(members[:n]):
            cx = LEFT_X + i * (card_w + gap)
            name = m.get('name', '')
            role = m.get('role', '')
            desc = m.get('desc', '')
            textbox(cx, y, card_w, emu(0.5), f'TName{i}',
                    [{'text': name, 'bold': True, 'color': BRAND_GREEN, 'sz': 2200}])
            textbox(cx, y + emu(0.55), card_w, emu(0.35), f'TRole{i}',
                    [{'text': role, 'bold': True, 'color': GOLD, 'sz': 1400}])
            if desc:
                textbox(cx, y + emu(0.95), card_w, emu(1.5), f'TDesc{i}',
                        [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1200}])
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
                    [{'text': result, 'bold': True, 'color': BRAND_GREEN, 'sz': 2000}])
            ry += emu(0.6)
        if detail:
            textbox(BODY_X, ry, BODY_W, emu(2.0), 'CaseDetail',
                    [{'text': detail, 'bold': False, 'color': '333333', 'sz': item_sz}])
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
            fill = FUNNEL_COLORS[i % len(FUNNEL_COLORS)]
            _rect(cx, y, w, stage_h, fill, f'FunnelBg{i}', [], radius=40000)
            textbox(cx, y, w, stage_h, f'Funnel{i}',
                    [{'text': f'{lbl}  {val}', 'bold': True, 'color': WHITE, 'sz': 1600, 'align': 'ctr'}])
            y += stage_h + gap
        return y

    def _layout_takeaway(data, y):
        takeaways = data.get('takeaways', ['（暂无内容）'])
        paras = [{'text': t, 'bold': False, 'color': '333333', 'sz': 2000, 'bullet': True} for t in takeaways]
        return textbox(BODY_X, y, BODY_W, BODY_H, 'Takeaway', paras)

    # ── dashiai-ppt 风格新增布局 ────────────────────────────

    CARD_COLORS = ['46A53B', '213A25', 'FBB03B', '2D7A35', 'D4940A', '3B8A52']  # 品牌绿/深绿/金/中绿/深金/翠绿

    def _rect(x, y, cx, cy, fill, name, paras, **kw):
        r = make_rect(x, y, cx, cy, fill, name, paras, **kw)
        spTree.append(r)
        return y + cy

    def _circle(x, y, size, fill, name, paras):
        c = make_circle(x, y, size, fill, name, paras)
        spTree.append(c)
        return y + size

    def _accent_bar(x, y, cx, color=BRAND_GREEN):
        """左侧竖条装饰"""
        bar = make_rect(x, y, emu(0.08), emu(0.5), color, 'Accent', [], radius=0)
        spTree.append(bar)

    def _layout_comparison_cards(data, y):
        """两张并排圆角卡片对比（dashiai-ppt 风格）"""
        lt = data.get('left_title', '方案 A')
        rt = data.get('right_title', '方案 B')
        left_items = data.get('left_items', [])
        right_items = data.get('right_items', [])
        card_w = emu(5.6)
        card_h = emu(3.2)
        gap = emu(0.3)
        # 左卡片 - 浅绿背景
        _rect(LEFT_X, y, card_w, card_h, BRAND_GREEN, 'CardL', [], radius=30000,
              line_color=BRAND_GREEN, line_w=12700)
        textbox(LEFT_X + emu(0.2), y + emu(0.15), card_w - emu(0.4), emu(0.5), 'CLTitle',
                [{'text': lt, 'bold': True, 'color': WHITE, 'sz': 2000}])
        spTree.append(make_line(LEFT_X + emu(0.2), y + emu(0.7), card_w - emu(0.4), BRAND_GREEN, 12700))
        paras_l = [{'text': i, 'bold': False, 'color': WHITE, 'sz': 1500, 'bullet': True} for i in left_items]
        textbox(LEFT_X + emu(0.2), y + emu(0.85), card_w - emu(0.4), card_h - emu(1.0), 'CLBody', paras_l)
        # 右卡片 - 浅黄背景
        rx = LEFT_X + card_w + gap
        _rect(rx, y, card_w, card_h, DARK_GREEN, 'CardR', [], radius=30000,
              line_color=GOLD, line_w=12700)
        textbox(rx + emu(0.2), y + emu(0.15), card_w - emu(0.4), emu(0.5), 'CRTitle',
                [{'text': rt, 'bold': True, 'color': WHITE, 'sz': 2000}])
        spTree.append(make_line(rx + emu(0.2), y + emu(0.7), card_w - emu(0.4), GOLD, 12700))
        paras_r = [{'text': i, 'bold': False, 'color': WHITE, 'sz': 1500, 'bullet': True} for i in right_items]
        textbox(rx + emu(0.2), y + emu(0.85), card_w - emu(0.4), card_h - emu(1.0), 'CRBody', paras_r)
        return y + card_h + emu(0.2)

    def _layout_process_circles(data, y):
        """编号圆圈 + 连接线 + 描述（dashiai-ppt 步骤页风格）"""
        steps = data.get('steps', [])
        n = len(steps)
        if n == 0:
            return y
        gap = emu(0.2)
        circle_sz = emu(0.85)
        total_w = n * circle_sz + (n - 1) * gap
        # 如果太宽就压缩
        col_w = min((BODY_W - (n - 1) * emu(0.3)) // n, emu(3.0))
        actual_gap = (BODY_W - n * col_w) // max(n - 1, 1)
        circle_sz = min(circle_sz, col_w)

        for i, step in enumerate(steps):
            cx = LEFT_X + i * (col_w + actual_gap)
            # 圆圈
            circle_x = cx + (col_w - circle_sz) // 2
            _circle(circle_x, y, circle_sz, BRAND_GREEN, f'Circle{i}',
                    [{'text': str(i + 1), 'bold': True, 'color': WHITE, 'sz': 2400, 'font': 'Arial'}])
            # 标题
            st = step.get('title', f'步骤{i+1}')
            textbox(cx, y + circle_sz + emu(0.15), col_w, emu(0.5), f'PTitle{i}',
                    [{'text': st, 'bold': True, 'color': '333333', 'sz': 1800}])
            # 描述
            desc = step.get('desc', '')
            if desc:
                textbox(cx, y + circle_sz + emu(0.65), col_w, emu(1.5), f'PDesc{i}',
                        [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1400}])
            # 连接线（除了最后一个）
            if i < n - 1:
                arrow_x = cx + col_w
                arrow_y = y + circle_sz // 2
                spTree.append(make_line(arrow_x, arrow_y, actual_gap, GOLD, 19050))
        return y + circle_sz + emu(2.2)

    def _layout_tag_list(data, y):
        """彩色标签 + 描述列表（dashiai-ppt factors/tag 风格）"""
        tags = data.get('tags', [])
        if not tags:
            return y
        TAG_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'EE822F', '30C0B4', 'E54C5E']
        for i, tag in enumerate(tags):
            color = TAG_COLORS[i % len(TAG_COLORS)]
            term = tag.get('term', '')
            note = tag.get('note', '')
            # 标签圆角矩形
            tag_w = emu(2.0)
            tag_h = emu(0.45)
            _rect(BODY_X, y, tag_w, tag_h, color, f'Tag{i}',
                  [{'text': term, 'bold': True, 'color': WHITE, 'sz': 1400}],
                  radius=60000, anchor='ctr')
            # 描述文本
            if note:
                textbox(BODY_X + tag_w + emu(0.2), y, BODY_W - tag_w - emu(0.2), tag_h, f'TNote{i}',
                        [{'text': note, 'bold': False, 'color': '333333', 'sz': 1500}])
            y += tag_h + emu(0.15)
        return y

    def _layout_highlight_box(data, y):
        """大数字高亮框 + 副指标（dashiai-ppt 数据页风格）"""
        big_num = data.get('big_number', '0')
        big_label = data.get('big_label', '')
        secondaries = data.get('secondaries', [])
        # 主高亮卡片
        box_h = emu(2.2)
        _rect(BODY_X, y, BODY_W, box_h, BRAND_GREEN, 'HBox', [], radius=30000,
              line_color=BRAND_GREEN, line_w=19050)
        # 大数字
        textbox(BODY_X, y + emu(0.2), BODY_W, emu(1.2), 'BigNum',
                [{'text': str(big_num), 'bold': True, 'color': WHITE, 'sz': 6000, 'font': 'Arial'}])
        # 标签
        if big_label:
            textbox(BODY_X, y + emu(1.3), BODY_W, emu(0.5), 'BigLabel',
                    [{'text': big_label, 'bold': False, 'color': WHITE, 'sz': 1800}])
        y += box_h + emu(0.3)
        # 副指标卡片
        if secondaries:
            n = len(secondaries)
            sec_w = (BODY_W - (n - 1) * emu(0.2)) // n
            for i, sec in enumerate(secondaries):
                sx = BODY_X + i * (sec_w + emu(0.2))
                _rect(sx, y, sec_w, emu(1.3), CARD_COLORS[i % len(CARD_COLORS)], f'Sec{i}', [],
                      radius=30000)
                sv = sec.get('value', '')
                sl = sec.get('label', '')
                textbox(sx, y + emu(0.1), sec_w, emu(0.7), f'SV{i}',
                        [{'text': sv, 'bold': True, 'color': WHITE, 'sz': 2800, 'font': 'Arial'}])
                if sl:
                    textbox(sx, y + emu(0.8), sec_w, emu(0.4), f'SL{i}',
                            [{'text': sl, 'bold': False, 'color': WHITE, 'sz': 1200}])
            y += emu(1.5)
        return y

    def _layout_dual_stat(data, y):
        """两个大数字并排对比"""
        lv = data.get('left_value', '0')
        ll = data.get('left_label', '')
        rv = data.get('right_value', '0')
        rl = data.get('right_label', '')
        half_w = emu(5.5)
        gap = emu(0.5)
        # 左
        _rect(LEFT_X, y, half_w, emu(2.0), BRAND_GREEN, 'DStatL', [], radius=30000)
        textbox(LEFT_X, y + emu(0.2), half_w, emu(1.0), 'DLV',
                [{'text': str(lv), 'bold': True, 'color': WHITE, 'sz': 4800, 'font': 'Arial'}])
        if ll:
            textbox(LEFT_X, y + emu(1.2), half_w, emu(0.5), 'DLL',
                    [{'text': ll, 'bold': False, 'color': WHITE, 'sz': 1600}])
        # 右
        rx = LEFT_X + half_w + gap
        _rect(rx, y, half_w, emu(2.0), GOLD, 'DStatR', [], radius=30000)
        textbox(rx, y + emu(0.2), half_w, emu(1.0), 'DRV',
                [{'text': str(rv), 'bold': True, 'color': WHITE, 'sz': 4800, 'font': 'Arial'}])
        if rl:
            textbox(rx, y + emu(1.2), half_w, emu(0.5), 'DRL',
                    [{'text': rl, 'bold': False, 'color': WHITE, 'sz': 1600}])
        return y + emu(2.2)

    def _layout_card_grid(data, y):
        """网格卡片布局（2×N 或 3×N）"""
        cards = data.get('cards', [])
        if not cards:
            return y
        n = len(cards)
        cols = 3 if n >= 3 else 2
        rows_n = (n + cols - 1) // cols
        gap = emu(0.25)
        card_w = (BODY_W - (cols - 1) * gap) // cols
        card_h = emu(2.0)
        for i, card in enumerate(cards):
            row = i // cols
            col = i % cols
            cx = LEFT_X + col * (card_w + gap)
            cy = y + row * (card_h + gap)
            fill = CARD_COLORS[i % len(CARD_COLORS)]
            title = card.get('title', '')
            desc = card.get('desc', '')
            _rect(cx, cy, card_w, card_h, fill, f'Card{i}', [], radius=4374)
            # 左侧竖条装饰
            _accent_bar(cx + emu(0.1), cy + emu(0.15), emu(0.4), [BRAND_GREEN, GOLD, '4874CB'][i % 3])
            textbox(cx + emu(0.3), cy + emu(0.15), card_w - emu(0.4), emu(0.5), f'CT{i}',
                    [{'text': title, 'bold': True, 'color': WHITE, 'sz': 1600}])
            if desc:
                textbox(cx + emu(0.3), cy + emu(0.65), card_w - emu(0.4), card_h - emu(0.8), f'CD{i}',
                        [{'text': desc, 'bold': False, 'color': WHITE, 'sz': 1300}])
        return y + rows_n * (card_h + gap)

    def _layout_before_after(data, y):
        """Before/After 对比卡片 + 中间箭头（dashiai-ppt 对比页风格）"""
        bt = data.get('before_title', 'Before')
        bi = data.get('before_items', [])
        at = data.get('after_title', 'After')
        ai = data.get('after_items', [])
        card_w = emu(5.2)
        card_h = emu(3.2)
        arrow_w = emu(0.8)
        # Before 卡片
        _rect(LEFT_X, y, card_w, card_h, 'C0392B', 'BeforeCard', [], radius=30000,
              line_color='E54C5E', line_w=12700)
        textbox(LEFT_X + emu(0.2), y + emu(0.15), card_w - emu(0.4), emu(0.5), 'BTitle',
                [{'text': bt, 'bold': True, 'color': WHITE, 'sz': 2000}])
        spTree.append(make_line(LEFT_X + emu(0.2), y + emu(0.7), card_w - emu(0.4), 'E54C5E', 12700))
        paras_b = [{'text': i, 'bold': False, 'color': WHITE, 'sz': 1500, 'bullet': True} for i in bi]
        textbox(LEFT_X + emu(0.2), y + emu(0.85), card_w - emu(0.4), card_h - emu(1.0), 'BBody', paras_b)
        # 中间箭头区域
        ax = LEFT_X + card_w + emu(0.15)
        _rect(ax, y + card_h // 2 - emu(0.3), arrow_w, emu(0.6), BRAND_GREEN, 'Arrow',
              [{'text': '→', 'bold': True, 'color': WHITE, 'sz': 2400, 'align': 'ctr'}],
              radius=60000, anchor='ctr')
        # After 卡片
        rx = ax + arrow_w + emu(0.15)
        _rect(rx, y, card_w, card_h, BRAND_GREEN, 'AfterCard', [], radius=30000,
              line_color=BRAND_GREEN, line_w=12700)
        textbox(rx + emu(0.2), y + emu(0.15), card_w - emu(0.4), emu(0.5), 'ATitle',
                [{'text': at, 'bold': True, 'color': WHITE, 'sz': 2000}])
        spTree.append(make_line(rx + emu(0.2), y + emu(0.7), card_w - emu(0.4), BRAND_GREEN, 12700))
        paras_a = [{'text': i, 'bold': False, 'color': WHITE, 'sz': 1500, 'bullet': True} for i in ai]
        textbox(rx + emu(0.2), y + emu(0.85), card_w - emu(0.4), card_h - emu(1.0), 'ABody', paras_a)
        return y + card_h + emu(0.2)

    # ── 新增布局（参考 dashi-ppt 语义角色）────────────────────

    FUNNEL_COLORS = ['46A53B', '2D7A35', '3B8A52', '213A25', '1A5C20']

    def _layout_swot(data, y):
        """SWOT 四象限分析（dashi-ppt: swot 角色）"""
        quadrants = data.get('quadrants', {})
        labels = data.get('labels', {'S': '优势 S', 'W': '劣势 W', 'O': '机会 O', 'T': '威胁 T'})
        colors = [BRAND_GREEN, 'C0392B', GOLD, DARK_GREEN]
        keys = ['S', 'W', 'O', 'T']
        half_w = (BODY_W - emu(0.2)) // 2
        half_h = emu(1.6)   # 2.2→1.6，防止超出底部栏
        for i, key in enumerate(keys):
            col = i % 2
            row = i // 2
            cx = LEFT_X + col * (half_w + emu(0.2))
            cy = y + row * (half_h + emu(0.2))
            fill = colors[i]
            _rect(cx, cy, half_w, half_h, fill, f'SWOT{i}', [], radius=10000)
            label = labels.get(key, key)
            items = quadrants.get(key, [])
            textbox(cx + emu(0.2), cy + emu(0.15), half_w - emu(0.4), emu(0.45), f'SWOTLbl{i}',
                    [{'text': label, 'bold': True, 'color': WHITE, 'sz': 1800}])
            if items:
                paras = [{'text': t, 'bold': False, 'color': WHITE, 'sz': 1300, 'bullet': True} for t in items]
                textbox(cx + emu(0.2), cy + emu(0.65), half_w - emu(0.4), half_h - emu(0.8), f'SWOTBd{i}', paras)
        return y + 2 * half_h + emu(0.4)

    def _layout_quadrant(data, y):
        """四象限矩阵（dashi-ppt: quadrant/matrix 角色）"""
        items = data.get('items', {'q1': [], 'q2': [], 'q3': [], 'q4': []})
        axis_x = data.get('axis_x', '')
        axis_y = data.get('axis_y', '')
        labels = data.get('labels', {'q1': 'Ⅰ', 'q2': 'Ⅱ', 'q3': 'Ⅲ', 'q4': 'Ⅳ'})
        fills = ['F0F7EE', 'FFF8EC', 'F5F0FA', 'FAF0F0']
        border_colors = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E']
        half_w = (BODY_W - emu(0.2)) // 2
        half_h = emu(1.6)   # 2.2→1.6，防止超出底部栏
        for i, key in enumerate(['q1', 'q2', 'q3', 'q4']):
            col = i % 2
            row = i // 2
            cx = LEFT_X + col * (half_w + emu(0.2))
            cy = y + row * (half_h + emu(0.2))
            _rect(cx, cy, half_w, half_h, border_colors[i], f'Quad{i}', [], radius=10000)
            lbl = labels.get(key, '')
            q_items = items.get(key, [])
            textbox(cx + emu(0.2), cy + emu(0.15), half_w - emu(0.4), emu(0.45), f'QLbl{i}',
                    [{'text': lbl, 'bold': True, 'color': WHITE, 'sz': 1800}])
            if q_items:
                paras = [{'text': t, 'bold': False, 'color': WHITE, 'sz': 1300, 'bullet': True} for t in q_items]
                textbox(cx + emu(0.2), cy + emu(0.65), half_w - emu(0.4), half_h - emu(0.8), f'QBd{i}', paras)
        return y + 2 * half_h + emu(0.4)

    def _layout_checklist(data, y):
        """清单打卡（dashi-ppt: checklist 角色）"""
        items = data.get('items', [])
        if not items:
            return y
        n = min(len(items), 6)
        for i, item in enumerate(items[:n]):
            text = item.get('text', '') if isinstance(item, dict) else str(item)
            done = item.get('done', False) if isinstance(item, dict) else False
            row_h = emu(0.6)
            # 勾选框
            box_color = BRAND_GREEN if done else 'CCCCCC'
            _rect(BODY_X, y + (row_h - emu(0.35)) // 2, emu(0.35), emu(0.35), box_color, f'ChkBox{i}',
                  [{'text': '✓' if done else '', 'bold': True, 'color': WHITE, 'sz': 1400, 'align': 'ctr'}],
                  radius=30000, anchor='ctr')
            # 文字
            txt_color = '333333' if not done else BRAND_GREEN
            textbox(BODY_X + emu(0.55), y, BODY_W - emu(0.55), row_h, f'ChkTxt{i}',
                    [{'text': text, 'bold': done, 'color': txt_color, 'sz': 1600}])
            y += row_h + emu(0.08)
        return y

    def _layout_scorecard(data, y):
        """评分卡（dashi-ppt: scorecard 角色）"""
        items = data.get('items', [])
        if not items:
            return y
        row_h = emu(0.7)
        for i, item in enumerate(items):
            label = item.get('label', '')
            score = item.get('score', 0)
            max_score = item.get('max', 5)
            desc = item.get('desc', '')
            # 标签
            textbox(BODY_X, y, emu(3.0), emu(0.4), f'ScLbl{i}',
                    [{'text': label, 'bold': True, 'color': '333333', 'sz': 1600}])
            # 分数
            textbox(BODY_X + emu(3.2), y, emu(1.0), emu(0.4), f'ScVal{i}',
                    [{'text': f'{score}/{max_score}', 'bold': True, 'color': BRAND_GREEN, 'sz': 1600, 'font': 'Arial'}])
            # 进度条背景
            bar_x = BODY_X
            bar_y = y + emu(0.42)
            bar_w = BODY_W
            bar_h = emu(0.15)
            _rect(bar_x, bar_y, bar_w, bar_h, 'E8E8E8', f'ScBg{i}', [], radius=20000)
            # 进度条填充
            fill_w = int(bar_w * score / max_score) if max_score > 0 else 0
            if fill_w > 0:
                fill_color = BRAND_GREEN if score / max_score >= 0.7 else GOLD if score / max_score >= 0.4 else 'C0392B'
                _rect(bar_x, bar_y, fill_w, bar_h, fill_color, f'ScFill{i}', [], radius=20000)
            if desc:
                textbox(BODY_X + emu(4.5), y, BODY_W - emu(4.5), emu(0.4), f'ScDesc{i}',
                        [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1200}])
            y += row_h
        return y

    def _layout_stair(data, y):
        """阶梯递进（dashi-ppt: stair 角色）"""
        steps = data.get('steps', [])
        if not steps:
            return y
        n = len(steps)
        step_w = (BODY_W - (n - 1) * emu(0.15)) // n
        base_h = emu(0.8)
        step_inc = emu(0.5)
        for i, step in enumerate(steps):
            sx = LEFT_X + i * (step_w + emu(0.15))
            sh = base_h + i * step_inc
            sy = y + emu(3.0) - sh  # 底部对齐
            fill = FUNNEL_COLORS[i % len(FUNNEL_COLORS)]
            _rect(sx, sy, step_w, sh, fill, f'Stair{i}', [], radius=30000)
            title = step.get('title', f'阶段{i+1}')
            desc = step.get('desc', '')
            textbox(sx + emu(0.1), sy + emu(0.1), step_w - emu(0.2), emu(0.45), f'StairT{i}',
                    [{'text': title, 'bold': True, 'color': WHITE, 'sz': 1500}])
            if desc:
                textbox(sx + emu(0.1), sy + emu(0.55), step_w - emu(0.2), sh - emu(0.65), f'StairD{i}',
                        [{'text': desc, 'bold': False, 'color': WHITE, 'sz': 1200}])
        return y + emu(3.2)

    def _layout_flywheel(data, y):
        """飞轮循环（dashi-ppt: flywheel/cyclewheel 角色）"""
        items = data.get('items', [])
        center_label = data.get('center', '')
        if not items:
            return y
        n = min(len(items), 6)
        area_sz = emu(3.5)
        cx_center = LEFT_X + (BODY_W - area_sz) // 2
        cy_center = y
        # 中心圆
        circle_sz = emu(1.5)
        circle_x = cx_center + (area_sz - circle_sz) // 2
        circle_y = cy_center + (area_sz - circle_sz) // 2
        _circle(circle_x, circle_y, circle_sz, BRAND_GREEN, 'FlyCenter',
                [{'text': center_label, 'bold': True, 'color': WHITE, 'sz': 1400, 'align': 'ctr'}])
        # 周围节点
        node_sz = emu(1.2)
        radius = (area_sz - node_sz) // 2
        for i, item in enumerate(items[:6]):
            angle = 2 * math.pi * i / min(n, 6) - math.pi / 2
            nx = int(cx_center + area_sz // 2 + radius * math.cos(angle) - node_sz // 2)
            ny = int(cy_center + area_sz // 2 + radius * math.sin(angle) - node_sz // 2)
            fill = FUNNEL_COLORS[i % len(FUNNEL_COLORS)]
            label = item.get('label', '') if isinstance(item, dict) else str(item)
            _circle(nx, ny, node_sz, fill, f'FlyNode{i}',
                    [{'text': label, 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}])
        return y + area_sz + emu(0.2)

    def _layout_statement(data, y):
        """大字宣言（dashi-ppt: statement/manifesto 角色）"""
        text = data.get('text', '')
        sub = data.get('sub', '')
        # 大号引号装饰
        textbox(BODY_X, y, emu(1.0), emu(1.0), 'QuoteMark',
                [{'text': '"', 'bold': True, 'color': GOLD, 'sz': 7200, 'font': 'Arial'}])
        # 主文字
        textbox(BODY_X + emu(0.3), y + emu(0.8), BODY_W - emu(0.6), emu(2.8), 'Statement',
                [{'text': text, 'bold': True, 'color': DARK_GREEN, 'sz': 2800}])
        if sub:
            textbox(BODY_X + emu(0.3), y + emu(3.5), BODY_W - emu(0.6), emu(0.5), 'StmtSub',
                    [{'text': sub, 'bold': False, 'color': GOLD, 'sz': 1600}])
        return y + emu(4.0) if sub else y + emu(3.6)

    def _layout_journey(data, y):
        """用户旅程图（dashi-ppt: journey 角色）"""
        stages = data.get('stages', [])
        if not stages:
            return y
        n = len(stages)
        col_w = (BODY_W - (n - 1) * emu(0.15)) // n
        # 顶部阶段标签
        for i, stage in enumerate(stages):
            sx = LEFT_X + i * (col_w + emu(0.15))
            label = stage.get('stage', f'阶段{i+1}')
            _rect(sx, y, col_w, emu(0.5), BRAND_GREEN, f'JStage{i}',
                  [{'text': label, 'bold': True, 'color': WHITE, 'sz': 1400, 'align': 'ctr'}],
                  radius=30000, anchor='ctr')
        # 下方内容区
        content_y = y + emu(0.7)
        rows = data.get('rows', [])
        if rows:
            row_h = emu(0.6)
            for j, row in enumerate(rows):
                row_label = row.get('label', '')
                textbox(BODY_X, content_y + j * (row_h + emu(0.1)), emu(2.0), row_h, f'JRowLbl{j}',
                        [{'text': row_label, 'bold': True, 'color': '333333', 'sz': 1300}])
                cells = row.get('cells', [])
                for i, cell in enumerate(cells[:n]):
                    sx = LEFT_X + i * (col_w + emu(0.15))
                    fill = 'F0F7EE' if j % 2 == 0 else 'FFFFFF'
                    _rect(sx, content_y + j * (row_h + emu(0.1)), col_w, row_h, fill, f'JCell{j}{i}', [], radius=20000)
                    textbox(sx + emu(0.08), content_y + j * (row_h + emu(0.1)), col_w - emu(0.16), row_h, f'JTxt{j}{i}',
                            [{'text': str(cell), 'bold': False, 'color': '333333', 'sz': 1200}])
            return content_y + len(rows) * (row_h + emu(0.1))
        return content_y + emu(0.5)

    def _layout_pricing(data, y):
        """价格方案（dashi-ppt: pricing 角色）"""
        plans = data.get('plans', [])
        if not plans:
            return y
        n = len(plans)
        card_w = (BODY_W - (n - 1) * emu(0.25)) // n
        card_h = emu(3.0)   # 3.5→3.0，更紧凑，防止超出
        for i, plan in enumerate(plans):
            cx = LEFT_X + i * (card_w + emu(0.25))
            name = plan.get('name', '')
            price = plan.get('price', '')
            unit = plan.get('unit', '')
            features = plan.get('features', [])
            highlight = plan.get('highlight', False)
            fill = BRAND_GREEN if highlight else DARK_GREEN
            _rect(cx, y, card_w, card_h, fill, f'Price{i}', [], radius=8000)
            # 方案名
            textbox(cx, y + emu(0.15), card_w, emu(0.35), f'PName{i}',
                    [{'text': name, 'bold': True, 'color': WHITE, 'sz': 1800, 'align': 'ctr'}])
            # 价格
            textbox(cx, y + emu(0.55), card_w, emu(0.55), f'PPrice{i}',
                    [{'text': price, 'bold': True, 'color': GOLD if highlight else WHITE, 'sz': 3000, 'font': 'Arial', 'align': 'ctr'}])
            if unit:
                textbox(cx, y + emu(1.10), card_w, emu(0.25), f'PUnit{i}',
                        [{'text': unit, 'bold': False, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}])
            # 特性列表
            if features:
                paras = [{'text': f, 'bold': False, 'color': WHITE, 'sz': 1300, 'bullet': True} for f in features]
                textbox(cx + emu(0.15), y + emu(1.40), card_w - emu(0.3), card_h - emu(1.55), f'PFeat{i}', paras)
        return y + card_h + emu(0.2)

    def _layout_faq(data, y):
        """FAQ 常见问题（dashi-ppt: faq 角色）"""
        items = data.get('items', [])
        if not items:
            return y
        n = min(len(items), 5)
        for i, item in enumerate(items[:n]):
            q = item.get('q', '') if isinstance(item, dict) else ''
            a = item.get('a', '') if isinstance(item, dict) else ''
            # Q 标签
            _rect(BODY_X, y, emu(0.45), emu(0.45), BRAND_GREEN, f'FAQQ{i}',
                  [{'text': 'Q', 'bold': True, 'color': WHITE, 'sz': 1600, 'align': 'ctr'}],
                  radius=30000, anchor='ctr')
            textbox(BODY_X + emu(0.6), y, BODY_W - emu(0.6), emu(0.45), f'FAQQt{i}',
                    [{'text': q, 'bold': True, 'color': '333333', 'sz': 1600}])
            y += emu(0.55)
            if a:
                _rect(BODY_X, y, emu(0.45), emu(0.45), GOLD, f'FAQAl{i}',
                      [{'text': 'A', 'bold': True, 'color': WHITE, 'sz': 1600, 'align': 'ctr'}],
                      radius=30000, anchor='ctr')
                textbox(BODY_X + emu(0.6), y, BODY_W - emu(0.6), emu(0.7), f'FAQAt{i}',
                        [{'text': a, 'bold': False, 'color': '555555', 'sz': 1400}])
                y += emu(0.7)
            y += emu(0.15)
        return y

    # ── 本轮新增 9 种布局（参考 dashi-ppt 语义角色）─────────

    PIE_COLORS = [BRAND_GREEN, GOLD, '2D7A35', 'D4940A', '3B8A52', DARK_GREEN, '4874CB', 'E54C5E']

    def _make_sp(sp_tree, shape):
        sp_tree.append(shape)

    def _layout_pie_chart(data, y):
        """饼图 / 环形图（dashi-ppt: pie/donut 角色）"""
        slices = data.get('slices', [])
        if not slices:
            return y
        is_donut = data.get('donut', False)
        total = sum(s.get('value', 0) for s in slices) or 1
        # 左侧圆图区域
        chart_sz = emu(2.8)
        chart_x = LEFT_X
        chart_y = y
        # 画扇形（简化：用圆形 + 百分比标签覆盖）
        _circle(chart_x, chart_y, chart_sz, 'E8E8E8', 'PieBg', [])
        # 按百分比画扇形色块（用 pie 几何形状）
        angle_start = 0
        for i, sl in enumerate(slices[:8]):
            val = sl.get('value', 0)
            pct = val / total
            angle_span = pct * 360  # 本扇形的角度跨度
            adj = int(angle_span * 60000 / 360)  # adj = 角度跨度（60000制）
            rot = int(angle_start * 60000)  # rot = 起始角度（60000制）
            color = PIE_COLORS[i % len(PIE_COLORS)]
            # 用 make_shape 画扇形（adj=跨度，不是起始位置）
            pie = make_shape(chart_x, chart_y, chart_sz, chart_sz, 'pie', color, f'Pie{i}',
                            [], adj=adj)
            spTree.append(pie)
            # 旋转形状：通过 xfrm rot 属性设置起始角度
            xfrm_el = pie.find(f'.//{{{A}}}xfrm')
            if xfrm_el is not None:
                xfrm_el.set('rot', str(rot))
            angle_start += angle_span
        # 环形图中心圆（覆盖中心）
        if is_donut:
            hole_sz = emu(1.4)
            hole_off = (chart_sz - hole_sz) // 2
            _circle(chart_x + hole_off, chart_y + hole_off, hole_sz, WHITE, 'DonutHole', [])
        # 右侧图例
        legend_x = LEFT_X + chart_sz + emu(0.4)
        legend_y = y + emu(0.2)
        for i, sl in enumerate(slices[:8]):
            val = sl.get('value', 0)
            pct = val / total * 100
            label = sl.get('label', f'项目{i+1}')
            color = PIE_COLORS[i % len(PIE_COLORS)]
            ly = legend_y + i * emu(0.38)
            _rect(legend_x, ly, emu(0.28), emu(0.28), color, f'Leg{i}', [], radius=20000)
            textbox(legend_x + emu(0.38), ly - emu(0.02), emu(4.5), emu(0.35), f'LegT{i}',
                    [{'text': f'{label}  {pct:.0f}%', 'bold': False, 'color': '333333', 'sz': 1300}])
        return y + chart_sz + emu(0.2)

    def _layout_bar_chart(data, y):
        """横向条形图（dashi-ppt: bar-chart 角色）"""
        bars = data.get('bars', [])
        if not bars:
            return y
        max_val = max(b.get('value', 0) for b in bars) or 1
        bar_h = emu(0.45)
        label_w = emu(2.2)
        bar_max_w = BODY_W - label_w - emu(1.2)   # 留空给数值文字
        for i, bar in enumerate(bars[:8]):
            label = bar.get('label', '')
            value = bar.get('value', 0)
            unit = bar.get('unit', '')
            color = bar.get('color', PIE_COLORS[i % len(PIE_COLORS)])
            bar_w = int(bar_max_w * value / max_val) if max_val > 0 else 0
            bar_w = max(bar_w, emu(0.3))
            row_y = y + i * (bar_h + emu(0.18))
            # 标签
            textbox(LEFT_X, row_y, label_w, bar_h, f'BarLbl{i}',
                    [{'text': label, 'bold': False, 'color': '333333', 'sz': 1400}])
            # 背景条
            _rect(LEFT_X + label_w, row_y + emu(0.05), bar_max_w, bar_h - emu(0.1), 'F0F0F0', f'BarBg{i}', [], radius=20000)
            # 数据条
            _rect(LEFT_X + label_w, row_y + emu(0.05), bar_w, bar_h - emu(0.1), color, f'Bar{i}', [], radius=20000)
            # 数值
            textbox(LEFT_X + label_w + bar_w + emu(0.12), row_y, emu(1.0), bar_h, f'BarVal{i}',
                    [{'text': f'{value}{unit}', 'bold': True, 'color': '333333', 'sz': 1300, 'font': 'Arial'}])
        return y + len(bars[:8]) * (bar_h + emu(0.18))

    def _layout_dashboard(data, y):
        """仪表盘 2×3 指标卡（dashi-ppt: dashboard/metrics-grid 角色）"""
        cards = data.get('cards', [])
        if not cards:
            return y
        n = min(len(cards), 6)
        cols = min(3, n)
        rows_n = (n + cols - 1) // cols
        gap = emu(0.2)
        card_w = (BODY_W - (cols - 1) * gap) // cols
        card_h = emu(1.55)
        for i, card in enumerate(cards[:n]):
            col = i % cols
            row = i // cols
            cx = LEFT_X + col * (card_w + gap)
            cy = y + row * (card_h + gap)
            label = card.get('label', '')
            value = card.get('value', '0')
            unit = card.get('unit', '')
            trend = card.get('trend', '')   # 如 ↑12% ↓5%
            fill = CARD_COLORS[i % len(CARD_COLORS)]
            _rect(cx, cy, card_w, card_h, fill, f'DashBg{i}', [], radius=15000)
            textbox(cx + emu(0.15), cy + emu(0.15), card_w - emu(0.3), emu(0.35), f'DashLbl{i}',
                    [{'text': label, 'bold': False, 'color': WHITE, 'sz': 1300}])
            val_text = f'{value}{unit}'
            textbox(cx + emu(0.15), cy + emu(0.50), card_w - emu(0.3), emu(0.55), f'DashVal{i}',
                    [{'text': val_text, 'bold': True, 'color': WHITE, 'sz': 2600, 'font': 'Arial'}])
            if trend:
                textbox(cx + emu(0.15), cy + emu(1.10), card_w - emu(0.3), emu(0.30), f'DashTr{i}',
                        [{'text': trend, 'bold': False, 'color': GOLD, 'sz': 1200}])
        return y + rows_n * (card_h + gap)

    def _layout_hero_banner(data, y):
        """英雄横幅（dashi-ppt: hero/cover-card 角色）"""
        subtitle = data.get('subtitle', '')
        kicker = data.get('kicker', '')
        tagline = data.get('tagline', '')
        # 全宽大卡片
        banner_h = emu(3.0)
        _rect(LEFT_X, y, BODY_W, banner_h, BRAND_GREEN, 'HeroBg', [], radius=15000)
        # 金色装饰线
        spTree.append(make_line(LEFT_X + emu(0.5), y + emu(0.4), emu(1.2), GOLD, 25400))
        # kicker
        if kicker:
            textbox(LEFT_X + emu(0.5), y + emu(0.55), BODY_W - emu(1.0), emu(0.4), 'HeroKick',
                    [{'text': kicker, 'bold': True, 'color': GOLD, 'sz': 1600}])
        # 大标题
        title_text = data.get('title', '')
        textbox(LEFT_X + emu(0.5), y + emu(1.0), BODY_W - emu(1.0), emu(1.0), 'HeroTitle',
                [{'text': title_text, 'bold': True, 'color': WHITE, 'sz': 3600}])
        # 副标题
        if subtitle:
            textbox(LEFT_X + emu(0.5), y + emu(1.9), BODY_W - emu(1.0), emu(0.5), 'HeroSub',
                    [{'text': subtitle, 'bold': False, 'color': WHITE, 'sz': 1600}])
        # 标语
        if tagline:
            textbox(LEFT_X + emu(0.5), y + emu(2.4), BODY_W - emu(1.0), emu(0.4), 'HeroTag',
                    [{'text': tagline, 'bold': False, 'color': GOLD, 'sz': 1400}])
        return y + banner_h + emu(0.2)

    def _layout_numbered_list(data, y):
        """编号列表（dashi-ppt: numbered-list/ordered-steps 角色）"""
        items = data.get('items', [])
        if not items:
            return y
        for i, item in enumerate(items[:8]):
            text = item.get('text', '') if isinstance(item, dict) else str(item)
            desc = item.get('desc', '') if isinstance(item, dict) else ''
            row_h = emu(0.48) if not desc else emu(0.62)
            # 数字圆圈
            circle_sz = emu(0.38)
            circle_y = y + (row_h - circle_sz) // 2
            _circle(BODY_X, circle_y, circle_sz, BRAND_GREEN, f'NumC{i}',
                    [{'text': str(i+1), 'bold': True, 'color': WHITE, 'sz': 1400, 'font': 'Arial', 'align': 'ctr'}])
            # 文字
            paras = [{'text': text, 'bold': True, 'color': '333333', 'sz': 1400}]
            if desc:
                paras.append({'text': desc, 'bold': False, 'color': '555555', 'sz': 1100})
            textbox(BODY_X + emu(0.52), y, BODY_W - emu(0.52), row_h, f'NumTxt{i}', paras)
            y += row_h + emu(0.06)
        return y

    def _layout_matrix_2x2(data, y):
        """2×2 内容矩阵（dashi-ppt: matrix-2x2 角色）"""
        quadrants = data.get('quadrants', {})
        labels = data.get('labels', {'q1': '', 'q2': '', 'q3': '', 'q4': ''})
        fills = ['F0F7EE', 'FFF8EC', 'F5F0FA', 'FAF0F0']
        border_colors = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E']
        half_w = (BODY_W - emu(0.2)) // 2
        half_h = emu(1.6)
        for i, key in enumerate(['q1', 'q2', 'q3', 'q4']):
            col = i % 2
            row = i // 2
            cx = LEFT_X + col * (half_w + emu(0.2))
            cy = y + row * (half_h + emu(0.2))
            q_data = quadrants.get(key, {})
            title = q_data.get('title', labels.get(key, ''))
            items_list = q_data.get('items', [])
            # 卡片背景
            _rect(cx, cy, half_w, half_h, fills[i], f'Mat{i}', [], radius=10000,
                  line_color=border_colors[i], line_w=12700)
            # 顶部色条
            _rect(cx, cy, half_w, emu(0.08), border_colors[i], f'MatBar{i}', [])
            # 标题
            textbox(cx + emu(0.18), cy + emu(0.18), half_w - emu(0.36), emu(0.35), f'MatLbl{i}',
                    [{'text': title, 'bold': True, 'color': border_colors[i], 'sz': 1500}])
            # 内容
            if items_list:
                paras = [{'text': t, 'bold': False, 'color': '333333', 'sz': 1200, 'bullet': True} for t in items_list]
                textbox(cx + emu(0.18), cy + emu(0.55), half_w - emu(0.36), half_h - emu(0.7), f'MatBd{i}', paras)
        return y + 2 * half_h + emu(0.4)

    def _layout_chart_placeholder(data, y):
        """图表占位（dashi-ppt: chart/chart-area 角色）"""
        chart_type = data.get('chart_type', 'bar')
        caption = data.get('caption', '')
        note = data.get('note', '')
        # 虚线边框占位区域（用浅色矩形模拟）
        chart_h = emu(2.8)
        _rect(LEFT_X, y, BODY_W, chart_h, 'FAFAFA', 'ChartBg', [], radius=10000,
              line_color='CCCCCC', line_w=12700)
        # 图表类型标签
        type_labels = {'bar': '柱状图', 'line': '折线图', 'pie': '饼图', 'area': '面积图', 'scatter': '散点图'}
        type_cn = type_labels.get(chart_type, chart_type)
        textbox(LEFT_X, y + emu(1.0), BODY_W, emu(0.5), 'ChartLbl',
                [{'text': f'[{type_cn} 图表区域]', 'bold': False, 'color': '999999', 'sz': 1800, 'align': 'ctr'}])
        # 装饰性简单图形
        if chart_type == 'bar':
            bar_w = emu(0.5)
            heights = [emu(1.2), emu(1.8), emu(0.9), emu(2.0), emu(1.5)]
            bx = LEFT_X + (BODY_W - len(heights) * (bar_w + emu(0.2))) // 2
            for j, bh in enumerate(heights):
                color = PIE_COLORS[j % len(PIE_COLORS)]
                _rect(bx + j * (bar_w + emu(0.2)), y + chart_h - emu(0.3) - bh, bar_w, bh, color, f'CBar{j}', [], radius=5000)
        elif chart_type == 'line':
            # 用对角线连接数据点
            pts = [(0.0, 1.8), (0.2, 1.2), (0.4, 1.5), (0.6, 0.8), (0.8, 1.0), (1.0, 0.5)]
            line_w = BODY_W - emu(1.0)
            line_h = emu(1.5)
            lx = LEFT_X + emu(0.5)
            ly = y + emu(0.5)
            for j in range(len(pts)-1):
                x1 = int(lx + pts[j][0] * line_w)
                y1 = int(ly + (1-pts[j][1]) * line_h)
                x2 = int(lx + pts[j+1][0] * line_w)
                y2 = int(ly + (1-pts[j+1][1]) * line_h)
                dl = _diag_line(x1, y1, x2, y2, BRAND_GREEN, 25400)
                if dl is not None:
                    spTree.append(dl)
        # 标题
        textbox(LEFT_X, y + chart_h - emu(0.45), BODY_W, emu(0.35), 'ChartTitle',
                [{'text': caption, 'bold': False, 'color': '555555', 'sz': 1200, 'align': 'ctr'}])
        # 注释
        if note:
            textbox(LEFT_X, y + chart_h + emu(0.1), BODY_W, emu(0.3), 'ChartNote',
                    [{'text': note, 'bold': False, 'color': '999999', 'sz': 1100, 'align': 'ctr'}])
        return y + chart_h + (emu(0.5) if note else emu(0.2))

    def _layout_kpi_card(data, y):
        """KPI 指标卡（dashi-ppt: kpi-card/metric-spotlight 角色）"""
        cards = data.get('cards', [])
        if not cards:
            return y
        n = min(len(cards), 4)
        card_w = (BODY_W - (n - 1) * emu(0.2)) // n
        card_h = emu(2.0)
        for i, card in enumerate(cards[:4]):
            cx = LEFT_X + i * (card_w + emu(0.2))
            label = card.get('label', '')
            value = card.get('value', '0')
            unit = card.get('unit', '')
            target = card.get('target', '')
            trend = card.get('trend', '')   # 'up' / 'down' / 'flat'
            trend_val = card.get('trend_val', '')
            fill = CARD_COLORS[i % len(CARD_COLORS)]
            _rect(cx, y, card_w, card_h, fill, f'KpiBg{i}', [], radius=12000)
            # 标签
            textbox(cx + emu(0.15), y + emu(0.15), card_w - emu(0.3), emu(0.3), f'KpiLbl{i}',
                    [{'text': label, 'bold': False, 'color': WHITE, 'sz': 1300}])
            # 主数值
            textbox(cx + emu(0.15), y + emu(0.50), card_w - emu(0.3), emu(0.7), f'KpiVal{i}',
                    [{'text': f'{value}', 'bold': True, 'color': WHITE, 'sz': 3200, 'font': 'Arial'}])
            # 单位
            if unit:
                textbox(cx + emu(0.15), y + emu(1.15), card_w - emu(0.3), emu(0.3), f'KpiUnit{i}',
                        [{'text': unit, 'bold': False, 'color': WHITE, 'sz': 1300}])
            # 趋势
            if trend_val:
                arrow = '↑' if trend == 'up' else '↓' if trend == 'down' else '→'
                textbox(cx + emu(0.15), y + emu(1.50), card_w - emu(0.3), emu(0.3), f'KpiTr{i}',
                        [{'text': f'{arrow} {trend_val}', 'bold': True, 'color': GOLD, 'sz': 1400, 'font': 'Arial'}])
            elif target:
                textbox(cx + emu(0.15), y + emu(1.50), card_w - emu(0.3), emu(0.3), f'KpiTgt{i}',
                        [{'text': f'目标: {target}', 'bold': False, 'color': WHITE, 'sz': 1200}])
        return y + card_h + emu(0.2)

    def _layout_feature_grid(data, y):
        """特性网格（dashi-ppt: feature-grid/icon-list 角色）"""
        items = data.get('items', [])
        if not items:
            return y
        n = len(items)
        cols = 2 if n <= 4 else 3
        rows_n = (n + cols - 1) // cols
        gap = emu(0.18)
        card_w = (BODY_W - (cols - 1) * gap) // cols
        card_h = emu(1.35)
        ICON_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', '30C0B4', DARK_GREEN]
        for i, item in enumerate(items[:6]):
            col = i % cols
            row = i // cols
            cx = LEFT_X + col * (card_w + gap)
            cy = y + row * (card_h + gap)
            icon = item.get('icon', '')   # 文字图标符号
            title = item.get('title', '')
            desc = item.get('desc', '')
            color = ICON_COLORS[i % len(ICON_COLORS)]
            # 背景
            _rect(cx, cy, card_w, card_h, 'FFFFFF', f'FgBg{i}', [], radius=10000,
                  line_color='E8E8E8', line_w=6350)
            # 顶部色条
            _rect(cx, cy, card_w, emu(0.06), color, f'FgBar{i}', [])
            # 图标（文字符号或首字母）
            icon_text = icon if icon else str(i+1)
            _circle(cx + emu(0.12), cy + emu(0.2), emu(0.45), color, f'FgIc{i}',
                    [{'text': icon_text, 'bold': True, 'color': WHITE, 'sz': 1600, 'align': 'ctr'}])
            # 标题
            textbox(cx + emu(0.65), cy + emu(0.22), card_w - emu(0.8), emu(0.35), f'FgT{i}',
                    [{'text': title, 'bold': True, 'color': '333333', 'sz': 1500}])
            # 描述
            if desc:
                textbox(cx + emu(0.65), cy + emu(0.58), card_w - emu(0.8), card_h - emu(0.7), f'FgD{i}',
                        [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1200}])
        return y + rows_n * (card_h + gap)

    # ── 本轮新增 15 种布局（补齐 dashi-ppt 语义角色）─────────


    def _layout_radar_chart(data, y):
        """雷达图 / 能力图（dashi-ppt: radar 角色）"""
        raw_axes = data.get('axes', [])
        if not raw_axes:
            return y
        # 兼容两种格式: [{label,value}] 或 axes=[str] + values=[num]
        if isinstance(raw_axes[0], dict):
            labels = [a.get('label', '') for a in raw_axes]
            values = [a.get('value', 0) for a in raw_axes]
        else:
            labels = list(raw_axes)
            values = data.get('values', [0] * len(labels))
        n = min(len(labels), 8)
        if n < 3:
            return y
        chart_sz = emu(2.8)
        cx_c = LEFT_X + chart_sz // 2
        cy_c = y + chart_sz // 2
        radius = chart_sz // 2 - emu(0.15)
        # 画同心圆网格（3层）
        for ring in range(1, 4):
            r = radius * ring // 3
            _circle(cx_c - r, cy_c - r, r * 2, 'F5F5F5', f'RadarRing{ring}', [])
        # 画轴线 + 收集数据点
        max_val = max(values[:n]) or 1
        data_pts = []
        for i in range(n):
            angle = 2 * math.pi * i / n - math.pi / 2
            # 轴线（中心到边缘）
            ex = int(cx_c + radius * math.cos(angle))
            ey = int(cy_c + radius * math.sin(angle))
            ax = _diag_line(cx_c, cy_c, ex, ey, 'DDDDDD', 6350)
            if ax is not None:
                spTree.append(ax)
            # 数据点
            val = values[i] if i < len(values) else 0
            pt_r = int(radius * val / max_val)
            px = int(cx_c + pt_r * math.cos(angle))
            py = int(cy_c + pt_r * math.sin(angle))
            data_pts.append((px, py))
            dot_sz = emu(0.12)
            _circle(px - dot_sz // 2, py - dot_sz // 2, dot_sz, BRAND_GREEN, f'RDot{i}', [])
            # 轴标签
            lbl_r = radius + emu(0.25)
            lbl_x = int(cx_c + lbl_r * math.cos(angle))
            lbl_y = int(cy_c + lbl_r * math.sin(angle))
            textbox(lbl_x - emu(0.6), lbl_y - emu(0.15), emu(1.2), emu(0.3), f'RAxis{i}',
                    [{'text': labels[i], 'bold': False, 'color': '333333', 'sz': 1100, 'align': 'ctr'}])
        # 数据多边形连线
        for i in range(n):
            j = (i + 1) % n
            dl = _diag_line(data_pts[i][0], data_pts[i][1],
                            data_pts[j][0], data_pts[j][1], BRAND_GREEN, 15875)
            if dl is not None:
                spTree.append(dl)
        # 右侧图例
        legend_x = LEFT_X + chart_sz + emu(0.5)
        legend_y = y + emu(0.3)
        datasets = data.get('datasets', [{'label': '当前水平', 'color': BRAND_GREEN}])
        for i, ds in enumerate(datasets[:3]):
            color = ds.get('color', PIE_COLORS[i])
            label = ds.get('label', f'数据集{i+1}')
            _rect(legend_x, legend_y + i * emu(0.35), emu(0.25), emu(0.25), color, f'RLeg{i}', [], radius=20000)
            textbox(legend_x + emu(0.35), legend_y + i * emu(0.35) - emu(0.02), emu(3.0), emu(0.3), f'RLegT{i}',
                    [{'text': label, 'bold': False, 'color': '333333', 'sz': 1200}])
        return y + chart_sz + emu(0.2)

    def _layout_pyramid(data, y):
        """金字塔（dashi-ppt: pyramid 角色）"""
        levels = data.get('levels', [])
        if not levels:
            return y
        n = min(len(levels), 6)
        max_w = BODY_W
        level_h = emu(0.55)
        gap = emu(0.08)
        PYRAMID_COLORS = [BRAND_GREEN, '2D7A35', '3B8A52', GOLD, DARK_GREEN, 'D4940A']
        for i in range(n):
            w = max_w * (n - i) // n
            cx = (SLIDE_W - w) // 2
            cy = y + i * (level_h + gap)
            fill = PYRAMID_COLORS[i % len(PYRAMID_COLORS)]
            lvl = levels[i]
            label = lvl.get('label', f'层级{i+1}')
            value = lvl.get('value', '')
            _rect(cx, cy, w, level_h, fill, f'Pyr{i}', [], radius=8000, anchor='ctr')
            text = f'{label}  {value}' if value else label
            textbox(cx, cy, w, level_h, f'PyrT{i}',
                    [{'text': text, 'bold': True, 'color': WHITE, 'sz': 1400, 'align': 'ctr'}])
        return y + n * (level_h + gap)

    def _layout_roadmap(data, y):
        """路线图 / 阶段路线（dashi-ppt: roadmap 角色）"""
        phases = data.get('phases', [])
        if not phases:
            return y
        n = min(len(phases), 5)
        phase_w = (BODY_W - (n - 1) * emu(0.15)) // n
        card_h = emu(2.8)
        ROADMAP_COLORS = [BRAND_GREEN, GOLD, '2D7A35', '4874CB', DARK_GREEN]
        for i in range(n):
            cx = LEFT_X + i * (phase_w + emu(0.15))
            phase = phases[i]
            title = phase.get('title', f'阶段{i+1}')
            period = phase.get('period', '')
            desc = phase.get('desc', '')
            items = phase.get('items', [])
            color = ROADMAP_COLORS[i % len(ROADMAP_COLORS)]
            # 顶部圆形编号
            num_sz = emu(0.5)
            num_x = cx + (phase_w - num_sz) // 2
            _circle(num_x, y, num_sz, color, f'RoadNum{i}',
                    [{'text': str(i+1), 'bold': True, 'color': WHITE, 'sz': 1800, 'font': 'Arial', 'align': 'ctr'}])
            # 标题
            textbox(cx, y + num_sz + emu(0.08), phase_w, emu(0.35), f'RoadT{i}',
                    [{'text': title, 'bold': True, 'color': '333333', 'sz': 1400, 'align': 'ctr'}])
            # 时间段
            if period:
                textbox(cx, y + num_sz + emu(0.42), phase_w, emu(0.25), f'RoadP{i}',
                        [{'text': period, 'bold': False, 'color': GOLD, 'sz': 1100, 'align': 'ctr'}])
            # 内容卡片
            detail_y = y + num_sz + emu(0.72)
            _rect(cx, detail_y, phase_w, card_h - num_sz - emu(0.72), 'F8F8F8', f'RoadBg{i}', [],
                  radius=8000, line_color=color, line_w=9525)
            # 描述或 items
            if desc:
                textbox(cx + emu(0.08), detail_y + emu(0.08), phase_w - emu(0.16), card_h - num_sz - emu(0.88),
                        f'RoadD{i}', [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1100}])
            elif items:
                paras = [{'text': t, 'bold': False, 'color': '555555', 'sz': 1100, 'bullet': True} for t in items]
                textbox(cx + emu(0.08), detail_y + emu(0.08), phase_w - emu(0.16), card_h - num_sz - emu(0.88),
                        f'RoadI{i}', paras)
            # 连接线（到下一阶段）
            if i < n - 1:
                arrow_x = cx + phase_w
                arrow_y = y + num_sz // 2
                spTree.append(make_line(arrow_x, arrow_y, emu(0.15), GOLD, 19050))
        return y + card_h + emu(0.15)

    def _layout_venn(data, y):
        """维恩图 / 交集（dashi-ppt: venn 角色）"""
        groups = data.get('groups', [])
        if not groups:
            return y
        n = min(len(groups), 3)
        area_w = emu(5.0)
        area_h = emu(2.8)
        circle_sz = emu(2.2)
        # 3 圆圈品字排列
        positions = []
        if n == 2:
            positions = [(LEFT_X, y), (LEFT_X + emu(1.8), y)]
        elif n == 3:
            positions = [(LEFT_X, y + emu(0.3)), (LEFT_X + emu(1.8), y + emu(0.3)),
                         (LEFT_X + emu(0.9), y + emu(0.0))]
        else:
            positions = [(LEFT_X + emu(0.5), y)]
        VENN_COLORS = ['46A53B', 'FBB03B', '4874CB']
        for i in range(n):
            px, py = positions[i]
            color = VENN_COLORS[i % len(VENN_COLORS)]
            group = groups[i]
            label = group.get('label', f'集合{i+1}')
            items = group.get('items', [])
            # 半透明效果用浅色实填模拟
            light_colors = ['D5EDD3', 'FDE8B8', 'D0DCF5']
            _circle(px, py, circle_sz, light_colors[i % 3], f'Venn{i}', [])
            # 标签
            textbox(px, py + circle_sz // 2 - emu(0.3), circle_sz, emu(0.6), f'VennLbl{i}',
                    [{'text': label, 'bold': True, 'color': VENN_COLORS[i], 'sz': 1400, 'align': 'ctr'}])
        # 右侧说明
        legend_x = LEFT_X + emu(5.2)
        for i in range(n):
            group = groups[i]
            label = group.get('label', f'集合{i+1}')
            items = group.get('items', [])
            color = VENN_COLORS[i % len(VENN_COLORS)]
            ly = y + i * emu(1.0)
            _rect(legend_x, ly, emu(0.25), emu(0.25), color, f'VennLeg{i}', [], radius=20000)
            textbox(legend_x + emu(0.35), ly - emu(0.02), emu(4.0), emu(0.3), f'VennLegT{i}',
                    [{'text': label, 'bold': True, 'color': '333333', 'sz': 1300}])
            if items:
                textbox(legend_x + emu(0.35), ly + emu(0.28), emu(4.0), emu(0.6), f'VennLegD{i}',
                        [{'text': '、'.join(str(t) for t in items[:3]), 'bold': False, 'color': '555555', 'sz': 1100}])
        return y + area_h + emu(0.2)

    def _layout_ranking(data, y):
        """排行榜（dashi-ppt: ranking 角色）"""
        items = data.get('items', [])
        if not items:
            return y
        n = min(len(items), 6)
        row_h = emu(0.52)
        MEDAL_COLORS = [GOLD, 'C0C0C0', 'CD7F32']  # 金/银/铜
        for i in range(n):
            item = items[i]
            rank = item.get('rank', i + 1)
            name = item.get('name', '')
            value = item.get('value', '')
            desc = item.get('desc', '')
            # 背景交替（最先绘制，在奖牌/文字下方）
            if i % 2 == 0:
                _rect(BODY_X, y, BODY_W, row_h, 'F8FFF7', f'RankBg{i}', [], radius=20000)
            # 排名标记
            if i < 3:
                _circle(BODY_X, y + (row_h - emu(0.42)) // 2, emu(0.42), MEDAL_COLORS[i], f'Rank{i}',
                        [{'text': str(rank), 'bold': True, 'color': WHITE, 'sz': 1600, 'font': 'Arial', 'align': 'ctr'}])
            else:
                _rect(BODY_X, y + (row_h - emu(0.35)) // 2, emu(0.42), emu(0.35), 'E8E8E8', f'Rank{i}',
                      [{'text': str(rank), 'bold': True, 'color': '555555', 'sz': 1300, 'font': 'Arial', 'align': 'ctr'}],
                      radius=30000, anchor='ctr')
            # 名称
            textbox(BODY_X + emu(0.6), y, emu(4.0), row_h, f'RankN{i}',
                    [{'text': name, 'bold': True, 'color': '333333', 'sz': 1500}])
            # 数值
            textbox(BODY_X + emu(5.0), y, emu(3.0), row_h, f'RankV{i}',
                    [{'text': str(value), 'bold': True, 'color': BRAND_GREEN, 'sz': 1600, 'font': 'Arial'}])
            # 描述
            if desc:
                textbox(BODY_X + emu(8.0), y, emu(3.0), row_h, f'RankD{i}',
                        [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1200}])
            y += row_h + emu(0.06)
        return y

    def _layout_waterfall(data, y):
        """瀑布图（dashi-ppt: waterfall 角色）"""
        values = data.get('values', [])
        labels = data.get('labels', [])
        if not values:
            return y
        n = min(len(values), 8)
        # 计算累积值
        cum = 0
        bars = []
        for v in values[:n]:
            bars.append((cum, cum + v, v))
            cum += v
        max_h = max(abs(b[1]) for b in bars) or 1
        max_val = max(max_h, abs(min(b[0] for b in bars)))
        chart_h = emu(2.5)
        bar_w = (BODY_W - (n + 1) * emu(0.25)) // n
        # 基线
        base_y = y + chart_h
        spTree.append(make_line(LEFT_X, base_y, BODY_W, 'CCCCCC', 9525))
        for i, (start, end, val) in enumerate(bars):
            bx = LEFT_X + emu(0.25) + i * (bar_w + emu(0.25))
            # 柱子高度
            h = int(chart_h * abs(val) / max_val) if max_val > 0 else emu(0.3)
            h = max(h, emu(0.15))
            if val >= 0:
                by = int(base_y - chart_h * end / max_val)
                fill = BRAND_GREEN
            else:
                by = int(base_y - chart_h * start / max_val)
                fill = 'E54C5E'
            _rect(bx, by, bar_w, h, fill, f'WF{i}', [], radius=3000)
            # 数值标签
            textbox(bx, by - emu(0.28), bar_w, emu(0.25), f'WFV{i}',
                    [{'text': f'{val:+}' if isinstance(val, (int, float)) else str(val), 'bold': True,
                      'color': fill, 'sz': 1100, 'font': 'Arial', 'align': 'ctr'}])
            # 底部标签
            lbl = labels[i] if i < len(labels) else ''
            textbox(bx, base_y + emu(0.05), bar_w, emu(0.3), f'WFL{i}',
                    [{'text': lbl, 'bold': False, 'color': '555555', 'sz': 1000, 'align': 'ctr'}])
        return y + chart_h + emu(0.5)

    def _layout_heatmap(data, y):
        """热力矩阵（dashi-ppt: heatmap 角色）"""
        rows = data.get('rows', [])
        col_labels = data.get('col_labels', [])
        if not rows:
            return y
        n_rows = min(len(rows), 5)
        n_cols = max(len(col_labels), max(len(r.get('values', [])) for r in rows[:n_rows]))
        n_cols = min(n_cols, 8)
        label_w = emu(1.8)
        cell_w = (BODY_W - label_w) // n_cols
        cell_h = emu(0.52)
        # 列标签
        for j in range(n_cols):
            lbl = col_labels[j] if j < len(col_labels) else f'列{j+1}'
            textbox(LEFT_X + label_w + j * cell_w, y, cell_w, emu(0.35), f'HMCol{j}',
                    [{'text': lbl, 'bold': True, 'color': '555555', 'sz': 1100, 'align': 'ctr'}])
        # 行数据
        for i in range(n_rows):
            row = rows[i]
            row_label = row.get('label', f'行{i+1}')
            vals = row.get('values', [])
            ry = y + emu(0.4) + i * (cell_h + emu(0.06))
            # 行标签
            textbox(LEFT_X, ry, label_w, cell_h, f'HMRow{i}',
                    [{'text': row_label, 'bold': True, 'color': '333333', 'sz': 1200}])
            # 单元格
            for j in range(min(len(vals), n_cols)):
                v = vals[j]
                cx = LEFT_X + label_w + j * cell_w
                # 颜色深浅按数值（0-100 或原始值）
                try:
                    norm = min(max(float(v) / 100, 0), 1)
                except (ValueError, TypeError):
                    norm = 0.5
                # 从浅绿到深绿
                r_val = int(240 - norm * 170)
                g_val = int(247 - norm * 80)
                b_val = int(238 - norm * 180)
                color = f'{r_val:02X}{g_val:02X}{b_val:02X}'
                _rect(cx, ry, cell_w - emu(0.04), cell_h, color, f'HM{i}{j}', [], radius=4000)
                textbox(cx, ry, cell_w - emu(0.04), cell_h, f'HMT{i}{j}',
                        [{'text': str(v), 'bold': True, 'color': '333333' if norm < 0.5 else WHITE,
                          'sz': 1200, 'font': 'Arial', 'align': 'ctr'}])
        return y + emu(0.4) + n_rows * (cell_h + emu(0.06))

    def _layout_gantt(data, y):
        """甘特图 / 排期（dashi-ppt: gantt 角色）"""
        tasks = data.get('tasks', [])
        if not tasks:
            return y
        n = min(len(tasks), 6)
        periods = data.get('periods', [])  # 时间段标签
        n_periods = len(periods) if periods else 4
        label_w = emu(2.2)
        chart_w = BODY_W - label_w
        row_h = emu(0.5)
        period_w = chart_w // n_periods
        # 时间标签
        for j in range(n_periods):
            lbl = periods[j] if j < len(periods) else f'P{j+1}'
            textbox(LEFT_X + label_w + j * period_w, y, period_w, emu(0.3), f'GanttP{j}',
                    [{'text': lbl, 'bold': True, 'color': '555555', 'sz': 1100, 'align': 'ctr'}])
        # 任务条
        GANTT_COLORS = [BRAND_GREEN, GOLD, '4874CB', '2D7A35', DARK_GREEN, 'E54C5E']
        for i in range(n):
            task = tasks[i]
            name = task.get('name', f'任务{i+1}')
            start = task.get('start', 0)   # 从第几个 period 开始
            duration = task.get('duration', 1)  # 持续几个 period
            ry = y + emu(0.38) + i * (row_h + emu(0.08))
            # 任务名
            textbox(LEFT_X, ry, label_w, row_h, f'GanttN{i}',
                    [{'text': name, 'bold': False, 'color': '333333', 'sz': 1200}])
            # 背景条
            _rect(LEFT_X + label_w, ry + emu(0.08), chart_w, row_h - emu(0.16), 'F0F0F0', f'GanttBg{i}', [], radius=20000)
            # 进度条
            bar_x = LEFT_X + label_w + start * period_w
            bar_w = duration * period_w
            color = GANTT_COLORS[i % len(GANTT_COLORS)]
            _rect(bar_x, ry + emu(0.08), bar_w, row_h - emu(0.16), color, f'GanttBar{i}', [], radius=20000)
        return y + emu(0.38) + n * (row_h + emu(0.08))

    def _layout_cycle(data, y):
        """循环图 / 闭环（dashi-ppt: cycle/loop 角色）"""
        steps = data.get('steps', [])
        center = data.get('center', '')
        if not steps:
            return y
        n = min(len(steps), 6)
        area_sz = emu(3.5)
        cx_c = LEFT_X + (BODY_W - area_sz) // 2
        cy_c = y
        # 中心圆
        center_sz = emu(1.2)
        _circle(cx_c + (area_sz - center_sz) // 2, cy_c + (area_sz - center_sz) // 2,
                center_sz, BRAND_GREEN, 'CycleCenter',
                [{'text': center, 'bold': True, 'color': WHITE, 'sz': 1300, 'align': 'ctr'}])
        # 周围节点
        node_sz = emu(1.0)
        orbit_r = (area_sz - node_sz) // 2
        CYCLE_COLORS = [BRAND_GREEN, GOLD, '2D7A35', '4874CB', DARK_GREEN, 'E54C5E']
        for i in range(n):
            angle = 2 * math.pi * i / n - math.pi / 2
            nx = int(cx_c + area_sz // 2 + orbit_r * math.cos(angle) - node_sz // 2)
            ny = int(cy_c + area_sz // 2 + orbit_r * math.sin(angle) - node_sz // 2)
            color = CYCLE_COLORS[i % len(CYCLE_COLORS)]
            step = steps[i]
            label = step.get('label', '') if isinstance(step, dict) else str(step)
            _circle(nx, ny, node_sz, color, f'CycleN{i}',
                    [{'text': label, 'bold': True, 'color': WHITE, 'sz': 1100, 'align': 'ctr'}])
        return y + area_sz + emu(0.2)

    def _layout_big_number(data, y):
        """超大数字海报（dashi-ppt: bignumber/monolith 角色）"""
        number = data.get('number', '0')
        label = data.get('label', '')
        desc = data.get('desc', '')
        suffix = data.get('suffix', '')
        # 背景卡片
        box_h = emu(3.0)
        _rect(LEFT_X, y, BODY_W, box_h, DARK_GREEN, 'BigNumBg', [], radius=15000)
        # 金色装饰线
        spTree.append(make_line(LEFT_X + emu(0.5), y + emu(0.3), emu(1.0), GOLD, 25400))
        # 标签
        if label:
            textbox(LEFT_X + emu(0.5), y + emu(0.4), BODY_W - emu(1.0), emu(0.4), 'BNLabel',
                    [{'text': label, 'bold': True, 'color': GOLD, 'sz': 1600}])
        # 超大数字
        num_text = f'{number}{suffix}'
        textbox(LEFT_X, y + emu(0.8), BODY_W, emu(1.5), 'BNNumber',
                [{'text': num_text, 'bold': True, 'color': WHITE, 'sz': 7200, 'font': 'Arial', 'align': 'ctr'}])
        # 描述
        if desc:
            textbox(LEFT_X + emu(0.5), y + emu(2.3), BODY_W - emu(1.0), emu(0.5), 'BNDesc',
                    [{'text': desc, 'bold': False, 'color': WHITE, 'sz': 1600, 'align': 'ctr'}])
        return y + box_h + emu(0.2)

    def _layout_gallery(data, y):
        """作品/案例画廊（dashi-ppt: gallery/showcase 角色）"""
        items = data.get('items', [])
        if not items:
            return y
        n = min(len(items), 4)
        cols = 2 if n <= 2 else 4
        gap = emu(0.2)
        card_w = (BODY_W - (cols - 1) * gap) // cols
        card_h = emu(2.8)
        for i in range(n):
            col = i % cols
            row = i // cols
            cx = LEFT_X + col * (card_w + gap)
            cy = y + row * (card_h + gap)
            item = items[i]
            title = item.get('title', '')
            desc = item.get('desc', '')
            tag = item.get('tag', '')
            # 图片占位区
            img_h = emu(1.6)
            _rect(cx, cy, card_w, img_h, 'E8E8E8', f'GalImg{i}', [], radius=8000)
            textbox(cx, cy + img_h // 2 - emu(0.2), card_w, emu(0.4), f'GalPh{i}',
                    [{'text': '[图片]', 'bold': False, 'color': 'AAAAAA', 'sz': 1200, 'align': 'ctr'}])
            # 文字区
            text_y = cy + img_h + emu(0.1)
            if tag:
                _rect(cx, text_y, emu(0.8), emu(0.28), BRAND_GREEN, f'GalTag{i}',
                      [{'text': tag, 'bold': True, 'color': WHITE, 'sz': 1000, 'align': 'ctr'}],
                      radius=30000, anchor='ctr')
            textbox(cx, text_y + emu(0.32), card_w, emu(0.3), f'GalT{i}',
                    [{'text': title, 'bold': True, 'color': '333333', 'sz': 1400}])
            if desc:
                textbox(cx, text_y + emu(0.62), card_w, emu(0.6), f'GalD{i}',
                        [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1100}])
        return y + card_h + emu(0.15)

    def _layout_layers(data, y):
        """层级架构图（dashi-ppt: layers/stack 角色）"""
        layers = data.get('layers', [])
        if not layers:
            return y
        n = min(len(layers), 5)
        layer_h = emu(0.6)
        gap = emu(0.06)
        LAYER_COLORS = [BRAND_GREEN, '2D7A35', GOLD, '4874CB', DARK_GREEN]
        for i in range(n):
            layer = layers[i]
            label = layer.get('label', f'层级{i+1}')
            desc = layer.get('desc', '')
            color = LAYER_COLORS[i % len(LAYER_COLORS)]
            # 主条
            _rect(LEFT_X, y, BODY_W, layer_h, color, f'Layer{i}', [], radius=6000)
            textbox(LEFT_X + emu(0.2), y, emu(3.0), layer_h, f'LayerLbl{i}',
                    [{'text': label, 'bold': True, 'color': WHITE, 'sz': 1500}])
            if desc:
                textbox(LEFT_X + emu(3.5), y, BODY_W - emu(3.7), layer_h, f'LayerDesc{i}',
                        [{'text': desc, 'bold': False, 'color': WHITE, 'sz': 1200}])
            # 连接箭头（到下一层）
            if i < n - 1:
                arrow_x = LEFT_X + BODY_W // 2
                spTree.append(make_line(arrow_x, y + layer_h, emu(0.01), 'CCCCCC', 12700))
            y += layer_h + gap
        return y

    def _layout_bento(data, y):
        """便当格卡片（dashi-ppt: bento 角色）"""
        cells = data.get('cells', [])
        if not cells:
            return y
        # 便当格：2行3列，第一格占2列宽
        n = min(len(cells), 6)
        gap = emu(0.15)
        col_w = (BODY_W - 2 * gap) // 3
        row_h = emu(1.5)
        BENTO_COLORS = ['F0F7EE', 'FFF8EC', 'F5F0FA', 'FAF0F0', 'F0F7EE', 'FFF8EC']
        BORDER_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', BRAND_GREEN, GOLD]
        for i in range(n):
            # cell 0 占 row0 cols 0-1; cells 1-5 按 (i-1)%3, (i-1)//3 排列
            if i == 0:
                cw = 2 * col_w + gap
                cx = LEFT_X
                cy = y
            else:
                cw = col_w
                sub = i - 1  # 0-based index in remaining cells
                # cell 1 → (row0, col2), cells 2-4 → row1, cell 5 → row2
                if i == 1:
                    cx = LEFT_X + 2 * (col_w + gap)
                    cy = y
                else:
                    cx = LEFT_X + ((i - 2) % 3) * (col_w + gap)
                    cy = y + (1 + (i - 2) // 3) * (row_h + gap)
            cell = cells[i]
            title = cell.get('title', '')
            value = cell.get('value', '')
            desc = cell.get('desc', '')
            fill = BENTO_COLORS[i % len(BENTO_COLORS)]
            border = BORDER_COLORS[i % len(BORDER_COLORS)]
            _rect(cx, cy, cw, row_h, fill, f'Bento{i}', [], radius=10000,
                  line_color=border, line_w=6350)
            # 标题
            textbox(cx + emu(0.15), cy + emu(0.12), cw - emu(0.3), emu(0.3), f'BentoT{i}',
                    [{'text': title, 'bold': True, 'color': border, 'sz': 1300}])
            # 数值
            if value:
                textbox(cx + emu(0.15), cy + emu(0.45), cw - emu(0.3), emu(0.5), f'BentoV{i}',
                        [{'text': value, 'bold': True, 'color': '333333', 'sz': 2400, 'font': 'Arial'}])
            # 描述
            if desc:
                textbox(cx + emu(0.15), cy + emu(0.95), cw - emu(0.3), row_h - emu(1.05), f'BentoD{i}',
                        [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1100}])
        # 计算总行数
        if n <= 2:
            rows = 1  # cell 0 (row0) + cell 1 (row0 col2)
        elif n <= 5:
            rows = 2  # + cells 2-4 (row1)
        else:
            rows = 3  # + cell 5 (row2)
        return y + rows * row_h + (rows - 1) * gap

    def _layout_gauge(data, y):
        """仪表盘 / 达成率（dashi-ppt: gauge 角色）"""
        value = data.get('value', 0)
        label = data.get('label', '')
        target = data.get('target', 100)
        unit = data.get('unit', '')
        sub_metrics = data.get('sub_metrics', [])
        # 主仪表区域
        gauge_sz = emu(2.5)
        gauge_x = LEFT_X + (BODY_W - gauge_sz) // 2
        # 背景半圆（用圆形模拟，下半遮挡）
        _circle(gauge_x, y, gauge_sz, 'F0F0F0', 'GaugeBg', [])
        # 进度弧（用 pie 形状近似）
        try:
            pct = min(float(value) / float(target), 1.0) if target > 0 else 0
        except (ValueError, TypeError):
            pct = 0.5
        # 用不同颜色填充表示进度
        fill_color = BRAND_GREEN if pct >= 0.7 else GOLD if pct >= 0.4 else 'E54C5E'
        # 中心数值
        center_sz = emu(1.5)
        _circle(gauge_x + (gauge_sz - center_sz) // 2, y + (gauge_sz - center_sz) // 2,
                center_sz, WHITE, 'GaugeCenter', [])
        val_text = f'{value}{unit}'
        textbox(gauge_x + (gauge_sz - center_sz) // 2, y + (gauge_sz - center_sz) // 2 + emu(0.15),
                center_sz, emu(0.6), 'GaugeVal',
                [{'text': val_text, 'bold': True, 'color': fill_color, 'sz': 2800, 'font': 'Arial', 'align': 'ctr'}])
        textbox(gauge_x + (gauge_sz - center_sz) // 2, y + (gauge_sz - center_sz) // 2 + emu(0.75),
                center_sz, emu(0.3), 'GaugeLbl',
                [{'text': label, 'bold': False, 'color': '555555', 'sz': 1200, 'align': 'ctr'}])
        # 百分比
        textbox(gauge_x + (gauge_sz - center_sz) // 2, y + (gauge_sz - center_sz) // 2 + emu(1.05),
                center_sz, emu(0.25), 'GaugePct',
                [{'text': f'{pct*100:.0f}%', 'bold': True, 'color': fill_color, 'sz': 1400, 'font': 'Arial', 'align': 'ctr'}])
        gauge_bottom = y + gauge_sz + emu(0.2)
        # 副指标
        if sub_metrics:
            n_sub = min(len(sub_metrics), 3)
            sub_w = (BODY_W - (n_sub - 1) * emu(0.2)) // n_sub
            for i, sub in enumerate(sub_metrics[:3]):
                sx = LEFT_X + i * (sub_w + emu(0.2))
                _rect(sx, gauge_bottom, sub_w, emu(0.8), 'F8F8F8', f'GSubBg{i}', [], radius=8000)
                sv = sub.get('value', '')
                sl = sub.get('label', '')
                textbox(sx + emu(0.1), gauge_bottom + emu(0.08), sub_w - emu(0.2), emu(0.35), f'GSubV{i}',
                        [{'text': str(sv), 'bold': True, 'color': BRAND_GREEN, 'sz': 1600, 'font': 'Arial', 'align': 'ctr'}])
                textbox(sx + emu(0.1), gauge_bottom + emu(0.45), sub_w - emu(0.2), emu(0.25), f'GSubL{i}',
                        [{'text': sl, 'bold': False, 'color': '555555', 'sz': 1100, 'align': 'ctr'}])
            gauge_bottom += emu(1.0)
        return gauge_bottom

    def _layout_testimonial(data, y):
        """证言/引述卡（dashi-ppt: testimonial/voices 角色）"""
        quotes = data.get('quotes', [])
        if not quotes:
            # 单个引用模式
            quote = data.get('quote', '')
            author = data.get('author', '')
            role = data.get('role', '')
            if not quote:
                return y
            card_h = emu(2.8)
            _rect(LEFT_X, y, BODY_W, card_h, 'F0F7EE', 'TestiBg', [], radius=15000,
                  line_color=BRAND_GREEN, line_w=6350)
            # 引号装饰
            textbox(LEFT_X + emu(0.3), y + emu(0.1), emu(0.8), emu(0.8), 'TestiQ',
                    [{'text': '"', 'bold': True, 'color': GOLD, 'sz': 4800, 'font': 'Arial'}])
            # 引述文字
            textbox(LEFT_X + emu(0.5), y + emu(0.6), BODY_W - emu(1.0), emu(1.4), 'TestiText',
                    [{'text': quote, 'bold': False, 'color': '333333', 'sz': 1800}])
            # 作者
            if author:
                author_text = f'— {author}'
                if role:
                    author_text += f'，{role}'
                textbox(LEFT_X + emu(0.5), y + emu(2.1), BODY_W - emu(1.0), emu(0.4), 'TestiAuthor',
                        [{'text': author_text, 'bold': True, 'color': BRAND_GREEN, 'sz': 1400}])
            return y + card_h + emu(0.2)
        # 多个引用卡片模式
        n = min(len(quotes), 3)
        card_w = (BODY_W - (n - 1) * emu(0.2)) // n
        card_h = emu(2.8)
        TESTI_COLORS = ['F0F7EE', 'FFF8EC', 'F5F0FA']
        BORDER_COLORS = [BRAND_GREEN, GOLD, '4874CB']
        for i in range(n):
            q = quotes[i]
            cx = LEFT_X + i * (card_w + emu(0.2))
            quote_text = q.get('quote', '')
            author = q.get('author', '')
            role = q.get('role', '')
            fill = TESTI_COLORS[i % len(TESTI_COLORS)]
            border = BORDER_COLORS[i % len(BORDER_COLORS)]
            _rect(cx, y, card_w, card_h, fill, f'Testi{i}', [], radius=12000,
                  line_color=border, line_w=6350)
            # 引号
            textbox(cx + emu(0.15), y + emu(0.08), emu(0.5), emu(0.5), f'TestiQ{i}',
                    [{'text': '"', 'bold': True, 'color': GOLD, 'sz': 3200, 'font': 'Arial'}])
            # 引述
            textbox(cx + emu(0.15), y + emu(0.45), card_w - emu(0.3), emu(1.3), f'TestiT{i}',
                    [{'text': quote_text, 'bold': False, 'color': '333333', 'sz': 1300}])
            # 作者
            if author:
                author_text = f'— {author}'
                if role:
                    author_text += f'\n{role}'
                textbox(cx + emu(0.15), y + emu(1.85), card_w - emu(0.3), emu(0.8), f'TestiA{i}',
                        [{'text': author_text, 'bold': True, 'color': border, 'sz': 1200}])
        return y + card_h + emu(0.2)

    # ── 再次补齐 10 种（dashi-ppt 高频未覆盖模式）─────────

    def _layout_treemap(data, y):
        """矩形树图（dashi-ppt: treemap 角色，7 模板）"""
        items = data.get('items', [])
        if not items:
            return y
        n = min(len(items), 8)
        total = sum(it.get('value', 1) for it in items[:n]) or 1
        # 简化 treemap：按面积比例排列矩形，第一行大、下方小
        chart_h = emu(3.0)
        # 按值排序
        sorted_items = sorted(items[:n], key=lambda x: x.get('value', 0), reverse=True)
        # 分行：前几个大矩形占上行，其余占下行
        top_count = min(3, n)
        bot_count = n - top_count
        TREEMAP_COLORS = [BRAND_GREEN, GOLD, '2D7A35', 'D4940A', '4874CB', DARK_GREEN, 'E54C5E', '30C0B4']
        # 上行（按比例分配宽度）
        top_w = BODY_W
        top_total = sum(sorted_items[j].get('value', 1) for j in range(top_count)) or 1
        top_gap = emu(0.08)
        avail_w = top_w - (top_count - 1) * top_gap
        cx = LEFT_X
        for i in range(top_count):
            item = sorted_items[i]
            pct = item.get('value', 1) / top_total
            w = max(int(avail_w * pct), emu(0.5))
            color = TREEMAP_COLORS[i % len(TREEMAP_COLORS)]
            label = item.get('label', '')
            value = item.get('value', '')
            _rect(cx, y, w, chart_h // 2, color, f'TM{i}', [], radius=6000)
            textbox(cx + emu(0.1), y + emu(0.1), w - emu(0.2), chart_h // 2 - emu(0.2), f'TMT{i}',
                    [{'text': f'{label}\n{value}', 'bold': True, 'color': WHITE, 'sz': 1400, 'align': 'ctr'}])
            cx += w + top_gap
        # 下行
        if bot_count > 0:
            bot_w_each = (BODY_W - (bot_count - 1) * emu(0.08)) // bot_count
            for i in range(bot_count):
                idx = top_count + i
                item = sorted_items[idx]
                cx = LEFT_X + i * (bot_w_each + emu(0.08))
                color = TREEMAP_COLORS[idx % len(TREEMAP_COLORS)]
                label = item.get('label', '')
                value = item.get('value', '')
                _rect(cx, y + chart_h // 2 + emu(0.08), bot_w_each, chart_h // 2 - emu(0.08),
                      color, f'TM{idx}', [], radius=6000)
                textbox(cx + emu(0.08), y + chart_h // 2 + emu(0.16), bot_w_each - emu(0.16), chart_h // 2 - emu(0.24),
                        f'TMT{idx}', [{'text': f'{label}\n{value}', 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}])
        return y + chart_h + emu(0.15)

    def _layout_scatter(data, y):
        """散点/气泡图（dashi-ppt: scatter/bubble/dealmap 角色）"""
        points = data.get('points', [])
        x_label = data.get('x_label', 'X 轴')
        y_label = data.get('y_label', 'Y 轴')
        if not points:
            return y
        chart_w = BODY_W - emu(0.6)
        chart_h = emu(2.8)
        # 坐标轴
        ax_x = LEFT_X + emu(0.5)
        ax_y = y + chart_h
        spTree.append(make_line(ax_x, ax_y, chart_w, 'CCCCCC', 9525))  # X 轴
        spTree.append(make_line(ax_x, y, emu(0.01), 'CCCCCC', 9525))   # Y 轴（竖线用极窄矩形模拟）
        _rect(ax_x - emu(0.005), y, emu(0.01), chart_h, 'CCCCCC', 'YAxis', [], radius=0)
        # 轴标签
        textbox(ax_x + chart_w // 2, ax_y + emu(0.25), emu(2.0), emu(0.25), 'ScXLabel',
                [{'text': x_label, 'bold': False, 'color': '555555', 'sz': 1100, 'align': 'ctr'}])
        textbox(LEFT_X, y + chart_h // 2, emu(0.5), emu(0.3), 'ScYLabel',
                [{'text': y_label, 'bold': False, 'color': '555555', 'sz': 1100, 'align': 'ctr'}])
        # 画数据点
        xs = [p.get('x', 0) for p in points]
        ys = [p.get('y', 0) for p in points]
        max_x = max(xs) if xs else 1
        min_x = min(xs) if xs else 0
        max_y = max(ys) if ys else 1
        min_y = min(ys) if ys else 0
        range_x = max_x - min_x or 1
        range_y = max_y - min_y or 1
        SCATTER_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4']
        for i, pt in enumerate(points[:15]):
            px = pt.get('x', 0)
            py = pt.get('y', 0)
            size = pt.get('size', 0.3)
            label = pt.get('label', '')
            # 归一化坐标
            nx = (px - min_x) / range_x
            ny = (py - min_y) / range_y
            dot_x = int(ax_x + nx * chart_w - emu(size) // 2)
            dot_y = int(y + chart_h - ny * chart_h - emu(size) // 2)
            dot_sz = emu(size)
            dot_sz = max(dot_sz, emu(0.2))
            dot_sz = min(dot_sz, emu(0.8))
            color = SCATTER_COLORS[i % len(SCATTER_COLORS)]
            _circle(dot_x, dot_y, dot_sz, color, f'ScPt{i}',
                    [{'text': label[:3] if label else '', 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}])
        return y + chart_h + emu(0.5)

    def _layout_stacked_bar(data, y):
        """百分比堆叠柱状图（dashi-ppt: stacked/stacked-mix 角色）"""
        categories = data.get('categories', [])
        series = data.get('series', [])  # [{label, values[], color?}]
        if not categories or not series:
            return y
        n_cat = min(len(categories), 6)
        bar_h = emu(0.5)
        label_w = emu(1.8)
        bar_max_w = BODY_W - label_w - emu(1.5)
        STACK_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4']
        for i in range(n_cat):
            cat = categories[i]
            row_y = y + i * (bar_h + emu(0.2))
            # 类别标签
            textbox(LEFT_X, row_y, label_w, bar_h, f'StackLbl{i}',
                    [{'text': cat, 'bold': False, 'color': '333333', 'sz': 1300}])
            # 计算总和
            vals = []
            for s in series:
                v_list = s.get('values', [])
                vals.append(v_list[i] if i < len(v_list) else 0)
            total = sum(vals) or 1
            # 画堆叠段
            cx = LEFT_X + label_w
            for j, s in enumerate(series):
                v = vals[j]
                seg_w = max(int(bar_max_w * v / total), emu(0.1))
                color = s.get('color', STACK_COLORS[j % len(STACK_COLORS)])
                _rect(cx, row_y, seg_w, bar_h, color, f'Stack{i}{j}', [], radius=0)
                # 段内百分比
                if v / total > 0.12:
                    textbox(cx, row_y, seg_w, bar_h, f'StackT{i}{j}',
                            [{'text': f'{v/total*100:.0f}%', 'bold': True, 'color': WHITE, 'sz': 1100, 'align': 'ctr'}])
                cx += seg_w
        # 图例
        legend_y = y + n_cat * (bar_h + emu(0.2)) + emu(0.1)
        for j, s in enumerate(series):
            lx = LEFT_X + j * emu(2.5)
            color = s.get('color', STACK_COLORS[j % len(STACK_COLORS)])
            _rect(lx, legend_y, emu(0.2), emu(0.2), color, f'StackLeg{j}', [], radius=20000)
            textbox(lx + emu(0.28), legend_y - emu(0.02), emu(2.0), emu(0.25), f'StackLegT{j}',
                    [{'text': s.get('label', f'系列{j+1}'), 'bold': False, 'color': '333333', 'sz': 1100}])
        return legend_y + emu(0.3)

    def _layout_profile(data, y):
        """档案卡 / 公司概况（dashi-ppt: profile/dossier 角色，4 模板）"""
        name = data.get('name', '')
        subtitle = data.get('subtitle', '')
        metrics = data.get('metrics', [])  # [{label, value}]
        desc = data.get('desc', '')
        tags = data.get('tags', [])
        card_h = emu(3.2)
        _rect(LEFT_X, y, BODY_W, card_h, 'F8F8F8', 'ProfBg', [], radius=12000,
              line_color='E0E0E0', line_w=6350)
        # 左侧头像/Logo 占位
        avatar_sz = emu(1.2)
        _rect(LEFT_X + emu(0.3), y + emu(0.3), avatar_sz, avatar_sz, BRAND_GREEN, 'ProfAvatar',
              [{'text': name[:1] if name else '?', 'bold': True, 'color': WHITE, 'sz': 3600, 'align': 'ctr'}],
              radius=30000, anchor='ctr')
        # 名称
        info_x = LEFT_X + emu(1.8)
        textbox(info_x, y + emu(0.25), emu(5.0), emu(0.5), 'ProfName',
                [{'text': name, 'bold': True, 'color': '333333', 'sz': 2400}])
        if subtitle:
            textbox(info_x, y + emu(0.75), emu(5.0), emu(0.3), 'ProfSub',
                    [{'text': subtitle, 'bold': False, 'color': GOLD, 'sz': 1400}])
        # 标签
        if tags:
            tag_x = info_x
            for i, tag in enumerate(tags[:3]):
                tw = emu(1.0)
                _rect(tag_x + i * (tw + emu(0.1)), y + emu(1.1), tw, emu(0.28), BRAND_GREEN, f'ProfTag{i}',
                      [{'text': str(tag), 'bold': True, 'color': WHITE, 'sz': 1000, 'align': 'ctr'}],
                      radius=30000, anchor='ctr')
        # 指标行
        if metrics:
            met_y = y + emu(1.55)
            n_met = min(len(metrics), 4)
            met_w = (BODY_W - emu(2.1)) // n_met
            for i, met in enumerate(metrics[:4]):
                mx = info_x + i * met_w
                textbox(mx, met_y, met_w, emu(0.3), f'ProfML{i}',
                        [{'text': met.get('label', ''), 'bold': False, 'color': '555555', 'sz': 1100}])
                textbox(mx, met_y + emu(0.3), met_w, emu(0.35), f'ProfMV{i}',
                        [{'text': met.get('value', ''), 'bold': True, 'color': BRAND_GREEN, 'sz': 1800, 'font': 'Arial'}])
        # 描述
        if desc:
            textbox(LEFT_X + emu(0.3), y + emu(2.3), BODY_W - emu(0.6), emu(0.8), 'ProfDesc',
                    [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1200}])
        return y + card_h + emu(0.2)

    def _layout_spotlight(data, y):
        """聚焦特写 / 案例高亮（dashi-ppt: spotlight 角色，6 模板）"""
        title = data.get('spot_title', '')
        big_stat = data.get('big_stat', '')
        stat_label = data.get('stat_label', '')
        highlights = data.get('highlights', [])  # [{label, value}]
        desc = data.get('desc', '')
        card_h = emu(3.2)
        # 左侧大区域
        left_w = emu(6.0)
        _rect(LEFT_X, y, left_w, card_h, DARK_GREEN, 'SpotBg', [], radius=12000)
        # 大数字
        if big_stat:
            textbox(LEFT_X + emu(0.3), y + emu(0.2), left_w - emu(0.6), emu(1.2), 'SpotStat',
                    [{'text': big_stat, 'bold': True, 'color': GOLD, 'sz': 4800, 'font': 'Arial'}])
        if stat_label:
            textbox(LEFT_X + emu(0.3), y + emu(1.3), left_w - emu(0.6), emu(0.4), 'SpotLbl',
                    [{'text': stat_label, 'bold': False, 'color': WHITE, 'sz': 1600}])
        if desc:
            textbox(LEFT_X + emu(0.3), y + emu(1.8), left_w - emu(0.6), emu(1.2), 'SpotDesc',
                    [{'text': desc, 'bold': False, 'color': WHITE, 'sz': 1300}])
        # 右侧高亮列表
        rx = LEFT_X + left_w + emu(0.2)
        rw = BODY_W - left_w - emu(0.2)
        if title:
            textbox(rx, y, rw, emu(0.4), 'SpotTitle',
                    [{'text': title, 'bold': True, 'color': '333333', 'sz': 1800}])
        for i, hl in enumerate(highlights[:4]):
            hy = y + emu(0.5) + i * emu(0.65)
            _rect(rx, hy, rw, emu(0.55), 'F0F7EE', f'SpotH{i}', [], radius=8000)
            textbox(rx + emu(0.12), hy + emu(0.02), rw - emu(0.24), emu(0.25), f'SpotHL{i}',
                    [{'text': hl.get('label', ''), 'bold': False, 'color': '555555', 'sz': 1100}])
            textbox(rx + emu(0.12), hy + emu(0.27), rw - emu(0.24), emu(0.25), f'SpotHV{i}',
                    [{'text': hl.get('value', ''), 'bold': True, 'color': BRAND_GREEN, 'sz': 1600, 'font': 'Arial'}])
        return y + card_h + emu(0.2)

    def _layout_risk(data, y):
        """风险矩阵（dashi-ppt: risk 角色，6 模板）"""
        risks = data.get('risks', [])  # [{name, probability, impact, desc?}]
        if not risks:
            return y
        # 3×3 矩阵: probability(Y) × impact(X)
        grid_w = emu(3.0)
        grid_h = emu(0.8)
        ax_label_w = emu(1.2)
        # Y 轴标签
        textbox(LEFT_X, y, ax_label_w, emu(0.35), 'RiskYAxis',
                [{'text': '概率 →', 'bold': True, 'color': '555555', 'sz': 1200, 'align': 'ctr'}])
        # X 轴标签
        textbox(LEFT_X + ax_label_w + grid_w + emu(0.1), y + 3 * grid_h + emu(0.1),
                emu(3.0), emu(0.3), 'RiskXAxis',
                [{'text': '影响程度 →', 'bold': True, 'color': '555555', 'sz': 1200, 'align': 'ctr'}])
        # 风险等级颜色: 低=绿, 中=金, 高=红
        RISK_COLORS = {
            (0, 0): 'D5EDD3', (0, 1): 'D5EDD3', (0, 2): 'FFF8EC',
            (1, 0): 'D5EDD3', (1, 1): 'FFF8EC', (1, 2): 'FDE8B8',
            (2, 0): 'FFF8EC', (2, 1): 'FDE8B8', (2, 2): 'F8D0D0',
        }
        RISK_LABELS = ['低', '低', '中', '低', '中', '中', '中', '中', '高']
        # 画 3×3 网格 (impact 0-2 为列, probability 0-2 为行)
        for row in range(3):
            for col in range(3):
                cx = LEFT_X + ax_label_w + col * grid_w
                cy = y + emu(0.4) + (2 - row) * grid_h  # 翻转 Y 轴
                idx = row * 3 + col
                color = RISK_COLORS.get((row, col), 'F5F5F5')
                _rect(cx, cy, grid_w - emu(0.04), grid_h - emu(0.04), color, f'RiskCell{row}{col}', [],
                      radius=4000, line_color='E0E0E0', line_w=6350)
                # 该格中的风险项
                items_in_cell = [r for r in risks if r.get('probability', 0) == row and r.get('impact', 0) == col]
                if items_in_cell:
                    text = '\n'.join(r.get('name', '') for r in items_in_cell[:2])
                    textbox(cx + emu(0.08), cy + emu(0.04), grid_w - emu(0.16), grid_h - emu(0.08),
                            f'RiskT{row}{col}',
                            [{'text': text, 'bold': True, 'color': '333333', 'sz': 1100}])
        # 右侧风险列表
        list_x = LEFT_X + ax_label_w + 3 * grid_w + emu(0.3)
        for i, risk in enumerate(risks[:5]):
            ry = y + emu(0.4) + i * emu(0.52)
            name = risk.get('name', '')
            desc = risk.get('desc', '')
            prob = risk.get('probability', 0)
            imp = risk.get('impact', 0)
            level = '高' if prob + imp >= 3 else '中' if prob + imp >= 2 else '低'
            level_color = 'E54C5E' if level == '高' else GOLD if level == '中' else BRAND_GREEN
            _rect(list_x, ry, emu(0.4), emu(0.4), level_color, f'RiskLvl{i}',
                  [{'text': level, 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}],
                  radius=30000, anchor='ctr')
            textbox(list_x + emu(0.5), ry, emu(4.0), emu(0.4), f'RiskName{i}',
                    [{'text': f'{name}  {desc}', 'bold': False, 'color': '333333', 'sz': 1200}])
        return y + emu(0.4) + 3 * grid_h + emu(0.3)

    def _layout_swimlane(data, y):
        """泳道流程（dashi-ppt: swimlane 角色）"""
        lanes = data.get('lanes', [])  # [{label, steps[{text}]}]
        if not lanes:
            return y
        n_lanes = min(len(lanes), 4)
        label_w = emu(1.5)
        lane_h = emu(0.7)
        chart_w = BODY_W - label_w
        SWIM_COLORS = [BRAND_GREEN, GOLD, '4874CB', DARK_GREEN]
        for i in range(n_lanes):
            lane = lanes[i]
            ly = y + i * (lane_h + emu(0.1))
            color = SWIM_COLORS[i % len(SWIM_COLORS)]
            # 泳道标签
            _rect(LEFT_X, ly, label_w, lane_h, color, f'SwimLbl{i}',
                  [{'text': lane.get('label', f'角色{i+1}'), 'bold': True, 'color': WHITE, 'sz': 1300, 'align': 'ctr'}],
                  radius=6000, anchor='ctr')
            # 泳道背景
            _rect(LEFT_X + label_w + emu(0.08), ly, chart_w - emu(0.08), lane_h,
                  'F8F8F8', f'SwimBg{i}', [], radius=4000, line_color='E8E8E8', line_w=6350)
            # 步骤
            steps = lane.get('steps', [])
            if steps:
                n_steps = min(len(steps), 4)
                step_w = (chart_w - emu(0.3)) // n_steps
                for j, step in enumerate(steps[:4]):
                    sx = LEFT_X + label_w + emu(0.15) + j * step_w
                    text = step.get('text', '') if isinstance(step, dict) else str(step)
                    _rect(sx, ly + emu(0.1), step_w - emu(0.1), lane_h - emu(0.2), color, f'SwimS{i}{j}',
                          [{'text': text, 'bold': False, 'color': WHITE, 'sz': 1100, 'align': 'ctr'}],
                          radius=6000, anchor='ctr')
        return y + n_lanes * (lane_h + emu(0.1))

    def _layout_overview(data, y):
        """摘要概览页（dashi-ppt: overview/spec/summary/recap 角色）"""
        summary = data.get('summary', '')
        key_points = data.get('key_points', [])  # [{label, value, desc?}]
        # 顶部摘要
        if summary:
            _rect(LEFT_X, y, BODY_W, emu(0.8), 'F0F7EE', 'OvBg', [], radius=8000)
            textbox(LEFT_X + emu(0.2), y + emu(0.1), BODY_W - emu(0.4), emu(0.6), 'OvSummary',
                    [{'text': summary, 'bold': False, 'color': '333333', 'sz': 1500}])
            y += emu(1.0)
        # 关键指标卡
        if key_points:
            n = min(len(key_points), 4)
            card_w = (BODY_W - (n - 1) * emu(0.15)) // n
            card_h = emu(2.2)
            OV_COLORS = [BRAND_GREEN, GOLD, '4874CB', DARK_GREEN]
            for i in range(n):
                kp = key_points[i]
                cx = LEFT_X + i * (card_w + emu(0.15))
                color = OV_COLORS[i % len(OV_COLORS)]
                _rect(cx, y, card_w, card_h, color, f'OvCard{i}', [], radius=10000)
                textbox(cx + emu(0.15), y + emu(0.15), card_w - emu(0.3), emu(0.3), f'OvLbl{i}',
                        [{'text': kp.get('label', ''), 'bold': False, 'color': WHITE, 'sz': 1300}])
                textbox(cx + emu(0.15), y + emu(0.5), card_w - emu(0.3), emu(0.6), f'OvVal{i}',
                        [{'text': kp.get('value', ''), 'bold': True, 'color': WHITE, 'sz': 2800, 'font': 'Arial'}])
                desc = kp.get('desc', '')
                if desc:
                    textbox(cx + emu(0.15), y + emu(1.2), card_w - emu(0.3), emu(0.8), f'OvDesc{i}',
                            [{'text': desc, 'bold': False, 'color': WHITE, 'sz': 1200}])
            y += card_h + emu(0.15)
        return y

    def _layout_principles(data, y):
        """核心原则 / 价值观（dashi-ppt: principles 角色，2 模板）"""
        items = data.get('items', [])
        if not items:
            return y
        n = min(len(items), 5)
        PRIN_COLORS = [BRAND_GREEN, GOLD, '2D7A35', '4874CB', DARK_GREEN]
        # 固定行高，5行合计约 3.0in，适配内容区
        row_h = emu(0.5)
        for i in range(n):
            item = items[i]
            title = item.get('title', '') if isinstance(item, dict) else str(item)
            desc = item.get('desc', '') if isinstance(item, dict) else ''
            color = PRIN_COLORS[i % len(PRIN_COLORS)]
            # 编号圆圈
            num_sz = emu(0.38)
            _circle(BODY_X, y + (row_h - num_sz) // 2, num_sz, color, f'PrinN{i}',
                    [{'text': str(i+1), 'bold': True, 'color': WHITE, 'sz': 1400, 'font': 'Arial', 'align': 'ctr'}])
            # 标题
            textbox(BODY_X + emu(0.55), y, emu(4.5), row_h, f'PrinT{i}',
                    [{'text': title, 'bold': True, 'color': '333333', 'sz': 1500}])
            # 描述
            if desc:
                textbox(BODY_X + emu(5.5), y, BODY_W - emu(5.5), row_h, f'PrinD{i}',
                        [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1100}])
            y += row_h + emu(0.1)
        return y

    def _layout_org_chart(data, y):
        """组织/生态架构图（dashi-ppt: orgchart/ecosystem/nexus 角色）"""
        nodes = data.get('nodes', [])  # [{label, level, desc?}]
        if not nodes:
            return y
        # 按 level 分组
        from collections import defaultdict
        levels = defaultdict(list)
        for node in nodes:
            lvl = node.get('level', 0)
            levels[lvl].append(node)
        lvl_keys = sorted(levels.keys())
        n_levels = min(len(lvl_keys), 4)
        ORG_COLORS = [BRAND_GREEN, GOLD, '4874CB', DARK_GREEN]
        # 固定布局：4级节点合计约 3.2in
        node_h = emu(0.6)
        level_gap = emu(0.65)
        for li in range(n_levels):
            lvl = lvl_keys[li]
            nodes_at_lvl = levels[lvl][:5]  # 最多 5 个/层
            n_nodes = len(nodes_at_lvl)
            color = ORG_COLORS[li % len(ORG_COLORS)]
            node_w = (BODY_W - (n_nodes - 1) * emu(0.15)) // n_nodes
            node_w = min(node_w, emu(3.5))
            total_w = n_nodes * node_w + (n_nodes - 1) * emu(0.15)
            start_x = LEFT_X + (BODY_W - total_w) // 2
            ly = y + li * (node_h + level_gap)
            for ni, node in enumerate(nodes_at_lvl):
                nx = start_x + ni * (node_w + emu(0.15))
                label = node.get('label', '')
                desc = node.get('desc', '')
                _rect(nx, ly, node_w, node_h, color, f'Org{li}{ni}', [], radius=8000)
                text = f'{label}\n{desc}' if desc else label
                textbox(nx, ly, node_w, node_h, f'OrgT{li}{ni}',
                        [{'text': text, 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}])
        return y + n_levels * (node_h + level_gap) - level_gap

    # ── 补齐 dashi-ppt 高频未覆盖图表（11 种）────────

    def _layout_bump(data, y):
        """名次变迁图（dashi-ppt: bump/bump-rank，4模板）"""
        series = data.get('series', [])  # [{label, ranks[]/values[]}]
        periods = data.get('periods', [])  # ['Q1','Q2',...]
        if not series:
            return y
        n_series = min(len(series), 5)
        # 兼容 ranks 和 values 两种 key
        def _get_ranks(s):
            return s.get('ranks', s.get('values', []))
        n_periods = max(len(_get_ranks(s)) for s in series) if series else 0
        if n_periods < 2:
            return y
        chart_w = BODY_W - emu(1.5)
        chart_h = emu(2.8)
        BUMPCOLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN]
        # 画水平排名线
        for r in range(n_series):
            rank_y = y + int(chart_h * r / max(n_series - 1, 1))
            _rect(LEFT_X + emu(0.5), rank_y, chart_w, emu(0.01), 'E8E8E8', f'BumpLine{r}', [], radius=0)
            textbox(LEFT_X, rank_y - emu(0.1), emu(0.5), emu(0.25), f'BumpR{r}',
                    [{'text': f'#{r+1}', 'bold': True, 'color': '999999', 'sz': 1000, 'align': 'r'}])
        # 画数据点和连线
        for si, s in enumerate(series[:n_series]):
            ranks = _get_ranks(s)
            color = BUMPCOLORS[si % len(BUMPCOLORS)]
            label = s.get('label', '')
            pts = []
            for pi in range(min(len(ranks), n_periods)):
                rank = ranks[pi] - 1  # 0-indexed
                rank = max(0, min(rank, n_series - 1))
                cx = LEFT_X + emu(0.5) + int(chart_w * pi / max(n_periods - 1, 1))
                cy = y + int(chart_h * rank / max(n_series - 1, 1))
                pts.append((cx, cy))
                dot_sz = emu(0.3)
                _circle(cx - dot_sz // 2, cy - dot_sz // 2, dot_sz, color, f'BumpD{si}{pi}',
                        [{'text': str(ranks[pi]), 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}])
            # 用对角线连接数据点
            for pi in range(1, len(pts)):
                dl = _diag_line(pts[pi-1][0], pts[pi-1][1], pts[pi][0], pts[pi][1], color, 15875)
                if dl is not None:
                    spTree.append(dl)
            # 末端标签
            if pts:
                textbox(pts[-1][0] + emu(0.2), pts[-1][1] - emu(0.1), emu(1.5), emu(0.25), f'BumpL{si}',
                        [{'text': label, 'bold': True, 'color': color, 'sz': 1100}])
        # 时间段标签
        for pi in range(n_periods):
            px = LEFT_X + emu(0.5) + int(chart_w * pi / max(n_periods - 1, 1))
            period_label = periods[pi] if pi < len(periods) else f'P{pi+1}'
            textbox(px - emu(0.3), y + chart_h + emu(0.1), emu(0.6), emu(0.25), f'BumpP{pi}',
                    [{'text': period_label, 'bold': False, 'color': '999999', 'sz': 1000, 'align': 'ctr'}])
        return y + chart_h + emu(0.4)

    def _layout_dumbbell(data, y):
        """哑铃图（dashi-ppt: dumbbell，4模板）"""
        items = data.get('items', [])  # [{label, before, after, unit?}]
        if not items:
            return y
        n = min(len(items), 6)
        label_w = emu(1.5)
        chart_w = BODY_W - label_w - emu(1.5)
        row_h = emu(0.45)
        all_vals = []
        for it in items[:n]:
            all_vals.extend([it.get('before', 0), it.get('after', 0)])
        max_val = max(all_vals) if all_vals else 1
        DUMBCOLORS = ['before', 'after']
        for i in range(n):
            it = items[i]
            ry = y + i * (row_h + emu(0.15))
            label = it.get('label', '')
            before = it.get('before', 0)
            after = it.get('after', 0)
            unit = it.get('unit', '')
            textbox(LEFT_X, ry, label_w, row_h, f'DumbLbl{i}',
                    [{'text': label, 'bold': False, 'color': '333333', 'sz': 1200, 'align': 'r'}])
            bx = LEFT_X + label_w + int(chart_w * before / max_val)
            ax = LEFT_X + label_w + int(chart_w * after / max_val)
            # 连接线
            conn_x = min(bx, ax)
            conn_w = abs(ax - bx)
            if conn_w > emu(0.05):
                _rect(conn_x, ry + row_h // 2 - emu(0.02), conn_w, emu(0.04), 'CCCCCC', f'DumbC{i}', [], radius=0)
            # 左点 (before)
            dot_sz = emu(0.32)
            _circle(bx - dot_sz // 2, ry + row_h // 2 - dot_sz // 2, dot_sz, 'E54C5E', f'DumbB{i}',
                    [{'text': str(before), 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}])
            # 右点 (after)
            _circle(ax - dot_sz // 2, ry + row_h // 2 - dot_sz // 2, dot_sz, BRAND_GREEN, f'DumbA{i}',
                    [{'text': str(after), 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}])
        # 图例
        legend_y = y + n * (row_h + emu(0.15)) + emu(0.05)
        _rect(LEFT_X + label_w, legend_y, emu(0.15), emu(0.15), 'E54C5E', 'DumbLegB', [], radius=30000)
        textbox(LEFT_X + label_w + emu(0.2), legend_y - emu(0.02), emu(1.0), emu(0.2), 'DumbLegBT',
                [{'text': '初始', 'bold': False, 'color': '555555', 'sz': 1000}])
        _rect(LEFT_X + label_w + emu(1.5), legend_y, emu(0.15), emu(0.15), BRAND_GREEN, 'DumbLegA', [], radius=30000)
        textbox(LEFT_X + label_w + emu(1.7), legend_y - emu(0.02), emu(1.0), emu(0.2), 'DumbLegAT',
                [{'text': '最终', 'bold': False, 'color': '555555', 'sz': 1000}])
        return legend_y + emu(0.25)

    def _layout_lollipop(data, y):
        """棒棒糖图（dashi-ppt: lollipop，1模板）"""
        items = data.get('items', [])  # [{label, value}]
        if not items:
            return y
        n = min(len(items), 8)
        label_w = emu(1.5)
        chart_w = BODY_W - label_w - emu(1.0)
        max_val = max(it.get('value', 0) for it in items[:n]) or 1
        LOLLIPOP_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4', '2D7A35', 'D4940A']
        for i in range(n):
            it = items[i]
            ry = y + i * emu(0.38)
            label = it.get('label', '')
            value = it.get('value', 0)
            color = LOLLIPOP_COLORS[i % len(LOLLIPOP_COLORS)]
            textbox(LEFT_X, ry, label_w, emu(0.35), f'LolLbl{i}',
                    [{'text': label, 'bold': False, 'color': '333333', 'sz': 1200, 'align': 'r'}])
            bar_w = max(int(chart_w * value / max_val), emu(0.3))
            # 棍子（细矩形）
            stick_h = emu(0.04)
            _rect(LEFT_X + label_w, ry + emu(0.15), bar_w, stick_h, color, f'LolStick{i}', [], radius=0)
            # 糖果（圆点）
            dot_sz = emu(0.28)
            _circle(LEFT_X + label_w + bar_w - dot_sz // 2, ry + emu(0.17) - dot_sz // 2 + emu(0.02),
                    dot_sz, color, f'LolDot{i}',
                    [{'text': str(value), 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}])
        return y + n * emu(0.38) + emu(0.1)

    def _layout_waffle(data, y):
        """华夫饼图 / 百分比方格（dashi-ppt: waffle/dotplot/dotfield，5模板）"""
        items = data.get('items', [])  # [{label, value, color?}]
        total = data.get('total', 100)
        if not items:
            return y
        # 10×10 方格
        grid_sz = emu(0.22)
        gap = emu(0.03)
        grid_w = 10 * (grid_sz + gap)
        WAF_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4']
        # 计算每个方格归属
        cells = []
        filled = 0
        for ii, it in enumerate(items):
            count = int(round(it.get('value', 0) / total * 100))
            color = it.get('color', WAF_COLORS[ii % len(WAF_COLORS)])
            for _ in range(count):
                if filled >= 100:
                    break
                cells.append(color)
                filled += 1
        while len(cells) < 100:
            cells.append('E8E8E8')
        # 画方格
        for idx in range(100):
            row = idx // 10
            col = idx % 10
            cx = BODY_X + col * (grid_sz + gap)
            cy = y + row * (grid_sz + gap)
            color = cells[idx]
            _rect(cx, cy, grid_sz, grid_sz, color, f'Waff{idx}', [], radius=3000)
        # 图例
        legend_y = y + 10 * (grid_sz + gap) + emu(0.15)
        for ii, it in enumerate(items[:5]):
            lx = BODY_X + ii * emu(2.0)
            color = it.get('color', WAF_COLORS[ii % len(WAF_COLORS)])
            _rect(lx, legend_y, emu(0.15), emu(0.15), color, f'WaffLeg{ii}', [], radius=3000)
            textbox(lx + emu(0.2), legend_y - emu(0.02), emu(1.5), emu(0.2), f'WaffLegT{ii}',
                    [{'text': f"{it.get('label', '')} {it.get('value', '')}", 'bold': False, 'color': '333333', 'sz': 1100}])
        return legend_y + emu(0.25)

    def _layout_radial_bar(data, y):
        """径向条形图（dashi-ppt: radialbar，1模板）"""
        items = data.get('items', [])  # [{label, value, max?}]
        if not items:
            return y
        n = min(len(items), 8)
        chart_sz = emu(2.8)
        cx_c = LEFT_X + chart_sz // 2
        cy_c = y + chart_sz // 2
        max_r = chart_sz // 2 - emu(0.2)
        RADIAL_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4', '2D7A35', 'D4940A']
        max_val = max((it.get('value', 0) for it in items[:n]), default=1) or 1
        # 背景圆
        bg_r = max_r
        _circle(cx_c - bg_r, cy_c - bg_r, bg_r * 2, 'F8F8F8', 'RadBg', [])
        # 各方向条形
        for i in range(n):
            it = items[i]
            label = it.get('label', '')
            value = it.get('value', 0)
            color = RADIAL_COLORS[i % len(RADIAL_COLORS)]
            angle = 2 * math.pi * i / n - math.pi / 2
            bar_len = int(max_r * value / max_val)
            bar_len = max(bar_len, emu(0.15))
            # 条形终点
            ex = int(cx_c + bar_len * math.cos(angle))
            ey = int(cy_c + bar_len * math.sin(angle))
            # 画条形（粗对角线）
            dl = _diag_line(cx_c, cy_c, ex, ey, color, emu(0.12))
            if dl is not None:
                spTree.append(dl)
            # 端点圆点
            dot_sz = emu(0.1)
            _circle(ex - dot_sz // 2, ey - dot_sz // 2, dot_sz, color, f'RadDot{i}', [])
            # 标签
            lbl_r = max_r + emu(0.2)
            lbl_x = int(cx_c + lbl_r * math.cos(angle))
            lbl_y = int(cy_c + lbl_r * math.sin(angle))
            textbox(lbl_x - emu(0.6), lbl_y - emu(0.15), emu(1.2), emu(0.3), f'RadLbl{i}',
                    [{'text': f'{label} {value}', 'bold': False, 'color': '333333', 'sz': 1000, 'align': 'ctr'}])
        # 中心圆
        c_sz = emu(0.3)
        _circle(cx_c - c_sz // 2, cy_c - c_sz // 2, c_sz, WHITE, 'RadCtr', [])
        return y + chart_sz + emu(0.2)

    def _layout_diverging(data, y):
        """正负双向条形图（dashi-ppt: diverging，2模板）"""
        items = data.get('items', [])  # [{label, value, unit?}]
        if not items:
            return y
        n = min(len(items), 8)
        label_w = emu(1.5)
        chart_w = (BODY_W - label_w - emu(1.0)) // 2  # 半边宽度
        center_x = LEFT_X + label_w + chart_w
        max_abs = max(abs(it.get('value', 0)) for it in items[:n]) or 1
        row_h = emu(0.35)
        for i in range(n):
            it = items[i]
            ry = y + i * (row_h + emu(0.08))
            label = it.get('label', '')
            value = it.get('value', 0)
            textbox(LEFT_X, ry, label_w, row_h, f'DivLbl{i}',
                    [{'text': label, 'bold': False, 'color': '333333', 'sz': 1200, 'align': 'r'}])
            bar_w = max(int(chart_w * abs(value) / max_abs), emu(0.2))
            color = BRAND_GREEN if value >= 0 else 'E54C5E'
            if value >= 0:
                _rect(center_x + emu(0.05), ry, bar_w, row_h, color, f'DivBar{i}', [], radius=4000)
                textbox(center_x + bar_w + emu(0.12), ry, emu(0.8), row_h, f'DivVal{i}',
                        [{'text': str(value), 'bold': True, 'color': color, 'sz': 1100}])
            else:
                _rect(center_x - emu(0.05) - bar_w, ry, bar_w, row_h, color, f'DivBar{i}', [], radius=4000)
                textbox(center_x - bar_w - emu(0.8), ry, emu(0.75), row_h, f'DivVal{i}',
                        [{'text': str(value), 'bold': True, 'color': color, 'sz': 1100, 'align': 'r'}])
        # 中线
        spTree.append(make_line(center_x, y, emu(0.01), 'CCCCCC', 9525))
        _rect(center_x - emu(0.005), y, emu(0.01), n * (row_h + emu(0.08)), 'CCCCCC', 'DivCenter', [], radius=0)
        return y + n * (row_h + emu(0.08)) + emu(0.1)

    def _layout_tornado(data, y):
        """龙卷风图 / 背对背条形图（dashi-ppt: tornado，3模板）"""
        items = data.get('items', [])  # [{label, left_value, right_value, left_label?, right_label?}]
        left_name = data.get('left_label', '左侧')
        right_name = data.get('right_label', '右侧')
        if not items:
            return y
        n = min(len(items), 7)
        center_label_w = emu(1.8)
        side_w = (BODY_W - center_label_w) // 2
        max_val = max(max(it.get('left_value', 0), it.get('right_value', 0)) for it in items[:n]) or 1
        row_h = emu(0.38)
        center_x = LEFT_X + side_w
        for i in range(n):
            it = items[i]
            ry = y + i * (row_h + emu(0.1))
            label = it.get('label', '')
            lv = it.get('left_value', 0)
            rv = it.get('right_value', 0)
            # 中间标签
            textbox(center_x, ry, center_label_w, row_h, f'TorLbl{i}',
                    [{'text': label, 'bold': False, 'color': '333333', 'sz': 1200, 'align': 'ctr'}])
            # 左条
            lb_w = max(int(side_w * lv / max_val), emu(0.15))
            _rect(center_x - emu(0.05) - lb_w, ry, lb_w, row_h, 'E54C5E', f'TorL{i}', [], radius=4000)
            textbox(center_x - emu(0.05) - lb_w + emu(0.05), ry, lb_w, row_h, f'TorLV{i}',
                    [{'text': str(lv), 'bold': True, 'color': WHITE, 'sz': 1000, 'align': 'r'}])
            # 右条
            rb_w = max(int(side_w * rv / max_val), emu(0.15))
            _rect(center_x + center_label_w + emu(0.05), ry, rb_w, row_h, BRAND_GREEN, f'TorR{i}', [], radius=4000)
            textbox(center_x + center_label_w + emu(0.1), ry, rb_w, row_h, f'TorRV{i}',
                    [{'text': str(rv), 'bold': True, 'color': WHITE, 'sz': 1000}])
        # 图例
        legend_y = y + n * (row_h + emu(0.1)) + emu(0.05)
        _rect(LEFT_X, legend_y, emu(0.15), emu(0.15), 'E54C5E', 'TorLegL', [], radius=30000)
        textbox(LEFT_X + emu(0.2), legend_y - emu(0.02), emu(1.5), emu(0.2), 'TorLegLT',
                [{'text': left_name, 'bold': False, 'color': '555555', 'sz': 1000}])
        _rect(LEFT_X + BODY_W - emu(1.5), legend_y, emu(0.15), emu(0.15), BRAND_GREEN, 'TorLegR', [], radius=30000)
        textbox(LEFT_X + BODY_W - emu(1.3), legend_y - emu(0.02), emu(1.5), emu(0.2), 'TorLegRT',
                [{'text': right_name, 'bold': False, 'color': '555555', 'sz': 1000}])
        return legend_y + emu(0.25)

    def _layout_honeycomb(data, y):
        """蜂巢图（dashi-ppt: honeycomb/hive，2模板）"""
        items = data.get('items', [])  # [{label, value?, color?}]
        if not items:
            return y
        n = min(len(items), 12)
        HONEY_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4']
        # 简化蜂巢：六边形用圆角矩形近似，3列×N行
        cols = 4
        cell_w = emu(2.2)
        cell_h = emu(0.9)
        gap_x = emu(0.1)
        gap_y = emu(0.1)
        rows = (n + cols - 1) // cols
        for i in range(n):
            row = i // cols
            col = i % cols
            it = items[i]
            label = it.get('label', '')
            value = it.get('value', '')
            color = it.get('color', HONEY_COLORS[i % len(HONEY_COLORS)])
            offset_x = emu(1.15) if row % 2 == 1 else emu(0)
            cx = LEFT_X + col * (cell_w + gap_x) + offset_x
            cy = y + row * (cell_h + gap_y)
            # 使用六边形近似（圆角矩形+较大radius）
            _rect(cx, cy, cell_w, cell_h, color, f'Hex{i}', [], radius=20000)
            text = f'{label}\n{value}' if value else label
            textbox(cx, cy, cell_w, cell_h, f'HexT{i}',
                    [{'text': text, 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}])
        return y + rows * (cell_h + gap_y) + emu(0.1)

    def _layout_slope(data, y):
        """斜率图（dashi-ppt: slope，7模板）"""
        items = data.get('items', [])  # [{label, left, right}]
        left_label = data.get('left_label', '前期')
        right_label = data.get('right_label', '后期')
        if not items:
            return y
        n = min(len(items), 6)
        chart_h = emu(2.8)
        left_x = LEFT_X + emu(1.0)
        right_x = LEFT_X + BODY_W - emu(2.5)
        SLOPE_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4']
        # 获取所有值确定范围
        all_vals = []
        for it in items[:n]:
            all_vals.extend([it.get('left', 0), it.get('right', 0)])
        min_val = min(all_vals) if all_vals else 0
        max_val = max(all_vals) if all_vals else 1
        val_range = max_val - min_val or 1
        # 左右轴标签
        textbox(LEFT_X, y - emu(0.3), emu(1.0), emu(0.3), 'SlopeLL',
                [{'text': left_label, 'bold': True, 'color': '999999', 'sz': 1200, 'align': 'ctr'}])
        textbox(LEFT_X + BODY_W - emu(2.5), y - emu(0.3), emu(2.5), emu(0.3), 'SlopeRL',
                [{'text': right_label, 'bold': True, 'color': '999999', 'sz': 1200, 'align': 'ctr'}])
        for i in range(n):
            it = items[i]
            label = it.get('label', '')
            left_val = it.get('left', 0)
            right_val = it.get('right', 0)
            color = SLOPE_COLORS[i % len(SLOPE_COLORS)]
            # 计算Y位置
            ly = y + int(chart_h * (1 - (left_val - min_val) / val_range))
            ry = y + int(chart_h * (1 - (right_val - min_val) / val_range))
            # 左圆点+标签
            dot_sz = emu(0.28)
            _circle(left_x - dot_sz // 2, ly - dot_sz // 2, dot_sz, color, f'SlopeLD{i}',
                    [{'text': str(left_val), 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}])
            textbox(LEFT_X, ly - emu(0.12), emu(1.0), emu(0.25), f'SlopeLT{i}',
                    [{'text': label, 'bold': False, 'color': color, 'sz': 1000, 'align': 'r'}])
            # 右圆点+标签
            _circle(right_x - dot_sz // 2, ry - dot_sz // 2, dot_sz, color, f'SlopeRD{i}',
                    [{'text': str(right_val), 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}])
            textbox(right_x + emu(0.2), ry - emu(0.12), emu(2.0), emu(0.25), f'SlopeRT{i}',
                    [{'text': label, 'bold': False, 'color': color, 'sz': 1000}])
            # 连线
            spTree.append(make_line(left_x, ly, right_x - left_x, color, 15875))
        return y + chart_h + emu(0.15)

    def _layout_pictogram(data, y):
        """象形图 / 单位图（dashi-ppt: pictogram/isotype，2模板）"""
        items = data.get('items', [])  # [{label, value, icon?}]
        total = data.get('total', 100)
        if not items:
            return y
        # 简化：用彩色小方块表示每个单位
        PICTO_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN]
        unit_sz = emu(0.18)
        gap = emu(0.03)
        units_per_row = 20
        # 计算每个项目的方块数
        cells = []
        filled = 0
        for ii, it in enumerate(items):
            count = int(round(it.get('value', 0) / total * 100))
            color = PICTO_COLORS[ii % len(PICTO_COLORS)]
            for _ in range(count):
                if filled >= 100:
                    break
                cells.append((color, it.get('label', '')))
                filled += 1
        while len(cells) < 100:
            cells.append(('E8E8E8', ''))
        # 画方块
        rows = 5
        cols = units_per_row
        for idx in range(min(len(cells), 100)):
            row = idx // cols
            col = idx % cols
            cx = BODY_X + col * (unit_sz + gap)
            cy = y + row * (unit_sz + gap)
            color, _ = cells[idx]
            # 用圆形代替方形，更像象形图
            _rect(cx, cy, unit_sz, unit_sz, color, f'Picto{idx}', [], radius=30000)
        # 图例
        legend_y = y + rows * (unit_sz + gap) + emu(0.15)
        for ii, it in enumerate(items[:5]):
            lx = BODY_X + ii * emu(2.0)
            color = PICTO_COLORS[ii % len(PICTO_COLORS)]
            _rect(lx, legend_y, emu(0.15), emu(0.15), color, f'PictoLeg{ii}', [], radius=30000)
            textbox(lx + emu(0.2), legend_y - emu(0.02), emu(1.5), emu(0.2), f'PictoLegT{ii}',
                    [{'text': f"{it.get('label', '')} ({it.get('value', '')})", 'bold': False, 'color': '333333', 'sz': 1100}])
        return legend_y + emu(0.25)

    def _layout_sunburst(data, y):
        """旭日图（dashi-ppt: sunburst，2模板）"""
        inner = data.get('inner', [])  # [{label, value}]
        outer = data.get('outer', [])  # [{label, parent_idx, value}]
        if not inner:
            return y
        center_x = LEFT_X + emu(2.0)
        center_y = y + emu(1.5)
        inner_r = emu(0.8)
        outer_r = emu(1.4)
        SUNBURST_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4', '2D7A35', 'D4940A']
        # 内圈：用圆+扇形区域表示（简化：用色块圆环）
        total_inner = sum(it.get('value', 0) for it in inner) or 1
        # 画内圈（简化为同心圆环段，用圆形近似）
        _rect(center_x - outer_r, center_y - outer_r, outer_r * 2, outer_r * 2, 'F5F5F5', 'SunBg', [], radius=50000)
        _rect(center_x - inner_r, center_y - inner_r, inner_r * 2, inner_r * 2, WHITE, 'SunInner', [], radius=50000)
        # 内圈标签（简化：中心显示标题）
        textbox(center_x - inner_r, center_y - emu(0.2), inner_r * 2, emu(0.4), 'SunCenter',
                [{'text': '总计', 'bold': False, 'color': '999999', 'sz': 1100, 'align': 'ctr'}])
        # 右侧图例
        legend_x = LEFT_X + emu(4.2)
        for i, it in enumerate(inner):
            color = SUNBURST_COLORS[i % len(SUNBURST_COLORS)]
            ly = y + i * emu(0.4)
            _rect(legend_x, ly, emu(0.2), emu(0.2), color, f'SunIn{i}', [], radius=3000)
            pct = it.get('value', 0) / total_inner * 100
            textbox(legend_x + emu(0.28), ly - emu(0.02), emu(3.5), emu(0.25), f'SunInT{i}',
                    [{'text': f"{it.get('label', '')} ({pct:.0f}%)", 'bold': False, 'color': '333333', 'sz': 1200}])
        # 外圈标签
        if outer:
            outer_y = y + len(inner) * emu(0.4) + emu(0.2)
            textbox(legend_x, outer_y, emu(3.5), emu(0.25), 'SunOutH',
                    [{'text': '细分：', 'bold': True, 'color': '555555', 'sz': 1100}])
            for j, oit in enumerate(outer[:6]):
                oly = outer_y + emu(0.3) + j * emu(0.3)
                parent_idx = oit.get('parent_idx', 0)
                color = SUNBURST_COLORS[parent_idx % len(SUNBURST_COLORS)]
                _rect(legend_x + emu(0.2), oly, emu(0.15), emu(0.15), color, f'SunOut{j}', [], radius=3000)
                textbox(legend_x + emu(0.42), oly - emu(0.02), emu(3.3), emu(0.2), f'SunOutT{j}',
                        [{'text': f"{oit.get('label', '')} ({oit.get('value', '')})", 'bold': False, 'color': '555555', 'sz': 1100}])
            return oly + emu(0.3)
        return y + len(inner) * emu(0.4) + emu(0.3)

    # ── 补齐 dashi-ppt 剩余高频布局（15 种）────────

    def _layout_mekko(data, y):
        """变宽堆叠图 / Marimekko（dashi-ppt: mekko/marimekko，11模板）"""
        items = data.get('items', [])  # [{label, width, segments[{label, value}]}]
        if not items:
            return y
        n = min(len(items), 6)
        chart_h = emu(2.5)
        total_w = sum(it.get('width', 1) for it in items[:n]) or 1
        MEKKO_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4']
        cx = LEFT_X
        for i in range(n):
            it = items[i]
            w = int(BODY_W * it.get('width', 1) / total_w)
            w = max(w, emu(0.5))
            segments = it.get('segments', [])
            seg_total = sum(s.get('value', 1) for s in segments) or 1
            sy = y
            for j, seg in enumerate(segments[:5]):
                sh = int(chart_h * seg.get('value', 1) / seg_total)
                sh = max(sh, emu(0.15))
                color = MEKKO_COLORS[j % len(MEKKO_COLORS)]
                _rect(cx, sy, w - emu(0.03), sh, color, f'Mekk{i}{j}', [], radius=0)
                if sh > emu(0.25):
                    textbox(cx + emu(0.05), sy, w - emu(0.1), sh, f'MekkT{i}{j}',
                            [{'text': seg.get('label', ''), 'bold': True, 'color': WHITE, 'sz': 1000, 'align': 'ctr'}])
                sy += sh
            textbox(cx, y + chart_h + emu(0.05), w, emu(0.25), f'MekkLbl{i}',
                    [{'text': it.get('label', ''), 'bold': False, 'color': '333333', 'sz': 1100, 'align': 'ctr'}])
            cx += w
        return y + chart_h + emu(0.35)

    def _layout_grouped(data, y):
        """分组柱状图（dashi-ppt: grouped/groupbars，3模板）"""
        categories = data.get('categories', [])
        series = data.get('series', [])  # [{label, values[], color?}]
        if not categories or not series:
            return y
        n_cat = min(len(categories), 6)
        n_ser = min(len(series), 3)
        grp_gap = emu(0.3)
        bar_gap = emu(0.04)
        grp_w = (BODY_W - (n_cat - 1) * grp_gap) // n_cat
        bar_w = (grp_w - (n_ser - 1) * bar_gap) // n_ser
        max_val = max(v for s in series for v in s.get('values', [])) or 1
        chart_h = emu(2.2)
        GRP_COLORS = [BRAND_GREEN, GOLD, '4874CB']
        for ci in range(n_cat):
            gx = LEFT_X + ci * (grp_w + grp_gap)
            for si in range(n_ser):
                v = series[si].get('values', [])[ci] if ci < len(series[si].get('values', [])) else 0
                color = series[si].get('color', GRP_COLORS[si % len(GRP_COLORS)])
                bh = max(int(chart_h * v / max_val), emu(0.08))
                bx = gx + si * (bar_w + bar_gap)
                by = y + chart_h - bh
                _rect(bx, by, bar_w, bh, color, f'Grp{ci}{si}', [], radius=3000)
                textbox(bx, by - emu(0.2), bar_w, emu(0.2), f'GrpV{ci}{si}',
                        [{'text': str(v), 'bold': True, 'color': color, 'sz': 900, 'align': 'ctr'}])
            textbox(gx, y + chart_h + emu(0.05), grp_w, emu(0.25), f'GrpLbl{ci}',
                    [{'text': categories[ci], 'bold': False, 'color': '333333', 'sz': 1100, 'align': 'ctr'}])
        legend_y = y + chart_h + emu(0.35)
        for si in range(n_ser):
            lx = LEFT_X + si * emu(2.0)
            color = series[si].get('color', GRP_COLORS[si % len(GRP_COLORS)])
            _rect(lx, legend_y, emu(0.15), emu(0.15), color, f'GrpLeg{si}', [], radius=3000)
            textbox(lx + emu(0.2), legend_y - emu(0.02), emu(1.5), emu(0.2), f'GrpLegT{si}',
                    [{'text': series[si].get('label', f'系列{si+1}'), 'bold': False, 'color': '333333', 'sz': 1000}])
        return legend_y + emu(0.25)

    def _layout_trend(data, y):
        """趋势折线图（dashi-ppt: trend/curve，8模板）"""
        series = data.get('series', [])  # [{label, values[], color?}]
        labels = data.get('labels', [])  # x 轴标签
        if not series:
            return y
        n_pts = max(len(s.get('values', [])) for s in series)
        chart_w = BODY_W - emu(0.8)
        chart_h = emu(2.4)
        all_vals = [v for s in series for v in s.get('values', [])]
        max_val = max(all_vals) if all_vals else 1
        min_val = min(all_vals) if all_vals else 0
        val_range = max_val - min_val or 1
        TREND_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E']
        for si, s in enumerate(series[:4]):
            vals = s.get('values', [])
            color = s.get('color', TREND_COLORS[si % len(TREND_COLORS)])
            prev_cx, prev_cy = None, None
            for pi, v in enumerate(vals[:n_pts]):
                cx = LEFT_X + emu(0.3) + int(chart_w * pi / max(n_pts - 1, 1))
                cy = y + int(chart_h * (1 - (v - min_val) / val_range))
                dot_sz = emu(0.2)
                _circle(cx - dot_sz // 2, cy - dot_sz // 2, dot_sz, color, f'TrendD{si}{pi}',
                        [{'text': str(v), 'bold': True, 'color': WHITE, 'sz': 800, 'align': 'ctr'}])
                if prev_cx is not None:
                    spTree.append(make_line(prev_cx, prev_cy, cx - prev_cx, color, 15875))
                prev_cx, prev_cy = cx, cy
        # x 轴标签
        for pi in range(min(n_pts, len(labels))):
            px = LEFT_X + emu(0.3) + int(chart_w * pi / max(n_pts - 1, 1))
            textbox(px - emu(0.3), y + chart_h + emu(0.1), emu(0.6), emu(0.2), f'TrendLbl{pi}',
                    [{'text': labels[pi], 'bold': False, 'color': '999999', 'sz': 900, 'align': 'ctr'}])
        # 图例
        legend_y = y + chart_h + emu(0.4)
        for si, s in enumerate(series[:4]):
            lx = LEFT_X + si * emu(2.2)
            color = s.get('color', TREND_COLORS[si % len(TREND_COLORS)])
            _rect(lx, legend_y, emu(0.2), emu(0.05), color, f'TrendLeg{si}', [], radius=0)
            textbox(lx + emu(0.25), legend_y - emu(0.05), emu(1.8), emu(0.2), f'TrendLegT{si}',
                    [{'text': s.get('label', ''), 'bold': False, 'color': '333333', 'sz': 1000}])
        return legend_y + emu(0.25)

    def _layout_chain(data, y):
        """产业链 / 价值链（dashi-ppt: chain，6模板）"""
        stages = data.get('stages', [])  # [{label, items[{name}]}]
        if not stages:
            return y
        n_stages = min(len(stages), 5)
        stage_w = (BODY_W - (n_stages - 1) * emu(0.15)) // n_stages
        chain_h = emu(2.8)
        CHAIN_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN]
        for si in range(n_stages):
            stage = stages[si]
            color = CHAIN_COLORS[si % len(CHAIN_COLORS)]
            sx = LEFT_X + si * (stage_w + emu(0.15))
            # 阶段标题
            _rect(sx, y, stage_w, emu(0.4), color, f'ChainH{si}',
                  [{'text': stage.get('label', f'阶段{si+1}'), 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}],
                  radius=6000)
            # 阶段内项目
            items = stage.get('items', [])
            for ii, item in enumerate(items[:4]):
                iy = y + emu(0.5) + ii * (emu(0.5) + emu(0.08))
                name = item.get('name', '') if isinstance(item, dict) else str(item)
                _rect(sx + emu(0.05), iy, stage_w - emu(0.1), emu(0.5), 'F5F5F5', f'ChainI{si}{ii}',
                      [{'text': name, 'bold': False, 'color': '333333', 'sz': 1100, 'align': 'ctr'}],
                      radius=4000, line_color=color, line_w=9525)
            # 连接箭头到下一阶段
            if si < n_stages - 1:
                arr_x = sx + stage_w
                arr_y = y + emu(0.15)
                spTree.append(make_line(arr_x, arr_y, emu(0.15), color, 19050))
        return y + chain_h

    def _layout_calendar(data, y):
        """日历热力图（dashi-ppt: calendar，4模板）"""
        weeks = data.get('weeks', [])  # [{values[7], label?}]
        month_labels = data.get('month_labels', [])
        day_labels = data.get('day_labels', ['一', '二', '三', '四', '五', '六', '日'])
        if not weeks:
            return y
        cell_sz = emu(0.25)
        gap = emu(0.04)
        n_weeks = min(len(weeks), 16)
        # 行标签（周一到周日）
        for di in range(7):
            dy = y + di * (cell_sz + gap)
            label = day_labels[di] if di < len(day_labels) else ''
            textbox(LEFT_X, dy, emu(0.35), cell_sz, f'CalDay{di}',
                    [{'text': label, 'bold': False, 'color': '999999', 'sz': 900, 'align': 'r'}])
        # 画日历格
        for wi in range(n_weeks):
            week = weeks[wi]
            vals = week.get('values', [0]*7)
            for di in range(7):
                v = vals[di] if di < len(vals) else 0
                # 颜色深浅按值 (0-100)
                intensity = max(0, min(100, v))
                r = int(0x46 + (0xFF - 0x46) * (1 - intensity / 100))
                g = int(0xA5 + (0xFF - 0xA5) * (1 - intensity / 100))
                b = int(0x3B + (0xFF - 0x3B) * (1 - intensity / 100))
                color = f'{r:02X}{g:02X}{b:02X}' if intensity > 0 else 'F0F0F0'
                cx = LEFT_X + emu(0.4) + wi * (cell_sz + gap)
                cy = y + di * (cell_sz + gap)
                _rect(cx, cy, cell_sz, cell_sz, color, f'Cal{wi}{di}', [], radius=3000)
        return y + 7 * (cell_sz + gap) + emu(0.1)

    def _layout_orbit(data, y):
        """轨道枢纽图（dashi-ppt: orbit，4模板）"""
        center = data.get('center', '')
        nodes = data.get('nodes', [])  # [{label, orbit?, desc?}]
        if not nodes and not center:
            return y
        center_x = LEFT_X + emu(2.5)
        center_y = y + emu(1.3)
        center_sz = emu(0.9)
        # 中心节点
        _circle(center_x - center_sz // 2, center_y - center_sz // 2, center_sz, BRAND_GREEN, 'OrbitCenter',
                [{'text': center, 'bold': True, 'color': WHITE, 'sz': 1300, 'align': 'ctr'}])
        # 轨道环
        ORBIT_COLORS = [GOLD, '4874CB', 'E54C5E']
        orbits = {}
        for node in nodes:
            orb = node.get('orbit', 1)
            if orb not in orbits:
                orbits[orb] = []
            orbits[orb].append(node)
        for orb_idx, orb_nodes in sorted(orbits.items()):
            r = emu(0.8) + orb_idx * emu(0.55)
            # 轨道圆环（用矩形+大圆角近似）
            _rect(center_x - r, center_y - r, r * 2, r * 2, 'F0F0F0', f'OrbitRing{orb_idx}', [],
                  radius=50000, line_color='E0E0E0', line_w=6350)
            n = len(orb_nodes)
            for ni, node in enumerate(orb_nodes):
                angle = 2 * math.pi * ni / n
                nx = center_x + int(r * math.cos(angle))
                ny = center_y + int(r * math.sin(angle))
                dot_sz = emu(0.5)
                color = ORBIT_COLORS[orb_idx % len(ORBIT_COLORS)]
                label = node.get('label', '')
                _circle(nx - dot_sz // 2, ny - dot_sz // 2, dot_sz, color, f'OrbitN{orb_idx}{ni}',
                        [{'text': label[:4], 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}])
        return y + emu(2.7)

    def _layout_triptych(data, y):
        """三联面板 / 三栏编辑（dashi-ppt: triptych，4模板）"""
        panels = data.get('panels', [])  # [{title, items[], color?}]
        if not panels:
            return y
        n = min(len(panels), 3)
        panel_w = (BODY_W - (n - 1) * emu(0.12)) // n
        panel_h = emu(2.8)
        TRI_COLORS = [BRAND_GREEN, GOLD, '4874CB']
        for pi in range(n):
            p = panels[pi]
            color = p.get('color', TRI_COLORS[pi % len(TRI_COLORS)])
            px = LEFT_X + pi * (panel_w + emu(0.12))
            # 面板背景
            _rect(px, y, panel_w, panel_h, 'F8F8F8', f'TriBg{pi}', [], radius=8000, line_color='E8E8E8', line_w=6350)
            # 标题条
            _rect(px, y, panel_w, emu(0.4), color, f'TriHead{pi}',
                  [{'text': p.get('title', f'面板{pi+1}'), 'bold': True, 'color': WHITE, 'sz': 1400, 'align': 'ctr'}],
                  radius=8000)
            # 内容列表
            items = p.get('items', [])
            for ii, item in enumerate(items[:5]):
                iy = y + emu(0.5) + ii * emu(0.42)
                text = item if isinstance(item, str) else item.get('text', '')
                textbox(px + emu(0.1), iy, panel_w - emu(0.2), emu(0.38), f'TriI{pi}{ii}',
                        [{'text': f'• {text}', 'bold': False, 'color': '333333', 'sz': 1200}])
        return y + panel_h + emu(0.15)

    def _layout_meter(data, y):
        """计量条 / 进度指标行（dashi-ppt: meter，3模板）"""
        items = data.get('items', [])  # [{label, value, max?, benchmark?}]
        if not items:
            return y
        n = min(len(items), 5)
        label_w = emu(1.5)
        bar_w = BODY_W - label_w - emu(1.5)
        row_h = emu(0.55)
        METER_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN]
        for i in range(n):
            it = items[i]
            ry = y + i * (row_h + emu(0.12))
            label = it.get('label', '')
            value = it.get('value', 0)
            max_val = it.get('max', 100)
            benchmark = it.get('benchmark')
            color = METER_COLORS[i % len(METER_COLORS)]
            # 标签
            textbox(LEFT_X, ry, label_w, row_h, f'MeterLbl{i}',
                    [{'text': label, 'bold': False, 'color': '333333', 'sz': 1300, 'align': 'r'}])
            # 背景条
            _rect(LEFT_X + label_w, ry + emu(0.12), bar_w, emu(0.28), 'E8E8E8', f'MeterBg{i}', [], radius=30000)
            # 填充条
            fill_w = max(int(bar_w * value / max_val), emu(0.3))
            _rect(LEFT_X + label_w, ry + emu(0.12), fill_w, emu(0.28), color, f'MeterFill{i}', [], radius=30000)
            # 基准线
            if benchmark is not None:
                bm_x = LEFT_X + label_w + int(bar_w * benchmark / max_val)
                _rect(bm_x, ry + emu(0.05), emu(0.04), emu(0.42), 'E54C5E', f'MeterBm{i}', [], radius=0)
            # 数值
            textbox(LEFT_X + label_w + bar_w + emu(0.1), ry, emu(1.0), row_h, f'MeterVal{i}',
                    [{'text': f'{value}/{max_val}', 'bold': True, 'color': color, 'sz': 1200, 'font': 'Arial'}])
        return y + n * (row_h + emu(0.12))

    def _layout_pareto(data, y):
        """帕累托图（dashi-ppt: pareto，2模板）"""
        items = data.get('items', [])  # [{label, value}]
        if not items:
            return y
        n = min(len(items), 8)
        sorted_items = sorted(items[:n], key=lambda x: x.get('value', 0), reverse=True)
        total = sum(it.get('value', 0) for it in sorted_items) or 1
        max_val = sorted_items[0].get('value', 1) if sorted_items else 1
        chart_h = emu(2.2)
        bar_w = (BODY_W - emu(0.5)) // n
        # 柱状图
        for i, it in enumerate(sorted_items):
            bx = LEFT_X + i * bar_w
            bh = max(int(chart_h * it.get('value', 0) / max_val), emu(0.1))
            by = y + chart_h - bh
            _rect(bx + emu(0.03), by, bar_w - emu(0.06), bh, BRAND_GREEN, f'ParetoBar{i}', [], radius=3000)
            textbox(bx, y + chart_h + emu(0.05), bar_w, emu(0.25), f'ParetoLbl{i}',
                    [{'text': it.get('label', ''), 'bold': False, 'color': '555555', 'sz': 900, 'align': 'ctr'}])
        # 累积百分比线
        cum = 0
        prev_cx, prev_cy = None, None
        for i, it in enumerate(sorted_items):
            cum += it.get('value', 0)
            pct = cum / total
            cx = LEFT_X + i * bar_w + bar_w // 2
            cy = y + int(chart_h * (1 - pct))
            dot_sz = emu(0.12)
            _rect(cx - dot_sz // 2, cy - dot_sz // 2, dot_sz, dot_sz, GOLD, f'ParetoDot{i}', [], radius=30000)
            if prev_cx is not None:
                spTree.append(make_line(prev_cx, prev_cy, cx - prev_cx, GOLD, 12700))
            prev_cx, prev_cy = cx, cy
        # 80% 标注线
        eighty_y = y + int(chart_h * 0.2)
        spTree.append(make_line(LEFT_X, eighty_y, BODY_W, 'E54C5E', 9525))
        textbox(LEFT_X + BODY_W - emu(0.8), eighty_y - emu(0.2), emu(0.8), emu(0.2), 'Pareto80',
                [{'text': '80%', 'bold': True, 'color': 'E54C5E', 'sz': 1000, 'align': 'r'}])
        return y + chart_h + emu(0.35)

    def _layout_delta(data, y):
        """变化量对比（dashi-ppt: delta/deltahero，3模板）"""
        items = data.get('items', [])  # [{label, before, after, unit?}]
        if not items:
            return y
        n = min(len(items), 4)
        card_w = (BODY_W - (n - 1) * emu(0.12)) // n
        card_h = emu(2.0)
        DELTA_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E']
        for i in range(n):
            it = items[i]
            cx = LEFT_X + i * (card_w + emu(0.12))
            color = DELTA_COLORS[i % len(DELTA_COLORS)]
            label = it.get('label', '')
            before = it.get('before', 0)
            after = it.get('after', 0)
            delta = after - before
            unit = it.get('unit', '')
            pct = (delta / before * 100) if before != 0 else 0
            _rect(cx, y, card_w, card_h, 'F8F8F8', f'DeltaBg{i}', [], radius=8000, line_color='E8E8E8', line_w=6350)
            textbox(cx, y + emu(0.1), card_w, emu(0.3), f'DeltaLbl{i}',
                    [{'text': label, 'bold': False, 'color': '555555', 'sz': 1200, 'align': 'ctr'}])
            textbox(cx, y + emu(0.4), card_w, emu(0.4), f'DeltaBefore{i}',
                    [{'text': f'{before}{unit}', 'bold': False, 'color': '999999', 'sz': 1400, 'align': 'ctr', 'font': 'Arial'}])
            # 箭头
            arrow_color = BRAND_GREEN if delta >= 0 else 'E54C5E'
            arrow = '↑' if delta >= 0 else '↓'
            textbox(cx, y + emu(0.85), card_w, emu(0.3), f'DeltaArrow{i}',
                    [{'text': arrow, 'bold': True, 'color': arrow_color, 'sz': 1800, 'align': 'ctr'}])
            textbox(cx, y + emu(1.15), card_w, emu(0.4), f'DeltaAfter{i}',
                    [{'text': f'{after}{unit}', 'bold': True, 'color': '333333', 'sz': 2000, 'align': 'ctr', 'font': 'Arial'}])
            textbox(cx, y + emu(1.55), card_w, emu(0.3), f'DeltaPct{i}',
                    [{'text': f'{pct:+.1f}%', 'bold': True, 'color': arrow_color, 'sz': 1400, 'align': 'ctr', 'font': 'Arial'}])
        return y + card_h + emu(0.15)

    def _layout_milestones(data, y):
        """里程碑时间轴（dashi-ppt: milestones/phases，2模板）"""
        milestones = data.get('milestones', [])  # [{date, title, desc?}]
        if not milestones:
            return y
        n = min(len(milestones), 5)
        # 时间轴线
        axis_y = y + emu(0.8)
        spTree.append(make_line(LEFT_X, axis_y, BODY_W, BRAND_GREEN, 19050))
        ms_w = BODY_W // n
        MS_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN]
        for i in range(n):
            ms = milestones[i]
            mx = LEFT_X + i * ms_w + ms_w // 2
            color = MS_COLORS[i % len(MS_COLORS)]
            # 节点圆
            dot_sz = emu(0.3)
            _rect(mx - dot_sz // 2, axis_y - dot_sz // 2, dot_sz, dot_sz, color, f'MsDot{i}', [], radius=30000)
            # 日期（上方）
            textbox(mx - ms_w // 2, y, ms_w, emu(0.5), f'MsDate{i}',
                    [{'text': ms.get('date', ''), 'bold': True, 'color': color, 'sz': 1200, 'align': 'ctr', 'font': 'Arial'}])
            # 标题（下方）
            textbox(mx - ms_w // 2, axis_y + emu(0.3), ms_w, emu(0.3), f'MsTitle{i}',
                    [{'text': ms.get('title', ''), 'bold': True, 'color': '333333', 'sz': 1200, 'align': 'ctr'}])
            # 描述
            desc = ms.get('desc', '')
            if desc:
                textbox(mx - ms_w // 2, axis_y + emu(0.6), ms_w, emu(0.8), f'MsDesc{i}',
                        [{'text': desc, 'bold': False, 'color': '555555', 'sz': 1100, 'align': 'ctr'}])
        return axis_y + emu(1.5) if any(ms.get('desc') for ms in milestones[:n]) else axis_y + emu(0.7)

    def _layout_spectrum(data, y):
        """光谱定位图（dashi-ppt: spectrum，2模板）"""
        items = data.get('items', [])  # [{label, position(0-100)}]
        left_label = data.get('left_label', '低')
        right_label = data.get('right_label', '高')
        if not items:
            return y
        n = min(len(items), 6)
        bar_h = emu(0.4)
        bar_y = y + emu(0.5)
        # 渐变条（用多段矩形模拟）
        steps = 20
        step_w = BODY_W // steps
        for si in range(steps):
            pct = si / steps
            r = int(0xE5 + (0x46 - 0xE5) * pct)
            g = int(0x4C + (0xA5 - 0x4C) * pct)
            b = int(0x5E + (0x3B - 0x5E) * pct)
            color = f'{r:02X}{g:02X}{b:02X}'
            _rect(LEFT_X + si * step_w, bar_y, step_w + emu(0.02), bar_h, color, f'SpecGrad{si}', [], radius=0)
        # 标签
        textbox(LEFT_X, y, emu(1.0), emu(0.35), 'SpecLeft',
                [{'text': left_label, 'bold': True, 'color': 'E54C5E', 'sz': 1300}])
        textbox(LEFT_X + BODY_W - emu(1.0), y, emu(1.0), emu(0.35), 'SpecRight',
                [{'text': right_label, 'bold': True, 'color': BRAND_GREEN, 'sz': 1300, 'align': 'r'}])
        # 数据点
        SP_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4']
        for i in range(n):
            it = items[i]
            pos = max(0, min(100, it.get('position', 50)))
            px = LEFT_X + int(BODY_W * pos / 100)
            color = SP_COLORS[i % len(SP_COLORS)]
            # 标记圆点
            dot_sz = emu(0.3)
            _circle(px - dot_sz // 2, bar_y - dot_sz // 2 + bar_h // 2, dot_sz, color, f'SpecDot{i}',
                    [{'text': str(i+1), 'bold': True, 'color': WHITE, 'sz': 1000, 'align': 'ctr'}])
            # 标签
            label_y = bar_y + bar_h + emu(0.15) + (i % 2) * emu(0.3)
            textbox(px - emu(0.5), label_y, emu(1.0), emu(0.25), f'SpecLbl{i}',
                    [{'text': it.get('label', ''), 'bold': False, 'color': color, 'sz': 1100, 'align': 'ctr'}])
        return bar_y + bar_h + emu(0.8)

    def _layout_logowall(data, y):
        """Logo 墙 / 伙伴墙（dashi-ppt: logowall，2模板）"""
        items = data.get('items', [])  # [{label}]
        cols = data.get('cols', 5)
        if not items:
            return y
        n = min(len(items), 20)
        cols = min(cols, 6)
        rows = (n + cols - 1) // cols
        cell_w = BODY_W // cols
        cell_h = emu(0.7)
        for i in range(n):
            row = i // cols
            col = i % cols
            cx = LEFT_X + col * cell_w
            cy = y + row * (cell_h + emu(0.08))
            label = items[i].get('label', '') if isinstance(items[i], dict) else str(items[i])
            _rect(cx + emu(0.08), cy, cell_w - emu(0.16), cell_h, 'F8F8F8', f'Logo{i}', [],
                  radius=6000, line_color='E8E8E8', line_w=6350)
            textbox(cx + emu(0.08), cy, cell_w - emu(0.16), cell_h, f'LogoT{i}',
                    [{'text': label, 'bold': False, 'color': '555555', 'sz': 1200, 'align': 'ctr'}])
        return y + rows * (cell_h + emu(0.08))

    def _layout_masonry(data, y):
        """瀑布流网格（dashi-ppt: masonry，2模板）"""
        items = data.get('items', [])  # [{title, desc?, height?}]
        if not items:
            return y
        n = min(len(items), 8)
        cols = 3
        col_w = (BODY_W - (cols - 1) * emu(0.1)) // cols
        MASON_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4', '2D7A35', 'D4940A']
        col_heights = [0] * cols
        for i in range(n):
            it = items[i]
            col = i % cols
            cy = y + col_heights[col]
            h = emu(0.8) + int(emu(0.3) * (i % 3))  # 变化高度
            color = MASON_COLORS[i % len(MASON_COLORS)]
            title = it.get('title', '') if isinstance(it, dict) else str(it)
            desc = it.get('desc', '') if isinstance(it, dict) else ''
            cx = LEFT_X + col * (col_w + emu(0.1))
            _rect(cx, cy, col_w, h, color, f'Mason{i}', [], radius=6000)
            text = f'{title}\n{desc}' if desc else title
            textbox(cx + emu(0.1), cy + emu(0.05), col_w - emu(0.2), h - emu(0.1), f'MasonT{i}',
                    [{'text': text, 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}])
            col_heights[col] += h + emu(0.1)
        return y + max(col_heights) + emu(0.05)

    def _layout_ladder(data, y):
        """转化阶梯 / 留存阶梯（dashi-ppt: ladder，2模板）"""
        stages = data.get('stages', [])  # [{label, value, dropoff?}]
        if not stages:
            return y
        n = min(len(stages), 6)
        max_val = max(s.get('value', 0) for s in stages[:n]) or 1
        chart_w = BODY_W - emu(0.5)
        stair_h = emu(0.45)
        LADDER_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4']
        for i in range(n):
            s = stages[i]
            label = s.get('label', '')
            value = s.get('value', 0)
            dropoff = s.get('dropoff')
            bar_w = max(int(chart_w * value / max_val), emu(0.5))
            sy = y + i * (stair_h + emu(0.15))
            color = LADDER_COLORS[i % len(LADDER_COLORS)]
            # 阶梯条
            _rect(LEFT_X, sy, bar_w, stair_h, color, f'Ladder{i}', [], radius=4000)
            textbox(LEFT_X + emu(0.1), sy, bar_w - emu(0.2), stair_h, f'LadderT{i}',
                    [{'text': f'{label}: {value}', 'bold': True, 'color': WHITE, 'sz': 1300}])
            # 下降指示
            if dropoff is not None and i < n - 1:
                textbox(LEFT_X + bar_w + emu(0.1), sy, emu(1.5), stair_h, f'LadderDrop{i}',
                        [{'text': f'↓ {dropoff}', 'bold': True, 'color': 'E54C5E', 'sz': 1200}])
        return y + n * (stair_h + emu(0.15))

    # ── 动画布局静态化适配（11 种）────────

    def _layout_mindmap(data, y):
        """思维导图 / 放射树（dashi-ppt: mindmap，1模板）"""
        center = data.get('center', '')
        branches = data.get('branches', [])  # [{label, leaves[{label}]}]
        if not branches and not center:
            return y
        cx = LEFT_X + BODY_W // 2
        cy = y + emu(1.5)
        center_sz = emu(0.9)
        _circle(cx - center_sz // 2, cy - center_sz // 2, center_sz, BRAND_GREEN, 'MMCenter',
                [{'text': center, 'bold': True, 'color': WHITE, 'sz': 1400, 'align': 'ctr'}])
        n_br = min(len(branches), 6)
        BR_COLORS = [GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4', '2D7A35']
        for bi in range(n_br):
            br = branches[bi]
            angle = 2 * math.pi * bi / n_br - math.pi / 2
            br_r = emu(1.2)
            bx = cx + int(br_r * math.cos(angle))
            by = cy + int(br_r * math.sin(angle))
            color = BR_COLORS[bi % len(BR_COLORS)]
            br_sz = emu(0.55)
            _circle(bx - br_sz // 2, by - br_sz // 2, br_sz, color, f'MMBr{bi}',
                    [{'text': br.get('label', '')[:4], 'bold': True, 'color': WHITE, 'sz': 1000, 'align': 'ctr'}])
            # 中心到分支的对角连线
            dl = _diag_line(cx, cy, bx, by, color, 15875)
            if dl is not None:
                spTree.append(dl)
            leaves = br.get('leaves', [])
            for li, leaf in enumerate(leaves[:3]):
                la = angle + (li - 1) * 0.4
                lr = emu(0.8)
                lx = bx + int(lr * math.cos(la))
                ly = by + int(lr * math.sin(la))
                leaf_sz = emu(0.4)
                leaf_text = leaf.get('label', '') if isinstance(leaf, dict) else str(leaf)
                _rect(lx - leaf_sz // 2, ly - leaf_sz // 2, leaf_sz, leaf_sz, 'F0F7EE', f'MMLf{bi}{li}',
                      [{'text': leaf_text, 'bold': False, 'color': '333333', 'sz': 900, 'align': 'ctr'}],
                      radius=6000, line_color=color, line_w=9525)
                # 分支到叶子的对角连线
                dl2 = _diag_line(bx, by, lx, ly, color, 9525)
                if dl2 is not None:
                    spTree.append(dl2)
        return y + emu(3.1)

    def _layout_network(data, y):
        """网络节点图（dashi-ppt: network/alliance，2模板）"""
        nodes = data.get('nodes', [])  # [{label, value?}]
        edges = data.get('edges', [])  # [{source, target}]
        center = data.get('center', '')
        if not nodes:
            return y
        cx = LEFT_X + BODY_W // 2
        cy = y + emu(1.5)
        n = min(len(nodes), 8)
        NET_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4', '2D7A35', 'D4940A']
        r = emu(1.3)
        positions = {}
        for i in range(n):
            angle = 2 * math.pi * i / n - math.pi / 2
            nx = cx + int(r * math.cos(angle))
            ny = cy + int(r * math.sin(angle))
            positions[i] = (nx, ny)
            color = NET_COLORS[i % len(NET_COLORS)]
            dot_sz = emu(0.5)
            label = nodes[i].get('label', '') if isinstance(nodes[i], dict) else str(nodes[i])
            _circle(nx - dot_sz // 2, ny - dot_sz // 2, dot_sz, color, f'NetN{i}',
                    [{'text': label[:4], 'bold': True, 'color': WHITE, 'sz': 1000, 'align': 'ctr'}])
        # 连线
        for e in edges[:10]:
            si = e.get('source', 0)
            ti = e.get('target', 1)
            if si in positions and ti in positions:
                sx, sy = positions[si]
                tx, ty = positions[ti]
                spTree.append(make_line(sx, sy, tx - sx, 'CCCCCC', 9525))
        # 中心节点
        if center:
            csz = emu(0.6)
            _circle(cx - csz // 2, cy - csz // 2, csz, BRAND_GREEN, 'NetCenter',
                    [{'text': center[:4], 'bold': True, 'color': WHITE, 'sz': 1100, 'align': 'ctr'}])
        return y + emu(3.1)

    def _layout_mosaic(data, y):
        """图片拼贴 / 马赛克网格（dashi-ppt: mosaic，6模板）"""
        items = data.get('items', [])  # [{title, tag?}]
        if not items:
            return y
        n = min(len(items), 6)
        MOS_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4']
        # 1大+5小布局
        big_w = BODY_W * 2 // 3
        small_w = BODY_W - big_w - emu(0.1)
        big_h = emu(2.8)
        # 大图
        _rect(LEFT_X, y, big_w, big_h, MOS_COLORS[0], 'MosaicBig', [], radius=8000)
        textbox(LEFT_X + emu(0.15), y + big_h - emu(0.5), big_w - emu(0.3), emu(0.4), 'MosaicBigT',
                [{'text': items[0].get('title', '') if isinstance(items[0], dict) else str(items[0]),
                  'bold': True, 'color': WHITE, 'sz': 1600}])
        # 小图
        small_h = (big_h - emu(0.08) * (min(n - 1, 5) - 1)) // max(min(n - 1, 5), 1)
        small_h = min(small_h, emu(0.85))
        for i in range(1, min(n, 6)):
            sx = LEFT_X + big_w + emu(0.1)
            sy = y + (i - 1) * (small_h + emu(0.08))
            color = MOS_COLORS[i % len(MOS_COLORS)]
            _rect(sx, sy, small_w, small_h, color, f'MosaicS{i}', [], radius=6000)
            title = items[i].get('title', '') if isinstance(items[i], dict) else str(items[i])
            textbox(sx + emu(0.08), sy + emu(0.05), small_w - emu(0.16), small_h - emu(0.1), f'MosaicST{i}',
                    [{'text': title, 'bold': True, 'color': WHITE, 'sz': 1100, 'align': 'ctr'}])
        return y + big_h + emu(0.1)

    def _layout_sticker_bubble(data, y):
        """卫星数据 / 估值泡沫（dashi-ppt: sticker-bubble，1模板）"""
        big_number = data.get('big_number', '')
        big_label = data.get('big_label', '')
        satellites = data.get('satellites', [])  # [{label, value}]
        card_h = emu(2.8)
        # 中心大数字
        center_w = emu(4.5)
        _rect(LEFT_X, y, center_w, card_h, DARK_GREEN, 'StkBubble', [], radius=12000)
        textbox(LEFT_X, y + emu(0.3), center_w, emu(1.0), 'StkNum',
                [{'text': big_number, 'bold': True, 'color': GOLD, 'sz': 4800, 'font': 'Arial', 'align': 'ctr'}])
        if big_label:
            textbox(LEFT_X, y + emu(1.3), center_w, emu(0.4), 'StkLbl',
                    [{'text': big_label, 'bold': False, 'color': WHITE, 'sz': 1600, 'align': 'ctr'}])
        # 右侧卫星卡
        rx = LEFT_X + center_w + emu(0.15)
        rw = BODY_W - center_w - emu(0.15)
        n = min(len(satellites), 4)
        sat_h = (card_h - (n - 1) * emu(0.1)) // max(n, 1)
        STK_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E']
        for i in range(n):
            sat = satellites[i]
            sy = y + i * (sat_h + emu(0.1))
            color = STK_COLORS[i % len(STK_COLORS)]
            _rect(rx, sy, rw, sat_h, color, f'StkSat{i}', [], radius=8000)
            label = sat.get('label', '') if isinstance(sat, dict) else str(sat)
            value = sat.get('value', '') if isinstance(sat, dict) else ''
            textbox(rx + emu(0.1), sy + emu(0.05), rw - emu(0.2), sat_h * 0.5, f'StkSL{i}',
                    [{'text': label, 'bold': False, 'color': WHITE, 'sz': 1200}])
            textbox(rx + emu(0.1), sy + sat_h * 0.45, rw - emu(0.2), sat_h * 0.5, f'StkSV{i}',
                    [{'text': str(value), 'bold': True, 'color': WHITE, 'sz': 1800, 'font': 'Arial'}])
        return y + card_h + emu(0.15)

    def _layout_bubbletl(data, y):
        """气泡时间线（dashi-ppt: bubbletl，1模板）"""
        items = data.get('items', [])  # [{label, value, date?}]
        if not items:
            return y
        n = min(len(items), 10)
        chart_w = BODY_W - emu(0.4)
        max_val = max(it.get('value', 0) for it in items[:n]) or 1
        min_val = min(it.get('value', 0) for it in items[:n])
        val_range = max_val - min_val or 1
        base_y = y + emu(2.0)
        # 基线
        spTree.append(make_line(LEFT_X + emu(0.2), base_y, chart_w, BRAND_GREEN, 15875))
        for i in range(n):
            it = items[i]
            val = it.get('value', 0)
            label = it.get('label', '')
            date = it.get('date', '')
            px = LEFT_X + emu(0.2) + int(chart_w * i / max(n - 1, 1))
            # 气泡大小按值
            norm = (val - min_val) / val_range
            dot_sz = emu(0.3) + int(emu(0.5) * norm)
            dot_sz = min(dot_sz, emu(0.8))
            cy = base_y - dot_sz // 2 - emu(0.1)
            color = BRAND_GREEN if val >= (max_val + min_val) / 2 else GOLD
            _circle(px - dot_sz // 2, cy - dot_sz // 2, dot_sz, color, f'BubTl{i}',
                    [{'text': str(val), 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}])
            # 日期标签
            textbox(px - emu(0.3), base_y + emu(0.08), emu(0.6), emu(0.2), f'BubTlD{i}',
                    [{'text': date or label, 'bold': False, 'color': '555555', 'sz': 900, 'align': 'ctr'}])
        return base_y + emu(0.35)

    def _layout_icicle(data, y):
        """冰柱图 / 层级分解（dashi-ppt: icicle，1模板）"""
        items = data.get('items', [])  # [{label, value, children[{label, value}]}]
        if not items:
            return y
        n = min(len(items), 5)
        total = sum(it.get('value', 0) for it in items[:n]) or 1
        ICICLE_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN]
        # 顶层
        top_h = emu(0.5)
        cx = LEFT_X
        for i in range(n):
            it = items[i]
            w = max(int(BODY_W * it.get('value', 0) / total), emu(0.5))
            color = ICICLE_COLORS[i % len(ICICLE_COLORS)]
            _rect(cx, y, w - emu(0.03), top_h, color, f'IcTop{i}',
                  [{'text': it.get('label', ''), 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}],
                  radius=4000)
            children = it.get('children', [])
            if children:
                child_total = sum(c.get('value', 0) for c in children) or 1
                ccx = cx
                child_w_each = (w - emu(0.03)) // len(children[:4])
                for ci, child in enumerate(children[:4]):
                    cw = max(int((w - emu(0.03)) * child.get('value', 0) / child_total), emu(0.3))
                    _rect(ccx, y + top_h + emu(0.06), cw - emu(0.02), emu(0.4), color, f'IcCh{i}{ci}',
                          [], radius=3000, line_color=WHITE, line_w=6350)
                    label = child.get('label', '') if isinstance(child, dict) else str(child)
                    textbox(ccx, y + top_h + emu(0.06), cw - emu(0.02), emu(0.4), f'IcChT{i}{ci}',
                            [{'text': label, 'bold': False, 'color': WHITE, 'sz': 1000, 'align': 'ctr'}])
                    ccx += cw
            cx += w
        return y + top_h + emu(0.52)

    def _layout_candles(data, y):
        """K线图 / 蜡烛图（dashi-ppt: candles，1模板）"""
        items = data.get('items', [])  # [{open, close, high, low, label?}]
        if not items:
            return y
        n = min(len(items), 15)
        chart_h = emu(2.4)
        all_vals = []
        for it in items[:n]:
            all_vals.extend([it.get('high', 0), it.get('low', 0)])
        max_val = max(all_vals) if all_vals else 1
        min_val = min(all_vals) if all_vals else 0
        val_range = max_val - min_val or 1
        candle_w = (BODY_W - emu(0.4)) // n
        for i in range(n):
            it = items[i]
            o = it.get('open', 0)
            c = it.get('close', 0)
            h = it.get('high', 0)
            l = it.get('low', 0)
            is_up = c >= o
            color = BRAND_GREEN if is_up else 'E54C5E'
            cx = LEFT_X + emu(0.2) + i * candle_w + candle_w // 2
            body_top = y + int(chart_h * (1 - (max(o, c) - min_val) / val_range))
            body_bot = y + int(chart_h * (1 - (min(o, c) - min_val) / val_range))
            body_h = max(body_bot - body_top, emu(0.04))
            wick_top = y + int(chart_h * (1 - (h - min_val) / val_range))
            wick_bot = y + int(chart_h * (1 - (l - min_val) / val_range))
            # 上下影线
            spTree.append(make_line(cx, wick_top, emu(0.01), color, 9525))
            _rect(cx - emu(0.005), wick_top, emu(0.01), body_top - wick_top, color, f'CdlWt{i}', [], radius=0)
            _rect(cx - emu(0.005), body_bot, emu(0.01), wick_bot - body_bot, color, f'CdlWb{i}', [], radius=0)
            # 实体
            bar_w = max(candle_w - emu(0.1), emu(0.1))
            _rect(cx - bar_w // 2, body_top, bar_w, body_h, color, f'CdlBody{i}', [], radius=2000)
            # 日期标签
            label = it.get('label', '')
            if label:
                textbox(cx - candle_w // 2, y + chart_h + emu(0.05), candle_w, emu(0.2), f'CdlLbl{i}',
                        [{'text': label, 'bold': False, 'color': '999999', 'sz': 800, 'align': 'ctr'}])
        return y + chart_h + emu(0.3)

    def _layout_hypecycle(data, y):
        """成熟度曲线 /  hype cycle（dashi-ppt: hypecycle，1模板）"""
        items = data.get('items', [])  # [{label, position(0-100)}]
        phases = data.get('phases', ['技术萌芽', '期望膨胀', '泡沫低谷', '稳步爬升', '成熟期'])
        chart_h = emu(2.2)
        chart_w = BODY_W - emu(0.4)
        # 用折线模拟曲线
        # 关键控制点: 0→20% 上升, 20-35% 峰值, 35-55% 低谷, 55-100% 缓慢上升
        key_points = [
            (0.0, 0.2), (0.1, 0.5), (0.2, 0.85), (0.3, 1.0),
            (0.4, 0.6), (0.5, 0.4), (0.55, 0.35),
            (0.65, 0.45), (0.75, 0.55), (0.85, 0.65), (1.0, 0.75)
        ]
        # 画曲线（用折线近似）
        for ki in range(len(key_points) - 1):
            x1 = LEFT_X + emu(0.2) + int(chart_w * key_points[ki][0])
            y1 = y + int(chart_h * (1 - key_points[ki][1]))
            x2 = LEFT_X + emu(0.2) + int(chart_w * key_points[ki + 1][0])
            y2 = y + int(chart_h * (1 - key_points[ki + 1][1]))
            spTree.append(make_line(x1, y1, x2 - x1, BRAND_GREEN, 19050))
        # 阶段标签
        n_phases = min(len(phases), 5)
        phase_w = chart_w // n_phases
        for pi in range(n_phases):
            px = LEFT_X + emu(0.2) + pi * phase_w
            textbox(px, y + chart_h + emu(0.1), phase_w, emu(0.25), f'HypeP{pi}',
                    [{'text': phases[pi], 'bold': False, 'color': '999999', 'sz': 900, 'align': 'ctr'}])
        # 数据点
        HC_COLORS = [GOLD, 'E54C5E', '4874CB', DARK_GREEN, '30C0B4']
        for i, it in enumerate(items[:6]):
            pos = max(0, min(100, it.get('position', 50)))
            # 在曲线上找到对应 y 值（线性插值）
            pct = pos / 100
            yv = 0.2
            for ki in range(len(key_points) - 1):
                if key_points[ki][0] <= pct <= key_points[ki + 1][0]:
                    t = (pct - key_points[ki][0]) / (key_points[ki + 1][0] - key_points[ki][0])
                    yv = key_points[ki][1] + t * (key_points[ki + 1][1] - key_points[ki][1])
                    break
            px = LEFT_X + emu(0.2) + int(chart_w * pct)
            py = y + int(chart_h * (1 - yv))
            color = HC_COLORS[i % len(HC_COLORS)]
            dot_sz = emu(0.25)
            _rect(px - dot_sz // 2, py - dot_sz // 2, dot_sz, dot_sz, color, f'HypeDot{i}', [], radius=30000)
            label = it.get('label', '')
            textbox(px - emu(0.4), py - emu(0.3), emu(0.8), emu(0.2), f'HypeLbl{i}',
                    [{'text': label, 'bold': True, 'color': color, 'sz': 900, 'align': 'ctr'}])
        return y + chart_h + emu(0.4)

    def _layout_typeriver(data, y):
        """字阵 / 标语流（dashi-ppt: typeriver，1模板）"""
        words = data.get('words', [])  # [{text, size?}]
        lead = data.get('lead', '')
        if not words:
            return y
        if lead:
            textbox(BODY_X, y, BODY_W, emu(0.3), 'TRLead',
                    [{'text': lead, 'bold': False, 'color': '555555', 'sz': 1400}])
            y += emu(0.35)
        SIZES = [3600, 2800, 2400, 2000, 1600, 1400]
        COLORS = [BRAND_GREEN, '333333', GOLD, '333333', '555555', '333333']
        for i, w in enumerate(words[:8]):
            text = w.get('text', '') if isinstance(w, dict) else str(w)
            sz = w.get('size', SIZES[i % len(SIZES)]) if isinstance(w, dict) else SIZES[i % len(SIZES)]
            color = COLORS[i % len(COLORS)]
            textbox(BODY_X + (i % 3) * emu(0.3), y, BODY_W - emu(0.6), emu(0.35), f'TR{i}',
                    [{'text': text, 'bold': True if i == 0 else False, 'color': color, 'sz': sz}])
            y += emu(0.35)
        return y

    def _layout_ribbon(data, y):
        """全幅比例带（dashi-ppt: ribbon，1模板）"""
        items = data.get('items', [])  # [{label, value}]
        if not items:
            return y
        n = min(len(items), 6)
        total = sum(it.get('value', 0) for it in items[:n]) or 1
        ribbon_h = emu(0.8)
        RIBBON_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4']
        cx = LEFT_X
        for i in range(n):
            it = items[i]
            w = max(int(BODY_W * it.get('value', 0) / total), emu(0.5))
            color = RIBBON_COLORS[i % len(RIBBON_COLORS)]
            _rect(cx, y, w - emu(0.02), ribbon_h, color, f'Ribbon{i}', [], radius=0)
            pct = it.get('value', 0) / total * 100
            label = it.get('label', '')
            if w > emu(0.8):
                textbox(cx, y, w, ribbon_h, f'RibbonT{i}',
                        [{'text': f'{label}\n{pct:.0f}%', 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}])
            cx += w
        return y + ribbon_h + emu(0.1)

    def _layout_vinyl(data, y):
        """唱片 / 播放列表（dashi-ppt: vinyl，1模板）"""
        title = data.get('title', '')
        tracks = data.get('tracks', [])  # [{label, duration?}]
        card_h = emu(2.8)
        # 左侧唱片
        vinyl_sz = emu(2.5)
        vinyl_cx = LEFT_X + vinyl_sz // 2
        vinyl_cy = y + card_h // 2
        _rect(LEFT_X, y, vinyl_sz, vinyl_sz, '333333', 'VinylBg', [], radius=50000)
        # 同心圆纹
        for ri in range(4):
            r = vinyl_sz // 2 - emu(0.15) - ri * emu(0.25)
            if r > emu(0.3):
                _rect(vinyl_cx - r, vinyl_cy - r, r * 2, r * 2, '333333', f'VinylR{ri}', [],
                      radius=50000, line_color='555555', line_w=6350)
        # 中心标签
        center_sz = emu(0.6)
        _rect(vinyl_cx - center_sz // 2, vinyl_cy - center_sz // 2, center_sz, center_sz, BRAND_GREEN, 'VinylCenter',
              [{'text': title[:2] if title else '♪', 'bold': True, 'color': WHITE, 'sz': 1400, 'align': 'ctr'}],
              radius=50000)
        # 右侧曲目列表
        list_x = LEFT_X + vinyl_sz + emu(0.2)
        if title:
            textbox(list_x, y, BODY_W - vinyl_sz - emu(0.2), emu(0.35), 'VinylTitle',
                    [{'text': title, 'bold': True, 'color': '333333', 'sz': 1800}])
        for i, tr in enumerate(tracks[:6]):
            ty = y + emu(0.4) + i * emu(0.35)
            label = tr.get('label', '') if isinstance(tr, dict) else str(tr)
            dur = tr.get('duration', '') if isinstance(tr, dict) else ''
            textbox(list_x, ty, emu(0.3), emu(0.3), f'VinylN{i}',
                    [{'text': f'{i+1:02d}', 'bold': True, 'color': GOLD, 'sz': 1100, 'font': 'Arial'}])
            textbox(list_x + emu(0.35), ty, emu(4.0), emu(0.3), f'VinylTr{i}',
                    [{'text': label, 'bold': False, 'color': '333333', 'sz': 1200}])
            if dur:
                textbox(list_x + emu(4.5), ty, emu(1.0), emu(0.3), f'VinylDur{i}',
                        [{'text': dur, 'bold': False, 'color': '999999', 'sz': 1100, 'font': 'Arial', 'align': 'r'}])
        return y + card_h + emu(0.15)

    # ── dashi-ppt 蒸馏迁移（20 种）────────

    def _layout_polar_rose(data, y):
        """玫瑰图 / 南丁格尔图（dashi-ppt: polar-rose/rose/polar，3模板）"""
        items = data.get('items', [])  # [{label, value}]
        if not items:
            return y
        n = min(len(items), 8)
        cx = LEFT_X + BODY_W // 2
        cy = y + emu(1.5)
        base_r = emu(1.3)
        max_val = max(it.get('value', 0) for it in items[:n]) or 1
        ROSE_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4', '2D7A35', 'D4940A']
        for i in range(n):
            it = items[i]
            val = it.get('value', 0)
            label = it.get('label', '')
            color = ROSE_COLORS[i % len(ROSE_COLORS)]
            angle_start = 2 * math.pi * i / n - math.pi / 2
            angle_end = 2 * math.pi * (i + 1) / n - math.pi / 2
            r = int(base_r * (val / max_val) ** 0.5)
            r = max(r, emu(0.3))
            # 用矩形近似扇形条
            mid_angle = (angle_start + angle_end) / 2
            bx = cx + int(r * 0.5 * math.cos(mid_angle))
            by = cy + int(r * 0.5 * math.sin(mid_angle))
            dot_sz = int(r * 0.6)
            _rect(bx - dot_sz // 2, by - dot_sz // 2, dot_sz, dot_sz, color, f'Rose{i}',
                  [{'text': f'{label}\n{val}', 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}],
                  radius=max(dot_sz // 3, 3000))
            # 图例
            lx = LEFT_X + BODY_W - emu(1.8)
            ly = y + emu(0.1) + i * emu(0.3)
            _rect(lx, ly, emu(0.2), emu(0.2), color, f'RoseL{i}', [], radius=3000)
            textbox(lx + emu(0.25), ly, emu(1.5), emu(0.2), f'RoseLT{i}',
                    [{'text': f'{label} {val}', 'bold': False, 'color': '333333', 'sz': 1000}])
        return y + emu(3.2)

    def _layout_histogram(data, y):
        """直方图 / 频率分布（dashi-ppt: histogram/distribution，2模板）"""
        bins = data.get('bins', [])  # [{range, count}]
        x_label = data.get('x_label', '')
        y_label = data.get('y_label', '频次')
        if not bins:
            return y
        n = min(len(bins), 15)
        chart_h = emu(2.2)
        chart_w = BODY_W - emu(0.8)
        max_count = max(b.get('count', 0) for b in bins[:n]) or 1
        bar_w = chart_w // n
        for i in range(n):
            b = bins[i]
            count = b.get('count', 0)
            h = int(chart_h * count / max_count)
            h = max(h, emu(0.05))
            bx = LEFT_X + emu(0.4) + i * bar_w
            by = y + chart_h - h
            color = BRAND_GREEN if count >= max_count * 0.7 else GOLD
            _rect(bx, by, bar_w - emu(0.03), h, color, f'Hist{i}', [], radius=0)
            range_label = b.get('range', '')
            textbox(bx, y + chart_h + emu(0.05), bar_w, emu(0.2), f'HistL{i}',
                    [{'text': range_label, 'bold': False, 'color': '999999', 'sz': 800, 'align': 'ctr'}])
        # 基线
        spTree.append(make_line(LEFT_X + emu(0.4), y + chart_h, chart_w, 'CCCCCC', 9525))
        return y + chart_h + emu(0.35)

    def _layout_quotewall(data, y):
        """群言墙 / 多引述马赛克（dashi-ppt: quotewall/voices，2模板）"""
        quotes = data.get('quotes', [])  # [{text, author?}]
        if not quotes:
            return y
        n = min(len(quotes), 6)
        QW_COLORS = [BRAND_GREEN, GOLD, DARK_GREEN, '4874CB', 'E54C5E', '30C0B4']
        # 2行×3列
        cols = 3 if n >= 3 else n
        rows = (n + cols - 1) // cols
        card_w = (BODY_W - emu(0.1) * (cols - 1)) // cols
        card_h = emu(1.3)
        for i in range(n):
            q = quotes[i]
            row = i // cols
            col = i % cols
            cx = LEFT_X + col * (card_w + emu(0.1))
            cy = y + row * (card_h + emu(0.1))
            color = QW_COLORS[i % len(QW_COLORS)]
            _rect(cx, cy, card_w, card_h, 'F8F8F8', f'QW{i}', [], radius=8000, line_color=color, line_w=12700)
            text = q.get('text', '') if isinstance(q, dict) else str(q)
            author = q.get('author', '') if isinstance(q, dict) else ''
            textbox(cx + emu(0.12), cy + emu(0.1), card_w - emu(0.24), card_h - emu(0.4), f'QWT{i}',
                    [{'text': f'"{text}"', 'bold': False, 'color': '333333', 'sz': 1100}])
            if author:
                textbox(cx + emu(0.12), cy + card_h - emu(0.3), card_w - emu(0.24), emu(0.2), f'QWA{i}',
                        [{'text': f'— {author}', 'bold': True, 'color': color, 'sz': 1000, 'align': 'r'}])
        return y + rows * (card_h + emu(0.1))

    def _layout_metro(data, y):
        """地铁线路图（dashi-ppt: metro，1模板）"""
        lines = data.get('lines', [])  # [{name, color?, stations[{name}]}]
        if not lines:
            return y
        n_lines = min(len(lines), 3)
        LINE_COLORS = [BRAND_GREEN, GOLD, 'E54C5E']
        station_sz = emu(0.2)
        line_h = emu(0.8)
        for li in range(n_lines):
            ln = lines[li]
            color = ln.get('color', LINE_COLORS[li % len(LINE_COLORS)])
            ly = y + li * line_h
            # 线路名
            textbox(LEFT_X, ly, emu(1.0), emu(0.25), f'MetroName{li}',
                    [{'text': ln.get('name', f'Line {li+1}'), 'bold': True, 'color': color, 'sz': 1200}])
            stations = ln.get('stations', [])
            n_st = min(len(stations), 6)
            station_w = (BODY_W - emu(1.2)) // max(n_st - 1, 1)
            # 画线
            spTree.append(make_line(LEFT_X + emu(1.0), ly + emu(0.3), BODY_W - emu(1.2), color, 25400))
            for si in range(n_st):
                st = stations[si]
                sx = LEFT_X + emu(1.0) + si * station_w - station_sz // 2
                _rect(sx, ly + emu(0.3) - station_sz // 2, station_sz, station_sz, WHITE, f'Metro{li}{si}',
                      [], radius=50000, line_color=color, line_w=19050)
                name = st.get('name', '') if isinstance(st, dict) else str(st)
                textbox(sx - emu(0.15), ly + emu(0.5), station_sz + emu(0.3), emu(0.2), f'MetroLbl{li}{si}',
                        [{'text': name, 'bold': False, 'color': '333333', 'sz': 900, 'align': 'ctr'}])
        return y + n_lines * line_h + emu(0.1)

    def _layout_balance(data, y):
        """天平 / 权衡对比（dashi-ppt: balance，1模板）"""
        left = data.get('left', {})   # {title, items[]}
        right = data.get('right', {})  # {title, items[]}
        card_h = emu(2.5)
        half_w = BODY_W // 2 - emu(0.1)
        # 横梁
        beam_y = y + emu(0.3)
        spTree.append(make_line(LEFT_X + emu(0.3), beam_y, BODY_W - emu(0.6), GOLD, 19050))
        # 支点
        pivot_sz = emu(0.3)
        _rect(LEFT_X + BODY_W // 2 - pivot_sz // 2, beam_y - pivot_sz, pivot_sz, pivot_sz, GOLD, 'BalPivot',
              [], radius=50000)
        # 左盘
        _rect(LEFT_X, y + emu(0.6), half_w, card_h, 'F0F7EE', 'BalLeft', [], radius=8000, line_color=BRAND_GREEN, line_w=12700)
        left_title = left.get('title', '方案A') if isinstance(left, dict) else str(left)
        textbox(LEFT_X + emu(0.1), y + emu(0.7), half_w - emu(0.2), emu(0.3), 'BalLeftT',
                [{'text': left_title, 'bold': True, 'color': BRAND_GREEN, 'sz': 1600}])
        left_items = left.get('items', []) if isinstance(left, dict) else []
        for i, item in enumerate(left_items[:5]):
            textbox(LEFT_X + emu(0.15), y + emu(1.1) + i * emu(0.25), half_w - emu(0.3), emu(0.25), f'BalLI{i}',
                    [{'text': f'• {item}', 'bold': False, 'color': '333333', 'sz': 1200}])
        # 右盘
        rx = LEFT_X + BODY_W // 2 + emu(0.1)
        _rect(rx, y + emu(0.6), half_w, card_h, 'FFF8EC', 'BalRight', [], radius=8000, line_color=GOLD, line_w=12700)
        right_title = right.get('title', '方案B') if isinstance(right, dict) else str(right)
        textbox(rx + emu(0.1), y + emu(0.7), half_w - emu(0.2), emu(0.3), 'BalRightT',
                [{'text': right_title, 'bold': True, 'color': GOLD, 'sz': 1600}])
        right_items = right.get('items', []) if isinstance(right, dict) else []
        for i, item in enumerate(right_items[:5]):
            textbox(rx + emu(0.15), y + emu(1.1) + i * emu(0.25), half_w - emu(0.3), emu(0.25), f'BalRI{i}',
                    [{'text': f'• {item}', 'bold': False, 'color': '333333', 'sz': 1200}])
        return y + emu(0.6) + card_h + emu(0.15)

    def _layout_fiveforces(data, y):
        """波特五力模型（dashi-ppt: fiveforces，1模板）"""
        center = data.get('center', '行业竞争')
        forces = data.get('forces', [])  # [{name, level?}] 上/下/左/右/中心竞争
        card_h = emu(3.0)
        cx = LEFT_X + BODY_W // 2
        cy = y + card_h // 2
        # 中心
        center_sz = emu(1.2)
        _rect(cx - center_sz // 2, cy - center_sz // 2, center_sz, center_sz, BRAND_GREEN, 'FFCenter',
              [{'text': center, 'bold': True, 'color': WHITE, 'sz': 1400, 'align': 'ctr'}], radius=8000)
        # 5个力: 上/下/左/右 + 新进入者(上方偏移)
        positions = [
            (cx, y + emu(0.1)),                     # 上
            (cx, y + card_h - emu(0.8)),             # 下
            (LEFT_X + emu(0.1), cy - emu(0.35)),     # 左
            (LEFT_X + BODY_W - emu(2.3), cy - emu(0.35)),  # 右
            (cx - emu(2.2), y + emu(0.1)),           # 左上
        ]
        FF_COLORS = [GOLD, 'E54C5E', '4874CB', DARK_GREEN, '30C0B4']
        n = min(len(forces), 5)
        box_w = emu(2.0)
        box_h = emu(0.7)
        for i in range(n):
            f = forces[i]
            name = f.get('name', '') if isinstance(f, dict) else str(f)
            level = f.get('level', '') if isinstance(f, dict) else ''
            fx, fy = positions[i]
            color = FF_COLORS[i % len(FF_COLORS)]
            _rect(fx, fy, box_w, box_h, color, f'FF{i}',
                  [{'text': name, 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}], radius=6000)
            if level:
                textbox(fx, fy + box_h - emu(0.25), box_w, emu(0.2), f'FFLvl{i}',
                        [{'text': level, 'bold': False, 'color': WHITE, 'sz': 900, 'align': 'ctr'}])
            # 连线到中心
            spTree.append(make_line(fx + box_w // 2, fy + box_h // 2, cx - fx - box_w // 2, 'CCCCCC', 9525))
        return y + card_h + emu(0.15)

    def _layout_glossary(data, y):
        """术语表 / 词汇表（dashi-ppt: glossary，1模板）"""
        items = data.get('items', [])  # [{term, definition}]
        if not items:
            return y
        n = min(len(items), 8)
        row_h = emu(0.35)
        for i in range(n):
            it = items[i]
            term = it.get('term', '') if isinstance(it, dict) else ''
            defn = it.get('definition', '') if isinstance(it, dict) else ''
            ry = y + i * (row_h + emu(0.05))
            bg_color = 'F0F7EE' if i % 2 == 0 else 'FFFFFF'
            _rect(LEFT_X, ry, BODY_W, row_h, bg_color, f'Gloss{i}', [], radius=4000)
            textbox(LEFT_X + emu(0.15), ry, emu(2.0), row_h, f'GlossT{i}',
                    [{'text': term, 'bold': True, 'color': BRAND_GREEN, 'sz': 1400}])
            textbox(LEFT_X + emu(2.2), ry, BODY_W - emu(2.4), row_h, f'GlossD{i}',
                    [{'text': defn, 'bold': False, 'color': '333333', 'sz': 1200}])
        return y + n * (row_h + emu(0.05)) + emu(0.1)

    def _layout_album(data, y):
        """专辑 / 成就清单（dashi-ppt: album，1模板）"""
        title = data.get('title', '')
        tracks = data.get('tracks', [])  # [{name, detail?, year?}]
        card_h = emu(2.8)
        # 左侧封面区
        cover_sz = emu(2.5)
        _rect(LEFT_X, y, cover_sz, cover_sz, DARK_GREEN, 'AlbumCover', [], radius=8000)
        textbox(LEFT_X, y + emu(0.3), cover_sz, emu(0.4), 'AlbumTitle',
                [{'text': title[:6] if title else '♪', 'bold': True, 'color': GOLD, 'sz': 2000, 'align': 'ctr'}])
        # 右侧曲目
        list_x = LEFT_X + cover_sz + emu(0.2)
        list_w = BODY_W - cover_sz - emu(0.2)
        if title:
            textbox(list_x, y, list_w, emu(0.3), 'AlbumFullT',
                    [{'text': title, 'bold': True, 'color': '333333', 'sz': 1600}])
        n = min(len(tracks), 6)
        track_h = (card_h - emu(0.4)) // max(n, 1)
        for i in range(n):
            tr = tracks[i]
            ty = y + emu(0.4) + i * track_h
            name = tr.get('name', '') if isinstance(tr, dict) else str(tr)
            detail = tr.get('detail', '') if isinstance(tr, dict) else ''
            year = tr.get('year', '') if isinstance(tr, dict) else ''
            textbox(list_x, ty, emu(0.3), emu(0.25), f'AlbN{i}',
                    [{'text': f'{i+1:02d}', 'bold': True, 'color': GOLD, 'sz': 1200, 'font': 'Arial'}])
            textbox(list_x + emu(0.35), ty, list_w - emu(0.35), emu(0.25), f'AlbName{i}',
                    [{'text': name, 'bold': True, 'color': '333333', 'sz': 1300}])
            if detail or year:
                textbox(list_x + emu(0.35), ty + emu(0.2), list_w - emu(0.35), emu(0.2), f'AlbDet{i}',
                        [{'text': f'{year}  {detail}', 'bold': False, 'color': '999999', 'sz': 1000}])
        return y + card_h + emu(0.15)

    def _layout_bracket(data, y):
        """分组括号图（dashi-ppt: bracket，1模板）"""
        groups = data.get('groups', [])  # [{label, items[]}]
        if not groups:
            return y
        n = min(len(groups), 4)
        GRP_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E']
        group_h = emu(0.65)
        gap = emu(0.1)
        for i in range(n):
            g = groups[i]
            gy = y + i * (group_h + gap)
            color = GRP_COLORS[i % len(GRP_COLORS)]
            label = g.get('label', '') if isinstance(g, dict) else str(g)
            items = g.get('items', []) if isinstance(g, dict) else []
            # 标签
            label_w = emu(1.5)
            _rect(LEFT_X, gy, label_w, group_h, color, f'BktLbl{i}',
                  [{'text': label, 'bold': True, 'color': WHITE, 'sz': 1300, 'align': 'ctr'}], radius=6000)
            # 括号竖线
            spTree.append(make_line(LEFT_X + label_w + emu(0.1), gy, emu(0.01), color, 12700))
            # 子项横排
            n_items = min(len(items), 4)
            item_w = (BODY_W - label_w - emu(0.4)) // max(n_items, 1)
            for j in range(n_items):
                ix = LEFT_X + label_w + emu(0.2) + j * item_w
                item_text = items[j] if not isinstance(items[j], dict) else items[j].get('name', '')
                _rect(ix, gy + emu(0.05), item_w - emu(0.05), group_h - emu(0.1), 'F5F5F5', f'BktItem{i}{j}',
                      [{'text': item_text, 'bold': False, 'color': '333333', 'sz': 1100, 'align': 'ctr'}],
                      radius=4000, line_color=color, line_w=6350)
        return y + n * (group_h + gap) + emu(0.1)

    def _layout_horizon(data, y):
        """视野图 / 三视野（dashi-ppt: horizon，1模板）"""
        views = data.get('views', [])  # [{title, items[]}]
        if not views:
            return y
        n = min(len(views), 3)
        HOR_COLORS = ['4874CB', BRAND_GREEN, GOLD]
        HOR_LABELS = ['当前视野', '战略视野', '未来视野']
        panel_w = (BODY_W - emu(0.2) * (n - 1)) // n
        panel_h = emu(2.8)
        for i in range(n):
            v = views[i]
            px = LEFT_X + i * (panel_w + emu(0.2))
            color = HOR_COLORS[i % len(HOR_COLORS)]
            label = v.get('title', HOR_LABELS[i]) if isinstance(v, dict) else HOR_LABELS[i]
            items = v.get('items', []) if isinstance(v, dict) else []
            # 面板
            _rect(px, y, panel_w, panel_h, 'F8F8F8', f'Hor{i}', [], radius=8000, line_color=color, line_w=12700)
            # 顶部色条
            _rect(px, y, panel_w, emu(0.08), color, f'HorBar{i}', [], radius=0)
            # 标题
            textbox(px + emu(0.1), y + emu(0.15), panel_w - emu(0.2), emu(0.3), f'HorT{i}',
                    [{'text': label, 'bold': True, 'color': color, 'sz': 1400}])
            # 子项
            for j, item in enumerate(items[:5]):
                textbox(px + emu(0.15), y + emu(0.55) + j * emu(0.35), panel_w - emu(0.3), emu(0.3), f'HorI{i}{j}',
                        [{'text': f'• {item}', 'bold': False, 'color': '333333', 'sz': 1100}])
        return y + panel_h + emu(0.15)

    def _layout_stack(data, y):
        """架构栈 / 技术栈（dashi-ppt: stack，1模板）"""
        layers = data.get('layers', [])  # [{label, items[]}]
        if not layers:
            return y
        n = min(len(layers), 5)
        STK_COLORS = [DARK_GREEN, BRAND_GREEN, '4874CB', GOLD, '30C0B4']
        layer_h = emu(0.5)
        gap = emu(0.08)
        for i in range(n):
            ly = layers[i]
            lyy = y + i * (layer_h + gap)
            color = STK_COLORS[i % len(STK_COLORS)]
            label = ly.get('label', '') if isinstance(ly, dict) else str(ly)
            items = ly.get('items', []) if isinstance(ly, dict) else []
            # 层标题
            label_w = emu(1.8)
            _rect(LEFT_X, lyy, label_w, layer_h, color, f'StkLbl{i}',
                  [{'text': label, 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}], radius=4000)
            # 层内容
            n_items = min(len(items), 4)
            if n_items > 0:
                item_w = (BODY_W - label_w - emu(0.2)) // n_items
                for j in range(n_items):
                    ix = LEFT_X + label_w + emu(0.1) + j * item_w
                    item_text = items[j] if not isinstance(items[j], dict) else items[j].get('name', '')
                    _rect(ix, lyy, item_w - emu(0.05), layer_h, 'F5F5F5', f'StkI{i}{j}',
                          [{'text': item_text, 'bold': False, 'color': '333333', 'sz': 1100, 'align': 'ctr'}],
                          radius=4000, line_color=color, line_w=6350)
        return y + n * (layer_h + gap) + emu(0.1)

    def _layout_gate(data, y):
        """分层防线 / 门控模型（dashi-ppt: gate，1模板）"""
        layers = data.get('layers', [])  # [{label, desc?}]
        if not layers:
            return y
        n = min(len(layers), 5)
        GT_COLORS = ['E54C5E', 'D4940A', GOLD, BRAND_GREEN, DARK_GREEN]
        # 同心矩形（从外到内缩小）
        max_h = emu(2.8)
        max_w = BODY_W
        step_w = emu(0.8)
        step_h = emu(0.5)
        for i in range(n):
            color = GT_COLORS[i % len(GT_COLORS)]
            lx = LEFT_X + i * step_w
            ly = y + i * step_h
            lw = max_w - 2 * i * step_w
            lh = max_h - 2 * i * step_h
            lw = max(lw, emu(1.5))
            lh = max(lh, emu(0.5))
            _rect(lx, ly, lw, lh, color, f'Gate{i}', [], radius=6000)
            label = layers[i].get('label', '') if isinstance(layers[i], dict) else str(layers[i])
            desc = layers[i].get('desc', '') if isinstance(layers[i], dict) else ''
            textbox(lx, ly + lh // 2 - emu(0.15), lw, emu(0.3), f'GateT{i}',
                    [{'text': label, 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}])
        return y + max_h + emu(0.15)

    def _layout_triad(data, y):
        """三角 / 三球串联（dashi-ppt: triad/spheres，2模板）"""
        items = data.get('items', [])  # [{label, desc?}]
        center = data.get('center', '')
        if not items:
            return y
        n = min(len(items), 3)
        TRI_COLORS = [BRAND_GREEN, GOLD, '4874CB']
        cx = LEFT_X + BODY_W // 2
        cy = y + emu(1.3)
        positions = [
            (cx - emu(0.9), y + emu(0.2)),                    # 上
            (LEFT_X + emu(0.3), y + emu(1.7)),                # 左下
            (LEFT_X + BODY_W - emu(2.3), y + emu(1.7)),       # 右下
        ]
        node_sz = emu(1.4)
        for i in range(n):
            it = items[i]
            label = it.get('label', '') if isinstance(it, dict) else str(it)
            desc = it.get('desc', '') if isinstance(it, dict) else ''
            color = TRI_COLORS[i % len(TRI_COLORS)]
            nx, ny = positions[i]
            _rect(nx, ny, node_sz, node_sz, color, f'Triad{i}', [], radius=50000)
            textbox(nx, ny + emu(0.2), node_sz, emu(0.3), f'TriadT{i}',
                    [{'text': label, 'bold': True, 'color': WHITE, 'sz': 1400, 'align': 'ctr'}])
            if desc:
                textbox(nx + emu(0.1), ny + emu(0.6), node_sz - emu(0.2), emu(0.8), f'TriadD{i}',
                        [{'text': desc, 'bold': False, 'color': WHITE, 'sz': 1100, 'align': 'ctr'}])
            # 连线
            next_i = (i + 1) % n
            nx2, ny2 = positions[next_i]
            spTree.append(make_line(nx + node_sz // 2, ny + node_sz // 2,
                                    nx2 - nx, 'CCCCCC', 9525))
        return y + emu(1.7) + node_sz + emu(0.15)

    def _layout_loop(data, y):
        """闭环循环 / 无限循环（dashi-ppt: loop，1模板）"""
        steps = data.get('steps', [])  # [{label}]
        center = data.get('center', '')
        if not steps:
            return y
        n = min(len(steps), 6)
        cx = LEFT_X + BODY_W // 2
        cy = y + emu(1.5)
        r = emu(1.2)
        # 中心圆
        center_sz = emu(0.9)
        _circle(cx - center_sz // 2, cy - center_sz // 2, center_sz, BRAND_GREEN, 'LoopCenter',
                [{'text': center, 'bold': True, 'color': WHITE, 'sz': 1300, 'align': 'ctr'}])
        LOOP_COLORS = [GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4', '2D7A35']
        node_sz = emu(0.6)
        for i in range(n):
            angle = 2 * math.pi * i / n - math.pi / 2
            nx = cx + int(r * math.cos(angle)) - node_sz // 2
            ny = cy + int(r * math.sin(angle)) - node_sz // 2
            color = LOOP_COLORS[i % len(LOOP_COLORS)]
            label = steps[i].get('label', '') if isinstance(steps[i], dict) else str(steps[i])
            _rect(nx, ny, node_sz, node_sz, color, f'Loop{i}',
                  [{'text': label, 'bold': True, 'color': WHITE, 'sz': 1000, 'align': 'ctr'}], radius=8000)
        return y + emu(1.5) + r + node_sz // 2 + emu(0.2)

    def _layout_ecosystem(data, y):
        """生态网络 / 核心辐射（dashi-ppt: ecosystem/nexus，2模板）"""
        center = data.get('center', '')
        nodes = data.get('nodes', [])  # [{label, orbit?}]
        if not nodes and not center:
            return y
        cx = LEFT_X + BODY_W // 2
        cy = y + emu(1.5)
        # 中心
        center_sz = emu(1.0)
        _circle(cx - center_sz // 2, cy - center_sz // 2, center_sz, BRAND_GREEN, 'EcoCenter',
                [{'text': center, 'bold': True, 'color': WHITE, 'sz': 1300, 'align': 'ctr'}])
        n = min(len(nodes), 8)
        ECO_COLORS = [GOLD, '4874CB', 'E54C5E', DARK_GREEN, '30C0B4', '2D7A35', 'D4940A', '7B68EE']
        r = emu(1.4)
        node_sz = emu(0.5)
        for i in range(n):
            nd = nodes[i]
            label = nd.get('label', '') if isinstance(nd, dict) else str(nd)
            angle = 2 * math.pi * i / n - math.pi / 2
            nx = cx + int(r * math.cos(angle)) - node_sz // 2
            ny = cy + int(r * math.sin(angle)) - node_sz // 2
            color = ECO_COLORS[i % len(ECO_COLORS)]
            _rect(nx, ny, node_sz, node_sz, color, f'Eco{i}',
                  [{'text': label[:4], 'bold': True, 'color': WHITE, 'sz': 900, 'align': 'ctr'}], radius=6000)
            # 辐射线
            spTree.append(make_line(cx, cy, nx + node_sz // 2 - cx, 'DDDDDD', 6350))
        return y + emu(1.5) + r + node_sz // 2 + emu(0.2)

    def _layout_chronicle(data, y):
        """编年史 / 纵向时间线（dashi-ppt: chronicle，1模板）"""
        events = data.get('events', [])  # [{year, title, desc?}]
        if not events:
            return y
        n = min(len(events), 5)
        row_h = emu(0.55)
        gap = emu(0.08)
        CHR_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN]
        # 纵向轴线
        axis_x = LEFT_X + emu(0.9)
        spTree.append(make_line(axis_x, y, emu(0.01), BRAND_GREEN, 12700))
        for i in range(n):
            ev = events[i]
            ey = y + i * (row_h + gap)
            color = CHR_COLORS[i % len(CHR_COLORS)]
            # 年份
            year = ev.get('year', '') if isinstance(ev, dict) else ''
            textbox(LEFT_X, ey, emu(0.8), row_h, f'ChrY{i}',
                    [{'text': str(year), 'bold': True, 'color': color, 'sz': 1400, 'font': 'Arial', 'align': 'r'}])
            # 圆点
            dot_sz = emu(0.15)
            _rect(axis_x - dot_sz // 2, ey + row_h // 2 - dot_sz // 2, dot_sz, dot_sz, color, f'ChrDot{i}',
                  [], radius=50000)
            # 内容
            title = ev.get('title', '') if isinstance(ev, dict) else ''
            desc = ev.get('desc', '') if isinstance(ev, dict) else ''
            content_x = axis_x + emu(0.2)
            textbox(content_x, ey, BODY_W - emu(1.2), emu(0.25), f'ChrT{i}',
                    [{'text': title, 'bold': True, 'color': '333333', 'sz': 1300}])
            if desc:
                textbox(content_x, ey + emu(0.25), BODY_W - emu(1.2), emu(0.25), f'ChrD{i}',
                        [{'text': desc, 'bold': False, 'color': '999999', 'sz': 1100}])
        return y + n * (row_h + gap) + emu(0.1)

    def _layout_manifesto(data, y):
        """宣言 / 主张页（dashi-ppt: manifesto，1模板）"""
        text = data.get('text', '')
        sub = data.get('sub', '')
        if not text:
            return y
        card_h = emu(3.0)
        _rect(LEFT_X, y, BODY_W, card_h, DARK_GREEN, 'Manifesto', [], radius=0)
        # 金色装饰线
        _rect(LEFT_X + emu(0.3), y + emu(0.3), emu(0.06), card_h - emu(0.6), GOLD, 'ManDeco', [], radius=0)
        textbox(LEFT_X + emu(0.6), y + emu(0.5), BODY_W - emu(1.0), card_h - emu(1.2), 'ManText',
                [{'text': text, 'bold': True, 'color': WHITE, 'sz': 2800, 'align': 'ctr'}])
        if sub:
            textbox(LEFT_X + emu(0.6), y + card_h - emu(0.5), BODY_W - emu(1.0), emu(0.3), 'ManSub',
                    [{'text': sub, 'bold': False, 'color': GOLD, 'sz': 1400, 'align': 'ctr'}])
        return y + card_h + emu(0.15)

    def _layout_comparetable(data, y):
        """特性对照表（dashi-ppt: comparetable，1模板）"""
        headers = data.get('headers', [])  # [str]
        rows = data.get('rows', [])  # [{feature, values[]}]
        if not headers and not rows:
            return y
        n_cols = min(len(headers), 4)
        n_rows = min(len(rows), 7)
        col_w = BODY_W // max(n_cols, 1)
        row_h = emu(0.35)
        # 表头
        for c in range(n_cols):
            _rect(LEFT_X + c * col_w, y, col_w - emu(0.03), emu(0.4), BRAND_GREEN, f'CtH{c}',
                  [{'text': headers[c] if c < len(headers) else '', 'bold': True, 'color': WHITE, 'sz': 1200, 'align': 'ctr'}],
                  radius=4000)
        # 数据行
        for r in range(n_rows):
            row = rows[r]
            ry = y + emu(0.45) + r * (row_h + emu(0.03))
            bg = 'F0F7EE' if r % 2 == 0 else 'FFFFFF'
            feature = row.get('feature', '') if isinstance(row, dict) else ''
            values = row.get('values', []) if isinstance(row, dict) else []
            textbox(LEFT_X, ry, col_w - emu(0.05), row_h, f'CtF{r}',
                    [{'text': feature, 'bold': True, 'color': '333333', 'sz': 1100}])
            for c in range(min(len(values), n_cols - 1)):
                val = values[c]
                _rect(LEFT_X + (c + 1) * col_w, ry, col_w - emu(0.03), row_h, bg, f'CtR{r}{c}',
                      [{'text': str(val), 'bold': False, 'color': '333333', 'sz': 1100, 'align': 'ctr'}],
                      radius=3000)
        return y + emu(0.45) + n_rows * (row_h + emu(0.03)) + emu(0.1)

    def _layout_dotfield(data, y):
        """点阵计数 / 单位图（dashi-ppt: dotfield/dotplot，2模板）"""
        items = data.get('items', [])  # [{label, value, total?}]
        if not items:
            return y
        n = min(len(items), 5)
        total_default = 20
        DOT_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN]
        row_h = emu(0.5)
        for i in range(n):
            it = items[i]
            label = it.get('label', '') if isinstance(it, dict) else ''
            value = it.get('value', 0) if isinstance(it, dict) else 0
            total = it.get('total', total_default) if isinstance(it, dict) else total_default
            color = DOT_COLORS[i % len(DOT_COLORS)]
            ry = y + i * (row_h + emu(0.08))
            # 标签
            textbox(LEFT_X, ry, emu(1.5), row_h, f'DotLbl{i}',
                    [{'text': f'{label} {value}/{total}', 'bold': True, 'color': '333333', 'sz': 1100}])
            # 点阵
            dot_sz = emu(0.12)
            dot_gap = emu(0.04)
            dots_x = LEFT_X + emu(1.8)
            max_dots = min(total, 20)
            for d in range(max_dots):
                dx = dots_x + d * (dot_sz + dot_gap)
                if dx + dot_sz > LEFT_X + BODY_W:
                    break
                fill = color if d < value else 'E0E0E0'
                _rect(dx, ry + row_h // 2 - dot_sz // 2, dot_sz, dot_sz, fill, f'Dot{i}{d}', [], radius=50000)
        return y + n * (row_h + emu(0.08)) + emu(0.1)

    def _layout_bullet(data, y):
        """子弹图 / 目标达成（dashi-ppt: bullet/progress/goals，3模板）"""
        items = data.get('items', [])  # [{label, value, target, max?}]
        if not items:
            return y
        n = min(len(items), 5)
        bar_h = emu(0.25)
        row_h = emu(0.5)
        BLT_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN]
        for i in range(n):
            it = items[i]
            label = it.get('label', '') if isinstance(it, dict) else ''
            value = it.get('value', 0) if isinstance(it, dict) else 0
            target = it.get('target', 100) if isinstance(it, dict) else 100
            max_val = it.get('max', target * 1.2) if isinstance(it, dict) else target * 1.2
            color = BLT_COLORS[i % len(BLT_COLORS)]
            ry = y + i * row_h
            # 标签
            textbox(LEFT_X, ry, emu(1.5), row_h, f'BlLbl{i}',
                    [{'text': label, 'bold': True, 'color': '333333', 'sz': 1200}])
            # 背景条
            bar_x = LEFT_X + emu(1.8)
            bar_w = BODY_W - emu(2.0)
            _rect(bar_x, ry + (row_h - bar_h) // 2, bar_w, bar_h, 'E0E0E0', f'BlBg{i}', [], radius=4000)
            # 实际值条
            val_w = max(int(bar_w * value / max_val), emu(0.1))
            _rect(bar_x, ry + (row_h - bar_h) // 2, val_w, bar_h, color, f'BlVal{i}', [], radius=4000)
            # 目标线
            target_x = bar_x + int(bar_w * target / max_val)
            spTree.append(make_line(target_x, ry + emu(0.05), emu(0.01), '333333', 19050))
            # 数值
            textbox(bar_x + val_w + emu(0.05), ry, emu(1.0), row_h, f'BlNum{i}',
                    [{'text': str(value), 'bold': True, 'color': color, 'sz': 1200, 'font': 'Arial'}])
        return y + n * row_h + emu(0.1)

    # ── 最终补齐（2 种）────────

    def _layout_combo(data, y):
        """组合图 / 柱线双轴（dashi-ppt: rounds/combo，2模板）"""
        categories = data.get('categories', [])  # [str]
        bars = data.get('bars', [])              # [{label, values[], color?}]
        line = data.get('line', {})              # {label, values[], color?}
        if not categories:
            return y
        n = min(len(categories), 8)
        chart_h = emu(2.0)
        chart_w = BODY_W - emu(0.6)
        # 柱状图数据
        max_bar = 1
        for b in bars:
            vals = b.get('values', [])
            if vals:
                max_bar = max(max_bar, max(vals[:n]))
        # 折线数据
        line_vals = line.get('values', [])[:n]
        max_line = max(line_vals) if line_vals else 1
        # 颜色
        BAR_COLORS = [BRAND_GREEN, GOLD, '4874CB']
        line_color = line.get('color', 'E54C5E')
        n_bars = min(len(bars), 3)
        # 柱宽
        group_w = chart_w // n
        bar_w = group_w // (n_bars + 1) if n_bars > 0 else group_w // 2
        # 画柱
        for bi in range(n_bars):
            b = bars[bi]
            color = b.get('color', BAR_COLORS[bi % len(BAR_COLORS)])
            vals = b.get('values', [])[:n]
            for ci in range(n):
                val = vals[ci] if ci < len(vals) else 0
                h = max(int(chart_h * val / max_bar), emu(0.03))
                bx = LEFT_X + emu(0.3) + ci * group_w + bi * bar_w + emu(0.05)
                by = y + chart_h - h
                _rect(bx, by, bar_w - emu(0.04), h, color, f'CmbB{bi}{ci}', [], radius=2000)
        # 画折线（用线段连接点）
        if line_vals:
            points = []
            for ci in range(n):
                val = line_vals[ci] if ci < len(line_vals) else 0
                px = LEFT_X + emu(0.3) + ci * group_w + group_w // 2
                py = y + int(chart_h * (1 - val / max_line))
                points.append((px, py))
            for pi in range(len(points) - 1):
                x1, y1 = points[pi]
                x2, y2 = points[pi + 1]
                spTree.append(make_line(x1, y1, x2 - x1, line_color, 19050))
            # 数据点圆
            for pi, (px, py) in enumerate(points):
                dot_sz = emu(0.15)
                _rect(px - dot_sz // 2, py - dot_sz // 2, dot_sz, dot_sz, line_color, f'CmbD{pi}',
                      [], radius=50000)
        # X轴标签
        for ci in range(n):
            lx = LEFT_X + emu(0.3) + ci * group_w
            textbox(lx, y + chart_h + emu(0.05), group_w, emu(0.2), f'CmbX{ci}',
                    [{'text': categories[ci] if ci < len(categories) else '', 'bold': False, 'color': '999999', 'sz': 900, 'align': 'ctr'}])
        # 图例
        leg_x = LEFT_X + emu(0.3)
        leg_y = y + chart_h + emu(0.3)
        for bi in range(n_bars):
            b = bars[bi]
            color = b.get('color', BAR_COLORS[bi % len(BAR_COLORS)])
            _rect(leg_x, leg_y, emu(0.15), emu(0.12), color, f'CmbLg{bi}', [], radius=2000)
            textbox(leg_x + emu(0.18), leg_y, emu(1.0), emu(0.12), f'CmbLgT{bi}',
                    [{'text': b.get('label', ''), 'bold': False, 'color': '333333', 'sz': 900}])
            leg_x += emu(1.3)
        if line_vals:
            _rect(leg_x, leg_y + emu(0.04), emu(0.15), emu(0.04), line_color, 'CmbLgLine', [], radius=0)
            textbox(leg_x + emu(0.18), leg_y, emu(1.0), emu(0.12), 'CmbLgLineT',
                    [{'text': line.get('label', ''), 'bold': False, 'color': '333333', 'sz': 900}])
        # 基线
        spTree.append(make_line(LEFT_X + emu(0.3), y + chart_h, chart_w, 'CCCCCC', 9525))
        return y + chart_h + emu(0.55)

    def _layout_stream(data, y):
        """主题河流 / 中心流式堆叠（dashi-ppt: stream，1模板）"""
        series = data.get('series', [])  # [{label, values[], color?}]
        labels = data.get('labels', [])  # 时间点标签
        if not series:
            return y
        n_series = min(len(series), 5)
        n_points = min(len(series[0].get('values', [])), 6) if series else 0
        chart_h = emu(2.2)
        chart_w = BODY_W - emu(0.6)
        STREAM_COLORS = [BRAND_GREEN, GOLD, '4874CB', 'E54C5E', DARK_GREEN]
        # 计算每个时间点的总值
        totals = []
        for pi in range(n_points):
            t = 0
            for si in range(n_series):
                vals = series[si].get('values', [])
                t += vals[pi] if pi < len(vals) else 0
            totals.append(max(t, 1))
        max_total = max(totals) if totals else 1
        # 每个时间点的堆叠柱
        seg_w = chart_w // max(n_points, 1)
        for pi in range(n_points):
            px = LEFT_X + emu(0.3) + pi * seg_w
            # 从中心向两侧堆叠
            total = totals[pi]
            cum = 0
            for si in range(n_series):
                vals = series[si].get('values', [])
                val = vals[pi] if pi < len(vals) else 0
                h = max(int(chart_h * val / max_total), emu(0.02))
                color = series[si].get('color', STREAM_COLORS[si % len(STREAM_COLORS)])
                sy = y + int(chart_h * 0.5) - int(chart_h * total / max_total / 2) + int(chart_h * cum / max_total)
                _rect(px + emu(0.02), sy, seg_w - emu(0.04), h, color, f'Strm{si}{pi}', [], radius=2000)
                cum += val
            # 时间标签
            lbl = labels[pi] if pi < len(labels) else ''
            textbox(px, y + chart_h + emu(0.05), seg_w, emu(0.2), f'StrmLbl{pi}',
                    [{'text': lbl, 'bold': False, 'color': '999999', 'sz': 900, 'align': 'ctr'}])
        # 图例
        leg_x = LEFT_X + emu(0.3)
        leg_y = y + chart_h + emu(0.3)
        for si in range(n_series):
            color = series[si].get('color', STREAM_COLORS[si % len(STREAM_COLORS)])
            _rect(leg_x, leg_y, emu(0.15), emu(0.12), color, f'StrmLg{si}', [], radius=2000)
            textbox(leg_x + emu(0.18), leg_y, emu(1.2), emu(0.12), f'StrmLgT{si}',
                    [{'text': series[si].get('label', ''), 'bold': False, 'color': '333333', 'sz': 900}])
            leg_x += emu(1.5)
        return y + chart_h + emu(0.55)

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
        # ── dashiai-ppt 风格新增布局 ──
        'comparison_cards': _layout_comparison_cards,
        'process_circles':  _layout_process_circles,
        'tag_list':         _layout_tag_list,
        'highlight_box':    _layout_highlight_box,
        'dual_stat':        _layout_dual_stat,
        'card_grid':        _layout_card_grid,
        'before_after':     _layout_before_after,
        # ── 新增布局（参考 dashi-ppt）──
        'swot':             _layout_swot,
        'quadrant':         _layout_quadrant,
        'checklist':        _layout_checklist,
        'scorecard':        _layout_scorecard,
        'stair':            _layout_stair,
        'flywheel':         _layout_flywheel,
        'statement':        _layout_statement,
        'journey':          _layout_journey,
        'pricing':          _layout_pricing,
        'faq':              _layout_faq,
        # ── 本轮新增 9 种（图表 / 仪表盘 / 网格）──
        'pie_chart':        _layout_pie_chart,
        'bar_chart':        _layout_bar_chart,
        'dashboard':        _layout_dashboard,
        'hero_banner':      _layout_hero_banner,
        'numbered_list':    _layout_numbered_list,
        'matrix_2x2':       _layout_matrix_2x2,
        'chart_placeholder':_layout_chart_placeholder,
        'kpi_card':         _layout_kpi_card,
        'feature_grid':     _layout_feature_grid,
        # ── 本轮补齐 15 种（dashi-ppt 语义角色全覆盖）──
        'radar_chart':      _layout_radar_chart,
        'pyramid':          _layout_pyramid,
        'roadmap':          _layout_roadmap,
        'venn':             _layout_venn,
        'ranking':          _layout_ranking,
        'waterfall':        _layout_waterfall,
        'heatmap':          _layout_heatmap,
        'gantt':            _layout_gantt,
        'cycle':            _layout_cycle,
        'big_number':       _layout_big_number,
        'gallery':          _layout_gallery,
        'layers':           _layout_layers,
        'bento':            _layout_bento,
        'gauge':            _layout_gauge,
        'testimonial':      _layout_testimonial,
        'treemap':          _layout_treemap,
        'scatter':          _layout_scatter,
        'stacked_bar':      _layout_stacked_bar,
        'profile':          _layout_profile,
        'spotlight':        _layout_spotlight,
        'risk':             _layout_risk,
        'swimlane':         _layout_swimlane,
        'overview':         _layout_overview,
        'principles':       _layout_principles,
        'org_chart':        _layout_org_chart,
        'bump':             _layout_bump,
        'dumbbell':         _layout_dumbbell,
        'lollipop':         _layout_lollipop,
        'waffle':           _layout_waffle,
        'radial_bar':       _layout_radial_bar,
        'diverging':        _layout_diverging,
        'tornado':          _layout_tornado,
        'honeycomb':        _layout_honeycomb,
        'slope':            _layout_slope,
        'pictogram':        _layout_pictogram,
        'sunburst':         _layout_sunburst,
        'mekko':            _layout_mekko,
        'grouped':          _layout_grouped,
        'trend':            _layout_trend,
        'chain':            _layout_chain,
        'calendar':         _layout_calendar,
        'orbit':            _layout_orbit,
        'triptych':         _layout_triptych,
        'meter':            _layout_meter,
        'pareto':           _layout_pareto,
        'delta':            _layout_delta,
        'milestones':       _layout_milestones,
        'spectrum':         _layout_spectrum,
        'logowall':         _layout_logowall,
        'masonry':          _layout_masonry,
        'ladder':           _layout_ladder,
        'mindmap':          _layout_mindmap,
        'network':          _layout_network,
        'mosaic':           _layout_mosaic,
        'sticker_bubble':   _layout_sticker_bubble,
        'bubbletl':         _layout_bubbletl,
        'icicle':           _layout_icicle,
        'candles':          _layout_candles,
        'hypecycle':        _layout_hypecycle,
        'typeriver':        _layout_typeriver,
        'ribbon':           _layout_ribbon,
        'vinyl':            _layout_vinyl,
        'polar_rose':       _layout_polar_rose,
        'histogram':        _layout_histogram,
        'quotewall':        _layout_quotewall,
        'metro':            _layout_metro,
        'balance':          _layout_balance,
        'fiveforces':       _layout_fiveforces,
        'glossary':         _layout_glossary,
        'album':            _layout_album,
        'bracket':          _layout_bracket,
        'horizon':          _layout_horizon,
        'stack':            _layout_stack,
        'gate':             _layout_gate,
        'triad':            _layout_triad,
        'loop':             _layout_loop,
        'ecosystem':        _layout_ecosystem,
        'chronicle':        _layout_chronicle,
        'manifesto':        _layout_manifesto,
        'comparetable':     _layout_comparetable,
        'dotfield':         _layout_dotfield,
        'bullet':           _layout_bullet,
        'combo':            _layout_combo,
        'stream':           _layout_stream,
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
            for rp in p.findall('a:r', namespaces={'a': A}):
                rPr = rp.find('a:rPr', namespaces={'a': A})
                if rPr is not None:
                    rPr.set('sz', '2400')
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
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    SP_TAG = f'{{{P}}}sp'
    for i in range(1, 6):
        for p in slide.xpath('//a:p', namespaces={'a': A}):
            runs = p.findall('a:r', namespaces={'a': A})
            text = ''.join(r.find('a:t', namespaces={'a': A}).text or '' for r in runs)
            if text == f'目录标题{i}':
                if i <= len(items):
                    replace_para_text(p, items[i - 1])
                else:
                    # Remove the entire shape containing this paragraph + 对应序号的 sp
                    sp = p
                    while sp is not None and sp.tag != SP_TAG:
                        sp = sp.getparent()
                    if sp is not None and sp.getparent() is not None:
                        sp.getparent().remove(sp)
                    # 同时删除对应的序号圆圈形状（id 比标题小 1 的 sp）
                break
    # 删除多余项对应的序号圆圈（编号为5的项）
    if len(items) < 5:
        for i in range(len(items) + 1, 6):
            # 找到包含数字 i 的段落所在 sp 并删除
            for p in slide.xpath('//a:p', namespaces={'a': A}):
                runs = p.findall('a:r', namespaces={'a': A})
                text = ''.join(r.find('a:t', namespaces={'a': A}).text or '' for r in runs)
                if text.strip() == str(i):
                    sp = p
                    while sp is not None and sp.tag != SP_TAG:
                        sp = sp.getparent()
                    if sp is not None and sp.getparent() is not None:
                        sp.getparent().remove(sp)
                    break
    save(slide, path)

# ── 编辑章节页 ───────────────────────────────────────────
def edit_section(path, data):
    slide = etree.parse(path).getroot()

    # ── 调整标题文本框（id=30）：开启换行 + 加宽 ──────────────
    NS = {'a': A, 'p': P, 'r': R}
    for sp in slide.xpath('//p:sp', namespaces=NS):
        cNvPr = sp.xpath('.//p:cNvPr', namespaces=NS)
        if cNvPr and cNvPr[0].get('id') == '30':
            bodyPr = sp.xpath('.//a:bodyPr', namespaces=NS)
            if bodyPr:
                bodyPr[0].set('wrap', 'square')       # 开启文字换行
            # 保持 x 不变（竖线右侧），扩展 cx 至金色边框右边界（9109393 - 5608955 = 3500000）
            xfrm = sp.xpath('.//a:xfrm', namespaces=NS)
            if xfrm:
                ext = xfrm[0].xpath('a:ext', namespaces=NS)
                if ext: ext[0].set('cx', '3500000')
            break

    for p in slide.xpath('//a:p', namespaces={'a': A}):
        runs = p.findall('a:r', namespaces={'a': A})
        text = ''.join(r.find('a:t', namespaces={'a': A}).text or '' for r in runs)
        if 'PART' in text or 'Part' in text:
            replace_para_text(p, 'PART')
        elif text.strip().isdigit():
            replace_para_text(p, data.get('part', '01'))
        elif text.strip().startswith('20'):
            replace_para_text(p, data.get('year', ''))
        elif len(text.strip()) >= 2 and 'PART' not in text and not text.strip().startswith('20') and not text.strip().isdigit():
            new_title = data.get('title', '')
            replace_para_text(p, new_title)
            # 根据标题长度调整字号，保证最多两行
            # 框宽 3500000 EMU ≈ 3.83in，28pt 中文约 10 字/行，两行≈20字
            title_len = len(new_title)
            target_sz = '2800'
            if title_len > 20:
                target_sz = '2000'   # 超长标题（>20字）缩到 20pt
            elif title_len > 10:
                target_sz = '2400'   # 中等标题（11~20字）24pt
            for rp in p.findall('a:r', namespaces={'a': A}):
                rPr = rp.find('a:rPr', namespaces={'a': A})
                if rPr is not None:
                    rPr.set('sz', target_sz)
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
    # 先读取 presentation.xml.rels，建立 rId → slide filename 的映射
    pres_rels_path = os.path.join(work, 'ppt', '_rels', 'presentation.xml.rels')
    pres_rels = etree.parse(pres_rels_path).getroot()
    REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    rid_to_fname = {}
    for rel in pres_rels.findall(f'{{{REL_NS}}}Relationship'):
        rid = rel.get('Id')
        target = rel.get('Target', '')
        if target.startswith('slides/'):
            rid_to_fname[rid] = os.path.basename(target)

    old_ids = list(sldIdLst)
    # 建立 slide filename → element 的映射
    id_map = {}
    for e in old_ids:
        rid = e.get(f'{{{R}}}id')
        fname = rid_to_fname.get(rid)
        if fname:
            id_map[fname] = e
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
