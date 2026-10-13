# -*- coding: utf-8 -*-
"""生成《先进制造装备关键零部件研发及前沿技术》学术报告 PPT（约 100 分钟）。

用法:
    pip install python-pptx
    python scripts/build_key_components_ppt.py [输出路径]

每页均附有演讲者备注（时间分配 + 讲解要点），可在 PowerPoint "备注视图" 查看。
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ---------- 主题色 ----------
NAVY = RGBColor(0x0B, 0x1F, 0x3A)
NAVY_2 = RGBColor(0x12, 0x2E, 0x55)
TEAL = RGBColor(0x00, 0x9E, 0xB8)
TEAL_LIGHT = RGBColor(0xDD, 0xF4, 0xF8)
AMBER = RGBColor(0xF2, 0xA3, 0x1D)
AMBER_LIGHT = RGBColor(0xFD, 0xF1, 0xDC)
RED = RGBColor(0xD6, 0x45, 0x45)
GREEN = RGBColor(0x2E, 0x9E, 0x6B)
GRAY = RGBColor(0x5F, 0x6B, 0x7A)
GRAY_LIGHT = RGBColor(0xF2, 0xF4, 0xF7)
GRAY_MID = RGBColor(0xD9, 0xDE, 0xE5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x22, 0x22, 0x22)
STEEL = RGBColor(0x8D, 0x99, 0xAE)

FONT = "Microsoft YaHei"

W, H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.55)
CONTENT_TOP = Inches(1.35)
CONTENT_W = W - 2 * MARGIN

prs = Presentation()
prs.slide_width, prs.slide_height = W, H
BLANK = prs.slide_layouts[6]

SECTIONS = [
    ("一", "绪论：先进制造装备与关键零部件", "约 10 分钟", "绪论"),
    ("二", "核心基础零部件研发现状与关键技术", "约 35 分钟", "核心零部件"),
    ("三", "关键零部件研发的共性使能技术", "约 20 分钟", "共性技术"),
    ("四", "前沿技术与颠覆性方向", "约 25 分钟", "前沿技术"),
    ("五", "发展趋势、研究建议与总结", "约 10 分钟", "趋势与建议"),
]
_state = {"page": 0, "section": ""}


# ---------- 基础工具 ----------
def _style_run(run, size, bold, color, font=FONT, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", font)


def set_text(tf, text, size=18, bold=False, color=BLACK, align=PP_ALIGN.LEFT, italic=False):
    tf.text = ""
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    _style_run(run, size, bold, color, italic=italic)
    return p


def add_text(slide, left, top, width, height, text, size=18, bold=False, color=BLACK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False, line_spacing=None):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    lines = text.split("\n")
    p = set_text(tf, lines[0], size, bold, color, align, italic)
    if line_spacing:
        p.line_spacing = line_spacing
    for extra in lines[1:]:
        p = tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        r = p.add_run()
        r.text = extra
        _style_run(r, size, bold, color, italic=italic)
    return tb


def add_rect(slide, left, top, width, height, fill, line=None, shape=MSO_SHAPE.RECTANGLE,
             line_width=1.0):
    s = slide.shapes.add_shape(shape, left, top, width, height)
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_width)
    s.shadow.inherit = False
    return s


def shape_text(shape, text, size=14, bold=False, color=WHITE, align=PP_ALIGN.CENTER,
               anchor=MSO_ANCHOR.MIDDLE, margin=0.08):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    lines = text.split("\n")
    set_text(tf, lines[0], size, bold, color, align)
    for extra in lines[1:]:
        p = tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = extra
        _style_run(r, size, bold, color)


def add_bullets(slide, left, top, width, height, items, size=17, color=BLACK,
                bullet_color=TEAL, line_spacing=1.15, para_space=7):
    """items: list of str 或 (str, level)。支持 '**' 包裹的加粗片段。"""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    first = True
    for item in items:
        text, level = (item, 0) if isinstance(item, str) else item
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.line_spacing = line_spacing
        p.space_after = Pt(para_space if level == 0 else para_space - 3)
        marker = "■ " if level == 0 else "– "
        r0 = p.add_run()
        r0.text = ("      " * level) + marker
        _style_run(r0, size - 7 if level == 0 else size - 3,
                   False, bullet_color if level == 0 else GRAY)
        sz = size if level == 0 else size - 2
        col = color if level == 0 else GRAY
        parts = text.split("**")
        for i, part in enumerate(parts):
            if not part:
                continue
            r = p.add_run()
            r.text = part
            _style_run(r, sz, i % 2 == 1, NAVY_2 if i % 2 == 1 else col)
    return tb


def add_line(slide, x1, y1, x2, y2, color=STEEL, width=1.5, dash=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = Pt(width)
    if dash:
        from pptx.enum.dml import MSO_LINE_DASH_STYLE
        c.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    return c


def add_notes(slide, minutes, text):
    tf = slide.notes_slide.notes_text_frame
    tf.text = f"【建议用时 {minutes}】\n{text}"


# ---------- 版式 ----------
def content_slide(title, subtitle=None, notes=("1.5 分钟", "")):
    _state["page"] += 1
    slide = prs.slides.add_slide(BLANK)
    add_rect(slide, 0, 0, W, Inches(1.05), NAVY)
    add_rect(slide, 0, Inches(1.05), W, Inches(0.05), TEAL)
    add_text(slide, MARGIN, Inches(0.16), Inches(10.4), Inches(0.6), title,
             size=26, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, MARGIN, Inches(0.68), Inches(10.6), Inches(0.35), subtitle,
                 size=13, color=TEAL_LIGHT)
    if _state["section"]:
        add_text(slide, W - Inches(3.4), Inches(0.22), Inches(2.9), Inches(0.6), _state["section"],
                 size=12, color=STEEL, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    # 页脚
    add_line(slide, MARGIN, Inches(7.02), W - MARGIN, Inches(7.02), color=GRAY_MID, width=0.75)
    add_text(slide, MARGIN, Inches(7.05), Inches(8), Inches(0.35),
             "先进制造装备关键零部件研发及前沿技术", size=10, color=GRAY)
    add_text(slide, W - MARGIN - Inches(1.5), Inches(7.05), Inches(1.5), Inches(0.35),
             f"{_state['page']}", size=10, color=GRAY, align=PP_ALIGN.RIGHT)
    if notes:
        add_notes(slide, notes[0], notes[1])
    return slide


def section_slide(idx, notes_text=""):
    num, title, duration, short = SECTIONS[idx]
    _state["page"] += 1
    _state["section"] = f"第{num}部分 · {short}"
    slide = prs.slides.add_slide(BLANK)
    add_rect(slide, 0, 0, W, H, NAVY)
    add_rect(slide, 0, 0, Inches(4.6), H, NAVY_2)
    add_text(slide, Inches(0.7), Inches(2.3), Inches(3.5), Inches(2.0), f"0{idx + 1}",
             size=110, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(0.75), Inches(4.3), Inches(3.5), Inches(0.5), f"第{num}部分",
             size=18, color=STEEL)
    add_text(slide, Inches(5.3), Inches(2.4), Inches(7.6), Inches(1.7), title,
             size=32, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.15)
    add_rect(slide, Inches(5.3), Inches(4.2), Inches(1.2), Inches(0.06), AMBER)
    add_text(slide, Inches(5.3), Inches(4.4), Inches(7), Inches(0.5), duration,
             size=18, color=TEAL_LIGHT)
    # 进度条
    for i, (n, t, _, short) in enumerate(SECTIONS):
        x = Inches(5.3) + i * Inches(1.5)
        add_rect(slide, x, Inches(6.3), Inches(1.35), Inches(0.1), TEAL if i <= idx else STEEL)
        add_text(slide, x, Inches(6.45), Inches(1.35), Inches(0.4), short,
                 size=10, color=TEAL_LIGHT if i == idx else STEEL)
    add_notes(slide, "0.5 分钟", notes_text or f"进入第{num}部分：{title}。简述本部分要回答的问题与时间安排。")
    return slide


def add_cards(slide, cards, cols, top=CONTENT_TOP, left=MARGIN, width=CONTENT_W,
              height=None, gap=Inches(0.2), title_size=15, body_size=12.5,
              header_h=Inches(0.5), colors=None):
    """cards: list of (title, body)；body 支持 '\n' 分行。"""
    rows = (len(cards) + cols - 1) // cols
    if height is None:
        height = Inches(6.85) - top
    cw = (width - gap * (cols - 1)) / cols
    ch = (height - gap * (rows - 1)) / rows
    colors = colors or [TEAL, NAVY_2, AMBER, GREEN, STEEL, RED]
    for i, (t, body) in enumerate(cards):
        r, c = divmod(i, cols)
        x = left + c * (cw + gap)
        y = top + r * (ch + gap)
        add_rect(slide, x, y, cw, ch, WHITE, line=GRAY_MID)
        hdr = add_rect(slide, x, y, cw, header_h, colors[i % len(colors)])
        shape_text(hdr, t, size=title_size, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
        add_text(slide, x + Inches(0.08), y + header_h + Inches(0.08), cw - Inches(0.16),
                 ch - header_h - Inches(0.12), body, size=body_size, color=BLACK,
                 line_spacing=1.12)


def add_table(slide, data, left=MARGIN, top=CONTENT_TOP, width=CONTENT_W, height=None,
              col_widths=None, font_size=12, header_fill=NAVY, first_col_bold=True,
              row_h=Inches(0.42)):
    rows, cols = len(data), len(data[0])
    height = height or row_h * rows
    gt = slide.shapes.add_table(rows, cols, left, top, width, height)
    tbl = gt.table
    if col_widths:
        total = sum(col_widths)
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = int(width * cw / total)
    for r in range(rows):
        tbl.rows[r].height = row_h
        for c in range(cols):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.07)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            is_header = r == 0
            fill = header_fill if is_header else (GRAY_LIGHT if r % 2 == 0 else WHITE)
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill
            set_text(tf, str(data[r][c]),
                     size=font_size + (0.5 if is_header else 0),
                     bold=is_header or (first_col_bold and c == 0),
                     color=WHITE if is_header else (NAVY_2 if c == 0 and first_col_bold else BLACK),
                     align=PP_ALIGN.CENTER if is_header else PP_ALIGN.LEFT)
    return tbl


def add_chevrons(slide, steps, left=MARGIN, top=CONTENT_TOP, width=CONTENT_W,
                 height=Inches(0.8), size=13, colors=None):
    n = len(steps)
    colors = colors or [NAVY_2, TEAL, NAVY_2, TEAL, NAVY_2, TEAL, NAVY_2, TEAL]
    sw = width / n
    for i, s in enumerate(steps):
        shp = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        box = add_rect(slide, left + i * sw, top, sw + Inches(0.08), height, colors[i % len(colors)],
                       shape=shp)
        shape_text(box, s, size=size, bold=True, color=WHITE, margin=0.18)


def add_kpis(slide, items, top=CONTENT_TOP, left=MARGIN, width=CONTENT_W, height=Inches(1.3),
             big_size=28, label_size=11.5, gap=Inches(0.2)):
    """items: list of (big, label)."""
    n = len(items)
    cw = (width - gap * (n - 1)) / n
    big_h = Inches(0.55) if big_size <= 22 else Inches(0.7)
    for i, (big, label) in enumerate(items):
        x = left + i * (cw + gap)
        add_rect(slide, x, top, cw, height, TEAL_LIGHT)
        add_rect(slide, x, top, Inches(0.08), height, TEAL)
        add_text(slide, x + Inches(0.15), top + Inches(0.06), cw - Inches(0.2), big_h, big,
                 size=big_size, bold=True, color=NAVY_2, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, x + Inches(0.15), top + big_h + Inches(0.05), cw - Inches(0.2),
                 height - big_h - Inches(0.1), label, size=label_size, color=GRAY, line_spacing=1.05)


def add_callout(slide, text, top=Inches(6.2), left=MARGIN, width=CONTENT_W, height=Inches(0.6),
                fill=AMBER_LIGHT, bar=AMBER, size=13.5, color=NAVY_2):
    add_rect(slide, left, top, width, height, fill)
    add_rect(slide, left, top, Inches(0.08), height, bar)
    add_text(slide, left + Inches(0.18), top, width - Inches(0.25), height, text,
             size=size, bold=True, color=color, anchor=MSO_ANCHOR.MIDDLE)


# =====================================================================
# 0. 封面 & 提纲
# =====================================================================
def cover():
    _state["page"] += 1
    s = prs.slides.add_slide(BLANK)
    add_rect(s, 0, 0, W, H, NAVY)
    add_rect(s, 0, Inches(5.2), W, Inches(2.3), NAVY_2)
    # 装饰几何
    for i in range(6):
        add_rect(s, W - Inches(3.2) + i * Inches(0.45), Inches(0.5) + i * Inches(0.3),
                 Inches(0.3), Inches(0.3), TEAL if i % 2 == 0 else AMBER,
                 shape=MSO_SHAPE.OVAL)
    add_text(s, MARGIN, Inches(1.2), Inches(9), Inches(0.5), "学术报告", size=16, color=TEAL)
    add_text(s, MARGIN, Inches(1.75), Inches(12), Inches(1.6),
             "先进制造装备关键零部件\n研发及前沿技术", size=44, bold=True, color=WHITE,
             line_spacing=1.1)
    add_text(s, MARGIN, Inches(3.75), Inches(12), Inches(0.5),
             "Key Components of Advanced Manufacturing Equipment: R&D Status and Frontier Technologies",
             size=15, color=STEEL, italic=True)
    add_rect(s, MARGIN, Inches(4.45), Inches(1.6), Inches(0.06), AMBER)
    add_text(s, MARGIN, Inches(5.55), Inches(8), Inches(0.45), "报告人：________        单位：________",
             size=16, color=WHITE)
    add_text(s, MARGIN, Inches(6.05), Inches(8), Inches(0.45),
             "时长约 100 分钟（含 10 分钟讨论）    面向：高校与科研院所研究人员、研究生", size=13, color=TEAL_LIGHT)
    add_text(s, MARGIN, Inches(6.55), Inches(8), Inches(0.4), "2026 年", size=13, color=STEEL)
    add_notes(s, "1 分钟", "自我介绍、报告背景。说明报告面向学术同行，侧重科学问题与技术路线，而非产品介绍。")


def agenda():
    s = content_slide("报告提纲与时间安排", "五个部分 · 约 90 分钟报告 + 10 分钟讨论",
                      notes=("1.5 分钟", "简述五部分逻辑：从宏观背景到具体零部件，再到共性技术与前沿，最后给出趋势与研究建议。"))
    top = CONTENT_TOP + Inches(0.1)
    for i, (num, title, dur, _) in enumerate(SECTIONS):
        y = top + i * Inches(0.98)
        add_rect(s, MARGIN, y, Inches(0.9), Inches(0.8), TEAL if i % 2 == 0 else NAVY_2)
        add_text(s, MARGIN, y, Inches(0.9), Inches(0.8), num, size=26, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, MARGIN + Inches(0.9), y, Inches(8.6), Inches(0.8), GRAY_LIGHT)
        add_text(s, MARGIN + Inches(1.1), y, Inches(8.3), Inches(0.8), title, size=19, bold=True,
                 color=NAVY_2, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, MARGIN + Inches(9.5), y, Inches(2.7), Inches(0.8), WHITE, line=GRAY_MID)
        add_text(s, MARGIN + Inches(9.5), y, Inches(2.7), Inches(0.8), dur, size=16, color=AMBER,
                 bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_callout(s, "报告约定：聚焦共性科学问题与技术路线；数据为公开资料整理，具体数值请以最新统计为准。",
                top=Inches(6.35), height=Inches(0.5))


# =====================================================================
# 第一部分 绪论
# =====================================================================
def part1():
    section_slide(0, "本部分回答三个问题：什么是先进制造装备与关键零部件？为什么零部件决定装备？我们处于什么位置？")

    # 1.1 先进制造装备谱系
    s = content_slide("先进制造装备的内涵与谱系", "装备是制造业的“工作母机”，决定一个国家制造能力的上限",
                      notes=("2 分钟", "强调装备的层次性：整机—功能部件—基础零部件。举例说明不同装备门类对零部件需求的差异（机床重精度刚度，机器人重减速器，半导体装备重超精密运动与真空）。"))
    add_cards(s, [
        ("高档数控机床", "五轴联动加工中心、精密磨床、超精密车床\n核心：主轴、丝杠导轨、转台、数控系统"),
        ("工业机器人", "六轴串联、SCARA、协作机器人\n核心：减速器、伺服电机、控制器、力传感器"),
        ("半导体与泛半导体装备", "光刻机、刻蚀机、薄膜沉积、检测设备\n核心：超精密工件台、真空机械手、静电卡盘"),
        ("增材制造装备", "激光选区熔化、电子束熔融、定向能量沉积\n核心：振镜、激光器、铺粉系统、过程监测"),
        ("航空航天专用装备", "大型蒙皮镜像铣、复材铺放、精密装配\n核心：大行程高刚度运动系统、测量定位"),
        ("智能产线与检测装备", "柔性产线、在线检测、物流AGV\n核心：传感器、视觉系统、工业总线"),
    ], cols=3, body_size=12.5, height=Inches(4.5))
    add_callout(s, "共性特征：高精度 · 高速高动态 · 高可靠 · 高集成 · 极端工况 —— 均最终落脚于关键零部件的性能",
                top=Inches(6.1), height=Inches(0.6))

    # 1.2 关键零部件定义与分类
    s = content_slide("关键零部件：定义、分类与“工业四基”", "核心基础零部件是“四基”之首，是装备性能与可靠性的物理载体",
                      notes=("2 分钟", "解释“四基”概念（核心基础零部件/元器件、关键基础材料、先进基础工艺、产业技术基础），说明本报告聚焦零部件，但材料与工艺是研发不可分割的部分。"))
    add_bullets(s, MARGIN, CONTENT_TOP, Inches(6.2), Inches(5.3), [
        "**定义**：对装备的精度、速度、可靠性、寿命起决定作用，且技术门槛高、替代困难的基础部件",
        "**分类（按功能）**",
        ("支承类：精密轴承、气浮/磁悬浮支承、导轨", 1),
        ("传动类：滚珠丝杠、精密减速器、联轴器、丝杠螺母", 1),
        ("驱动类：电主轴、伺服电机、直线/力矩电机、液压气动元件", 1),
        ("控制类：数控系统、伺服驱动器、运动控制器", 1),
        ("感知类：光栅、编码器、测头、力/力矩传感器", 1),
        ("工艺执行类：刀具、夹具、密封、润滑与冷却单元", 1),
        "**特点**：多学科交叉（材料—力学—摩擦学—控制—测量），研发周期长、试错成本高、依赖工程积累",
    ], size=16)
    # 右侧：四基金字塔
    x0, y0 = Inches(7.3), CONTENT_TOP + Inches(0.2)
    layers = [("产业技术基础", STEEL, Inches(5.2)), ("先进基础工艺", GREEN, Inches(4.4)),
              ("关键基础材料", AMBER, Inches(3.6)), ("核心基础零部件", TEAL, Inches(2.8))]
    for i, (name, col, wdt) in enumerate(reversed(layers)):
        yy = y0 + i * Inches(1.05)
        cx = x0 + (Inches(5.2) - wdt) / 2
        b = add_rect(s, cx, yy, wdt, Inches(0.9), col, shape=MSO_SHAPE.TRAPEZOID if i == 0 else MSO_SHAPE.RECTANGLE)
        shape_text(b, name, size=15, bold=True)
    add_text(s, x0, y0 + Inches(4.35), Inches(5.2), Inches(0.6),
             "“工业四基”：中国制造2025 工业强基工程\n产业基础再造工程（2021—）延续并深化", size=12, color=GRAY,
             align=PP_ALIGN.CENTER)

    # 1.3 性能传递链
    s = content_slide("为什么零部件决定装备？—— 性能传递链与成本结构",
                      "零部件误差经运动链放大、叠加，整机精度和可靠性由“短板”决定",
                      notes=("2 分钟", "以机床为例讲误差传递：轴承跳动→主轴回转误差→加工表面质量；丝杠导程误差→定位误差。用木桶效应解释可靠性。成本占比数据为行业调研的大致区间。"))
    add_chevrons(s, ["材料与热处理", "零件精度", "部件装配精度", "整机几何/动态精度", "加工件精度与一致性"],
                 top=CONTENT_TOP, height=Inches(0.85))
    add_text(s, MARGIN, CONTENT_TOP + Inches(0.9), CONTENT_W, Inches(0.4),
             "误差源  →  传递与放大（阿贝误差、热变形、刚度链）  →  整机表现", size=12.5, color=GRAY, align=PP_ALIGN.CENTER)
    add_kpis(s, [
        ("60–70%", "高档数控机床中，功能部件（数控系统、主轴、丝杠导轨、转台、刀库等）占整机成本比例的大致区间"),
        ("~70%", "工业机器人成本中减速器、伺服系统、控制器三大核心部件合计占比（行业调研常见区间）"),
        ("1/R", "可靠性串联模型：整机 MTBF 受最薄弱部件制约，任一环节失效率上升都会显著拉低整机"),
    ], top=CONTENT_TOP + Inches(1.45), height=Inches(1.55))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(3.25), CONTENT_W, Inches(2.2), [
        "**精度**：主轴回转误差、滚珠丝杠导程误差、导轨直线度直接映射为工件形位误差",
        "**动态性能**：主轴/进给系统刚度与阻尼决定颤振边界与可达材料去除率",
        "**可靠性**：轴承疲劳、密封失效、电子元件退化是机床与机器人现场故障的主要来源",
        "**一致性**：批量零部件的性能离散度决定整机“标称性能”能否稳定复现 —— 这是产业化最大难点",
    ], size=15)

    # 1.4 现状与差距
    s = content_slide("现状与差距：高端零部件仍是“卡脖子”集中区",
                      "中低端基本实现自主，高端产品在精度保持性、寿命与一致性上仍存在代差",
                      notes=("2.5 分钟", "逐行讲解。强调：差距主要不在“做不出来”，而在于“精度保持性、批量一致性、寿命”——这些背后是材料、工艺与工程数据积累。数据为公开行业报告区间，提醒听众自行核实最新数据。"))
    add_table(s, [
        ["零部件", "国际领先企业", "国内代表", "主要差距（高端段）"],
        ["精密轴承", "SKF、Schaeffler、NSK、NTN、JTEKT、Timken", "洛轴、瓦轴、哈轴、人本", "钢材纯净度、热处理一致性、超精研、寿命与噪声"],
        ["滚珠丝杠/导轨", "THK、NSK、HIWIN、Bosch Rexroth、Schneeberger", "南京工艺、汉江机床、博特精工", "C0–C3 级大导程/长行程产品，精度保持性"],
        ["精密减速器", "Nabtesco（RV）、Harmonic Drive（谐波）", "绿的谐波、双环传动、中大力德", "RV 寿命与回差稳定性；谐波已形成一定替代"],
        ["数控系统", "FANUC、Siemens、Heidenhain、Mitsubishi", "华中数控、广州数控、科德数控", "五轴高端系统生态、工艺软件、可靠性口碑"],
        ["电主轴", "Kessler、Fischer、IBAG、Step-Tec、GMN", "洛阳轴研、昊志机电", "高速大功率段、热稳定性、寿命"],
        ["液压元件", "Bosch Rexroth、Moog、Parker、Kawasaki", "恒立液压、艾迪精密、航空工业609所", "高压柱塞泵、电液伺服阀的寿命与污染敏感性"],
        ["光栅/编码器", "Heidenhain、Renishaw、Magnescale", "禹衡光学、长春光机所体系", "纳米级绝对式光栅、高速高分辨率绝对编码器"],
    ], col_widths=[1.5, 3.4, 2.6, 4.2], font_size=11, row_h=Inches(0.56))
    add_callout(s, "结论：差距的本质是“材料—工艺—检测—数据”的工程积累不足，而非单点原理突破 —— 这恰恰是学术界可以系统发力之处",
                top=Inches(6.3), height=Inches(0.5), size=13)

    # 1.5 战略背景
    s = content_slide("战略背景与政策脉络", "国家层面持续、系统地支持基础零部件攻关",
                      notes=("1.5 分钟", "快速过时间轴，重点说明04专项和产业基础再造工程的定位差异：前者以整机带动部件，后者直接针对“四基”。"))
    tl = [
        ("2009", "“高档数控机床与基础制造装备”科技重大专项（04 专项）启动，2009—2020"),
        ("2015", "《中国制造 2025》发布，“工业强基工程”列为五大工程之一，提出核心基础零部件自主保障目标"),
        ("2016", "工信部等发布《工业强基工程实施指南（2016—2020）》，“一揽子”突破与“一条龙”应用"),
        ("2021", "“十四五”规划提出实施产业基础再造工程；《“十四五”智能制造发展规划》发布"),
        ("2023", "新一轮大规模设备更新与高端装备自主可控政策，强调核心零部件补短板"),
        ("2025+", "“十五五”时期：从补短板走向锻长板，前沿技术布局（智能零部件、极端制造）"),
    ]
    y_axis = CONTENT_TOP + Inches(1.6)
    add_line(s, MARGIN, y_axis, W - MARGIN, y_axis, color=TEAL, width=3)
    n = len(tl)
    step = CONTENT_W / n
    for i, (yr, desc) in enumerate(tl):
        x = MARGIN + i * step
        add_rect(s, x + step / 2 - Inches(0.14), y_axis - Inches(0.14), Inches(0.28), Inches(0.28),
                 AMBER if i % 2 else TEAL, shape=MSO_SHAPE.OVAL)
        add_text(s, x, y_axis - Inches(0.85), step, Inches(0.6), yr, size=20, bold=True, color=NAVY_2,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)
        add_rect(s, x + Inches(0.08), y_axis + Inches(0.35), step - Inches(0.16), Inches(2.4), GRAY_LIGHT)
        add_text(s, x + Inches(0.14), y_axis + Inches(0.42), step - Inches(0.28), Inches(2.3), desc,
                 size=11.5, color=BLACK, line_spacing=1.1)
    add_callout(s, "政策共识：以应用牵引（整机—部件“一条龙”）+ 基础研究（材料、工艺、可靠性）双轮驱动",
                top=Inches(6.1), height=Inches(0.6))

    # 1.6 共性技术特征
    s = content_slide("关键零部件的共性技术特征", "六个维度构成零部件研发的“目标函数”，彼此耦合甚至矛盾",
                      notes=("2 分钟", "强调多目标耦合：高速与长寿命矛盾、高刚度与低发热矛盾、高精度与低成本矛盾。研发本质是在约束下寻优。"))
    add_cards(s, [
        ("精度与精度保持性", "μm → sub-μm → nm 量级\n几何精度 / 运动精度 / 重复精度\n关键：精度随时间、温度、载荷的保持能力"),
        ("刚度与动态特性", "静刚度、动刚度、阻尼、固有频率\n决定颤振边界、轨迹跟踪误差\n关键：轻量化与高刚度的权衡"),
        ("热稳定性", "热源：轴承摩擦、电机损耗、切削热\n热误差可占机床总误差 40–70%\n关键：热对称设计、冷却、热误差补偿"),
        ("可靠性与寿命", "L10 寿命、MTBF、失效率曲线\n失效机理：疲劳、磨损、腐蚀、电子退化\n关键：加速试验与寿命预测模型"),
        ("一致性与可制造性", "批量性能离散度（Cpk）\n工艺窗口、公差链、装配工艺性\n关键：从“样机好”到“批量好”"),
        ("极端工况适应性", "高速、高温、真空、洁净、强辐射\n特殊润滑、材料与密封\n关键：面向半导体、航天、核能等场景"),
    ], cols=3, body_size=12.5, height=Inches(4.6))
    add_callout(s, "研发的核心矛盾：高速 vs 寿命、高刚度 vs 低发热、高精度 vs 低成本 —— 需要系统级多目标优化思维",
                top=Inches(6.15), height=Inches(0.55))

    # 1.7 研发范式
    s = content_slide("关键零部件研发的典型范式：V 模型与正向设计",
                      "从“逆向测绘—仿制”走向“需求驱动—正向设计—试验验证—数据迭代”",
                      notes=("2 分钟", "讲 V 模型左侧（需求分解到零件设计）和右侧（零件验证到整机验证）。强调试验数据库与设计知识库的闭环是国外企业最大的壁垒。"))
    left_steps = ["装备需求与指标分解", "部件系统设计", "零件设计与材料选择", "工艺设计与制造"]
    right_steps = ["零件检测与评价", "部件台架试验", "整机集成验证", "现场数据反馈"]
    x_l, x_r = MARGIN + Inches(0.3), W - MARGIN - Inches(3.6)
    y0 = CONTENT_TOP + Inches(0.1)
    for i, (l, r) in enumerate(zip(left_steps, right_steps)):
        yy = y0 + i * Inches(1.05)
        bl = add_rect(s, x_l + i * Inches(0.75), yy, Inches(3.3), Inches(0.8), NAVY_2)
        shape_text(bl, l, size=14, bold=True)
        br = add_rect(s, x_r - i * Inches(0.75), yy, Inches(3.3), Inches(0.8), TEAL)
        shape_text(br, r, size=14, bold=True)
        add_line(s, x_l + i * Inches(0.75) + Inches(3.3), yy + Inches(0.4),
                 x_r - i * Inches(0.75), yy + Inches(0.4), color=GRAY_MID, dash=True)
    base = add_rect(s, (W - Inches(3.4)) / 2, y0 + Inches(4.3), Inches(3.4), Inches(0.8), AMBER)
    shape_text(base, "制造 / 装配 / 调试", size=14, bold=True, color=NAVY)
    add_text(s, MARGIN, y0 + Inches(5.2), CONTENT_W, Inches(0.5),
             "闭环的关键：试验数据库 + 设计知识库 + 工艺参数库  →  这是先进企业数十年积累形成的核心壁垒",
             size=13.5, bold=True, color=NAVY_2, align=PP_ALIGN.CENTER)


# =====================================================================
# 第二部分 核心基础零部件
# =====================================================================
def part2():
    section_slide(1, "本部分逐类剖析八大类零部件：原理与关键指标 → 研发难点 → 国内外水平 → 研究方向。每类约 4 分钟。")

    # 2.1 轴承总览
    s = content_slide("2.1 高端轴承：装备的“关节”", "分类、关键指标与技术门槛",
                      notes=("2 分钟", "轴承是最典型的“看似简单实则极难”的零部件。先讲分类，再讲为什么难：几十个微米级公差要素同时保证，且要批量一致。"))
    add_cards(s, [
        ("滚动轴承", "角接触球轴承、圆柱滚子、圆锥滚子、交叉滚子\n应用：主轴、丝杠支承、转台、机器人关节\n关键：精度等级 P4/P2、DN 值、预紧、寿命"),
        ("滑动/流体轴承", "液体静压、动静压、气体静压、箔片轴承\n应用：超精密机床、高速透平、磨床主轴\n关键：间隙精度（μm）、刚度、稳定性"),
        ("磁悬浮轴承", "主动电磁、被动永磁、混合式\n应用：涡轮分子泵、高速电机、飞轮储能\n关键：控制带宽、功率放大器、容错"),
        ("特种轴承", "陶瓷/混合陶瓷、薄壁、真空/低温、固体润滑\n应用：航天、半导体、核能\n关键：材料、润滑、极端环境可靠性"),
    ], cols=4, body_size=12, height=Inches(2.9))
    add_table(s, [
        ["关键指标", "典型要求（高端主轴轴承）", "物理意义与难点"],
        ["旋转精度", "P4 / P2 级，径向跳动 ≤ 1–2 μm", "套圈沟道圆度、波纹度、滚动体直径一致性（亚微米）"],
        ["极限转速（DN 值）", "钢球 1.5–2.0×10⁶，陶瓷球 2.5–3.0×10⁶ mm·r/min", "离心力与发热随转速平方增长，润滑膜形成"],
        ["刚度与预紧", "预紧力精确可控，温升后预紧漂移小", "预紧—刚度—发热—寿命强耦合"],
        ["寿命与可靠性", "L10 寿命数万小时，噪声/振动等级低", "材料夹杂物、表面完整性、润滑脂寿命"],
    ], top=CONTENT_TOP + Inches(3.1), col_widths=[1.6, 3.6, 4.4], font_size=11.5, row_h=Inches(0.46))

    # 2.1b 精密主轴轴承
    s = content_slide("2.1 精密主轴轴承：材料—工艺—润滑的协同", "以高速角接触球轴承为例",
                      notes=("2.5 分钟", "重点讲三件事：材料（轴承钢纯净度、陶瓷球）、工艺（超精研与波纹度控制）、润滑（油气润滑）。说明为何国产轴承“装上能转”但“跑不长、精度掉得快”。"))
    add_bullets(s, MARGIN, CONTENT_TOP, Inches(6.4), Inches(5.4), [
        "**材料体系**",
        ("套圈：GCr15/GCr15SiMn、Cr4Mo4V（高温）、渗碳钢；氧含量 ≤ 5–8 ppm，夹杂物评级控制", 1),
        ("滚动体：Si₃N₄ 陶瓷球（密度为钢的 40%，弹性模量高 50%），降低离心力与发热", 1),
        ("保持架：酚醛层压布管、PEEK、黄铜；高速下的变形与引导方式", 1),
        "**制造工艺**",
        ("沟道超精研：圆度 ≤ 0.2 μm，波纹度与粗糙度 Ra ≤ 0.02 μm", 1),
        ("热处理：真空/可控气氛淬火、深冷处理、残余奥氏体控制与尺寸稳定性", 1),
        ("选配与成组：接触角、沟位置差控制，成组轴承预紧一致性", 1),
        "**润滑与热管理**",
        ("油气润滑、油雾润滑；脂润滑寿命模型；外圈冷却套", 1),
        ("轴承—主轴—壳体热—结构耦合：预紧随温升变化是寿命杀手", 1),
    ], size=15)
    add_kpis(s, [("≤ 2 μm", "P2 级主轴轴承径向跳动"),
                 ("3.0×10⁶", "陶瓷球混合轴承 DN 值上限"),
                 ("≤ 5 ppm", "高端轴承钢氧含量目标")],
             left=Inches(7.3), width=Inches(5.5), top=CONTENT_TOP, height=Inches(1.4),
             big_size=20, label_size=10.5, gap=Inches(0.12))
    add_rect(s, Inches(7.3), CONTENT_TOP + Inches(1.65), Inches(5.5), Inches(3.75), GRAY_LIGHT)
    add_text(s, Inches(7.45), CONTENT_TOP + Inches(1.75), Inches(5.2), Inches(0.4), "研究方向（学术视角）",
             size=14, bold=True, color=NAVY_2)
    add_bullets(s, Inches(7.4), CONTENT_TOP + Inches(2.15), Inches(5.3), Inches(3.2), [
        "多体动力学 + 弹流润滑（EHL）+ 热的耦合建模，预测高速下保持架稳定性与打滑",
        "滚动接触疲劳机理：白蚀区（WEA）、夹杂物起裂的多尺度模拟",
        "智能轴承：嵌入式温度/振动/载荷感知与预紧主动调控",
        "面向批量的工艺—性能映射与数字孪生质量控制",
    ], size=12.5, para_space=4)

    # 2.1c 气浮与磁悬浮
    s = content_slide("2.1 非接触支承：气体静压轴承与磁悬浮轴承", "超精密与超高速装备的首选",
                      notes=("2 分钟", "对比两种非接触支承。气浮：精度极高但刚度低、需洁净气源；磁悬浮：可主动控制但系统复杂。半导体装备（光刻机工件台）大量采用气浮。"))
    add_table(s, [
        ["对比项", "气体静压轴承", "主动磁悬浮轴承（AMB）"],
        ["工作原理", "外部供气在间隙（5–20 μm）形成压力膜", "电磁力悬浮，位移传感器+控制器闭环"],
        ["回转精度", "可达 ≤ 50 nm（误差均化效应）", "亚微米级，取决于传感与控制"],
        ["刚度/承载", "刚度中低（数十–数百 N/μm），承载有限", "可调，承载受电磁饱和限制"],
        ["速度/摩擦", "无磨损，摩擦极低；高速易产生气锤自激", "无接触，适合超高速（>100 000 r/min）"],
        ["典型应用", "超精密机床主轴/导轨、光刻机工件台、三坐标测量机", "涡轮分子泵、离心压缩机、高速电主轴、飞轮储能"],
        ["研发难点", "节流器微结构加工、气膜稳定性、供气洁净度", "控制带宽、功放、传感器噪声、断电保护轴承"],
    ], col_widths=[1.5, 4.0, 4.0], font_size=12, row_h=Inches(0.52), top=CONTENT_TOP)
    add_callout(s, "前沿：多孔质石墨节流气浮、气磁混合支承、无传感器（自感知）磁轴承、基于模型预测控制的 AMB",
                top=Inches(5.3), height=Inches(0.55))
    add_text(s, MARGIN, Inches(5.95), CONTENT_W, Inches(0.9),
             "科学问题：气膜微尺度流动的稳定性与非线性（气锤振荡）；磁悬浮转子—控制器—电磁系统的多场耦合与鲁棒性；非接触支承的误差均化机理与极限精度。",
             size=13, color=GRAY, line_spacing=1.1)

    # 2.1d 研发链路与差距
    s = content_slide("2.1 轴承研发全链条：差距在每一环节的“最后 10%”", "材料 → 热处理 → 加工 → 检测 → 试验 → 应用",
                      notes=("1.5 分钟", "用链条图说明每一环节与国际先进水平的差距点，强调它们是乘法关系而不是加法。"))
    add_chevrons(s, ["特种冶金", "热处理", "超精加工", "精密检测", "寿命试验", "装机应用"], height=Inches(0.85))
    gaps = [
        "钢材纯净度与均匀性\n大尺寸夹杂物概率\n陶瓷球烧结缺陷",
        "变形控制与尺寸稳定性\n深冷/回火工艺参数\n批次一致性",
        "沟道波纹度与表面完整性\n磨削烧伤与残余应力\n超精研机床本身精度",
        "波纹度仪、圆度仪\n接触角/沟位置差测量\n检测标准与数据追溯",
        "台架覆盖度与加速因子\n失效物理与寿命模型\n数据库规模（百万小时级）",
        "主机厂配套验证机会\n现场工况数据反馈\n“不敢用—用不好—不好用”循环",
    ]
    cw = CONTENT_W / 6
    for i, g in enumerate(gaps):
        x = MARGIN + i * cw
        add_rect(s, x + Inches(0.05), CONTENT_TOP + Inches(1.05), cw - Inches(0.1), Inches(2.5), GRAY_LIGHT)
        add_text(s, x + Inches(0.1), CONTENT_TOP + Inches(1.12), cw - Inches(0.2), Inches(2.4), g,
                 size=11.5, color=BLACK, line_spacing=1.15)
    add_callout(s, "破局路径：建立“材料—工艺—性能”全链条数据库；以应用验证平台（整机厂—零部件厂—高校）打通“首台套”困境",
                top=Inches(5.3), height=Inches(0.6))
    add_text(s, MARGIN, Inches(6.0), CONTENT_W, Inches(0.9),
             "案例思路：轴承试验数据库规模是先进企业最核心的资产之一；国内团队可通过标准化试验台 + 共享数据平台的方式，在学术合作框架内实现规模效应。",
             size=12.5, color=GRAY)

    # 2.2 滚珠丝杠与导轨
    s = content_slide("2.2 精密传动（一）：滚珠丝杠副与滚动直线导轨副", "进给系统精度与动态性能的决定环节",
                      notes=("2.5 分钟", "先讲指标体系（精度等级、预压、刚度、DN 值），再讲工艺难点（螺纹磨削、滚道型面、热伸长）。强调国产在中档已有竞争力，高档长行程、大导程、高速是难点。"))
    add_bullets(s, MARGIN, CONTENT_TOP, Inches(6.3), Inches(5.3), [
        "**滚珠丝杠副：关键指标**",
        ("精度等级：C0–C7（JIS）/ 1–7 级（ISO 3408）；C0 级行程偏差约 ±3.5 μm/300 mm", 1),
        ("预压与刚度：双螺母/变导程预压，轴向刚度与摩擦力矩的权衡", 1),
        ("DN 值与高速：≥ 1.5×10⁵ 时需端盖/内循环返向器优化，噪声与温升控制", 1),
        ("热伸长：1 m 丝杠温升 1 ℃ 伸长约 12 μm → 预拉伸 + 中空冷却 + 补偿", 1),
        "**滚动直线导轨副：关键指标**",
        ("精度：行走平行度（μm 级），预压等级与刚度；四列/两列圆弧沟槽接触形式", 1),
        ("寿命：滚道硬度（HRC 58–62）、表面完整性；返向器与保持链设计", 1),
        "**研发难点**",
        ("螺纹/滚道磨削机床与砂轮修整；滚道型面（哥特式弧）轮廓精度", 1),
        ("长行程（> 3 m）丝杠的直线度、热处理变形与精度保持性", 1),
    ], size=14.5, para_space=5)
    add_table(s, [
        ["方向", "研究内容"],
        ["滚动体动力学", "滚珠—滚道—返向器接触碰撞建模，噪声与振动机理"],
        ["摩擦与磨损", "混合润滑下摩擦力矩预测，磨损寿命与预压衰减"],
        ["热误差", "丝杠热伸长实时建模（热成像/温度场重构）与补偿"],
        ["智能化", "内置传感的“智能丝杠”：预压监测、润滑状态与剩余寿命"],
        ["新原理", "行星滚柱丝杠（高载荷密度）、静压丝杠（超精密）、磁力丝杠", ],
    ], left=Inches(7.1), width=Inches(5.7), top=CONTENT_TOP, col_widths=[1.3, 4.4], font_size=11.5,
        row_h=Inches(0.62))
    add_callout(s, "指标提示：高速加工中心进给速度 60–120 m/min、加速度 1–2 g，对丝杠导轨提出高速低温升的双重要求",
                left=Inches(7.1), width=Inches(5.7), top=CONTENT_TOP + Inches(3.95), height=Inches(1.3), size=12)

    # 2.2b 减速器
    s = content_slide("2.2 精密传动（二）：机器人精密减速器", "RV、谐波、行星三类减速器的对比与选型逻辑",
                      notes=("2 分钟", "讲三类减速器的性能特点与应用分工。RV 用于大负载基座关节，谐波用于腕部和协作机器人。强调回差、传动精度、刚度、寿命是核心指标。"))
    add_table(s, [
        ["指标", "RV 减速器", "谐波减速器", "精密行星减速器"],
        ["传动原理", "渐开线行星 + 摆线针轮两级", "柔轮弹性变形与刚轮错齿啮合", "多级行星轮系"],
        ["典型速比", "30–200", "50–320", "3–100（多级）"],
        ["回差 / 传动精度", "≤ 1 arcmin；传动误差 ≤ 1 arcmin", "近零回差；传动误差 ≤ 1–1.5 arcmin", "1–15 arcmin"],
        ["刚度 / 承载", "高刚度，重载（基座、大臂）", "低刚度，中小负载（腕部、协作）", "中等"],
        ["效率", "80–90%", "70–85%", "90–97%"],
        ["寿命", "6000 h 以上（额定）", "7000–10000 h（额定）", "较长"],
        ["国际/国内代表", "Nabtesco / 双环传动、中大力德", "Harmonic Drive / 绿的谐波、来福", "Wittenstein、Neugart / 多家"],
        ["主要难点", "摆线轮修形与针齿精度、装配回差控制、寿命", "柔轮材料疲劳、齿形设计、润滑脂", "齿轮精度、热处理变形"],
    ], col_widths=[1.6, 3.2, 3.2, 2.6], font_size=11.5, row_h=Inches(0.5))
    add_callout(s, "趋势：减速器—电机—编码器—驱动器一体化关节模组；轻量化（铝合金/复合材料）；面向人形机器人的高扭矩密度执行器",
                top=Inches(6.0), height=Inches(0.65), size=12.5)

    # 2.2c 减速器关键技术
    s = content_slide("2.2 减速器研发中的关键科学与工程问题", "从齿形设计到装配回差，每一环节都需要“理论 + 工艺 + 检测”",
                      notes=("2 分钟", "重点讲摆线轮修形、柔轮齿形和疲劳、以及精密装配检测。说明谐波减速器国产替代已较成功，RV 仍有差距，原因在于零件数量多、公差链长、装配选配要求高。"))
    add_cards(s, [
        ("齿形设计与修形", "RV：摆线轮等距/移距/转角修形的多目标优化，啮合刚度与传动误差\n谐波：双圆弧、IH 齿形、共轭齿形的空间啮合理论\n工具：齿面接触分析（TCA）、载荷接触分析（LTCA）"),
        ("材料与热处理", "摆线轮/针齿：轴承钢或渗碳钢，高硬度高耐磨\n柔轮：高强韧合金钢（如 40CrNiMoA 系），疲劳极限决定寿命\n挑战：薄壁柔轮热处理变形、残余应力控制"),
        ("精密制造", "摆线轮成形磨削（轮廓精度 ≤ 2–3 μm）\n针齿壳孔位精度、柔轮薄壁车削/滚齿变形\n柔性轴承（薄壁交叉滚子）本身是高端轴承"),
        ("装配与检测", "选配装配控制回差、预紧量；装配应力\n传动误差、回差、启动力矩、刚度曲线测试\n加速寿命台架与失效分析（点蚀、断齿、柔轮裂纹）"),
    ], cols=2, body_size=12.5, height=Inches(4.2))
    add_callout(s, "科学问题：多齿啮合的载荷分配与误差敏感性；柔轮多轴疲劳寿命预测；考虑制造误差分布的公差—性能概率设计",
                top=Inches(5.75), height=Inches(0.6), size=13)

    # 2.3 电主轴
    s = content_slide("2.3 高速电主轴：机电热一体化的典型", "“零传动”的核心功能部件，决定加工效率与表面质量",
                      notes=("2.5 分钟", "讲结构组成、主要指标（转速—功率—刚度三角）、研发难点（热、动平衡、轴承寿命、拉刀）。可以结合图示说明电机内置带来的热问题。"))
    # 结构示意
    x0, y0 = MARGIN, CONTENT_TOP
    add_text(s, x0, y0, Inches(6.2), Inches(0.35),
             "结构示意：壳体（冷却水套）— 轴承组 — 内置电机 — 主轴 — 拉刀机构 — 编码器 — 密封",
             size=11, color=GRAY)
    add_rect(s, x0, y0 + Inches(0.5), Inches(6.0), Inches(1.3), GRAY_MID)  # 壳体
    add_rect(s, x0 + Inches(0.3), y0 + Inches(1.0), Inches(5.4), Inches(0.3), STEEL)  # 轴
    for xx in (x0 + Inches(0.5), x0 + Inches(1.0), x0 + Inches(4.6), x0 + Inches(5.1)):
        add_rect(s, xx, y0 + Inches(0.65), Inches(0.35), Inches(1.0), TEAL)
    mot = add_rect(s, x0 + Inches(2.0), y0 + Inches(0.6), Inches(2.0), Inches(1.1), AMBER)
    shape_text(mot, "内置电机\n定子/转子", size=11, bold=True, color=NAVY)
    add_text(s, x0 + Inches(0.3), y0 + Inches(1.85), Inches(1.6), Inches(0.35), "前轴承组", size=11, color=NAVY_2)
    add_text(s, x0 + Inches(4.4), y0 + Inches(1.85), Inches(1.6), Inches(0.35), "后轴承组", size=11, color=NAVY_2)
    add_text(s, x0 + Inches(1.9), y0 + Inches(1.85), Inches(2.5), Inches(0.35), "冷却水套 / 油气润滑 / 气封", size=11, color=GRAY, align=PP_ALIGN.CENTER)
    add_table(s, [
        ["指标", "典型值（加工中心）", "难点"],
        ["最高转速", "12 000–40 000 r/min；PCB 钻孔 > 100 000", "轴承 DN 值、转子动力学"],
        ["功率/扭矩", "10–80 kW；低速大扭矩 vs 高速恒功率", "电磁设计、散热"],
        ["动平衡", "G0.4 级（ISO 21940）", "整体动平衡、热态平衡"],
        ["热漂移", "轴向伸长 ≤ 10–20 μm（稳态）", "热源分布、冷却、补偿"],
    ], left=MARGIN, width=Inches(6.2), top=y0 + Inches(2.35), col_widths=[1.2, 3.2, 1.8], font_size=10.5, row_h=Inches(0.44))
    add_bullets(s, Inches(7.1), CONTENT_TOP, Inches(5.7), Inches(5.4), [
        "**研发难点**",
        ("热—结构—转子动力学耦合：电机损耗与轴承摩擦热导致预紧变化、临界转速漂移", 1),
        ("高速轴承寿命与润滑：油气润滑参数、陶瓷球、保持架稳定性", 1),
        ("拉刀机构与接口（HSK/BT/CAPTO）：高速下的刀柄接触刚度", 1),
        ("电机驱动：高频（> 1 kHz）矢量控制、弱磁扩速、谐波损耗", 1),
        "**国内外格局**",
        ("Kessler、Fischer、IBAG、Step-Tec、GMN、Weiss 占据高端；国内在中低速、PCB 钻孔主轴已形成规模", 1),
        "**研究方向**",
        ("主轴数字孪生：实时热误差与刚度预测", 1),
        ("智能主轴：内置振动/温度/位移传感，颤振在线识别与主动抑制", 1),
        ("磁悬浮电主轴、超高速（> 200 000 r/min）微主轴", 1),
    ], size=13.5, para_space=4)

    # 2.3b 直驱
    s = content_slide("2.3 直驱技术：直线电机、力矩电机与 DD 转台", "取消中间传动环节，实现高动态、高精度",
                      notes=("2 分钟", "对比直驱与传统传动的优劣。直驱的核心挑战是推力波动（齿槽力）、发热、刚度（无减速比放大）与对控制的高要求。"))
    add_table(s, [
        ["", "滚珠丝杠进给", "直线电机进给", "力矩电机 / DD 转台"],
        ["运动精度", "受丝杠导程误差、反向间隙影响", "直接由光栅闭环决定，可达亚微米", "圆光栅闭环，角秒级"],
        ["动态性能", "加速度 ~1 g，带宽受限于机械谐振", "加速度 2–10 g，带宽高", "高角加速度，无回差"],
        ["刚度", "机械刚度高（无减速比时）", "伺服刚度，依赖控制器增益", "伺服刚度"],
        ["热", "丝杠热伸长", "初级线圈发热直接传入工作台", "电机发热影响转台精度"],
        ["典型应用", "通用数控机床", "高速加工中心、光刻机、激光加工、PCB 钻孔", "五轴机床 A/C 轴、精密转台"],
        ["关键技术", "丝杠制造工艺", "齿槽力/推力波动抑制、磁路设计、冷却、抗干扰", "多极对数设计、转矩波动抑制、编码器"],
    ], col_widths=[1.3, 3.0, 3.6, 3.0], font_size=11.5, row_h=Inches(0.52))
    add_callout(s, "光刻机工件台：直线电机 + 气浮/磁浮导向 + 激光干涉/光栅测量，定位精度纳米级、加速度数 g —— 直驱技术的极致",
                top=Inches(5.75), height=Inches(0.55), size=13)
    add_text(s, MARGIN, Inches(6.35), CONTENT_W, Inches(0.6),
             "研究方向：无铁芯/永磁同步平面电机；磁悬浮平面电机（六自由度）；超导直线电机；基于学习的推力波动前馈补偿。",
             size=12.5, color=GRAY)

    # 2.4 数控系统
    s = content_slide("2.4 数控系统：装备的“大脑”", "软硬件结合的复杂系统，生态壁垒高于技术壁垒",
                      notes=("2.5 分钟", "讲架构、核心算法与生态。强调高端数控系统的差距在工艺软件生态、五轴高端功能可靠性以及用户信任，而不仅是插补算法。"))
    layers = [("人机界面 HMI / 工艺编程 / CAM 接口", STEEL), ("NC 内核：译码、刀补、前瞻、插补、加减速、误差补偿", TEAL),
              ("PLC：逻辑控制、辅助功能、安全", NAVY_2), ("伺服驱动与现场总线：EtherCAT / PROFINET / NCUC", GREEN),
              ("执行与感知：电机、光栅、编码器、I/O", AMBER)]
    for i, (t, c) in enumerate(layers):
        b = add_rect(s, MARGIN, CONTENT_TOP + i * Inches(0.78), Inches(6.2), Inches(0.68), c)
        shape_text(b, t, size=13, bold=True, color=WHITE if c != AMBER else NAVY)
    add_text(s, MARGIN, CONTENT_TOP + Inches(3.95), Inches(6.2), Inches(1.6),
             "硬件平台：多核 CPU + FPGA/实时内核；控制周期 125 μs–1 ms；多通道多轴（≥ 32 轴）\n"
             "软件平台：实时操作系统、开放式架构（PLCopen、OPC UA）、工艺软件包（五轴 RTCP、车铣复合、齿轮加工）",
             size=12, color=GRAY, line_spacing=1.15)
    add_bullets(s, Inches(7.1), CONTENT_TOP, Inches(5.7), Inches(5.5), [
        "**核心算法**",
        ("小线段/样条前瞻与速度规划（S 型、jerk 限制），NURBS 插补", 1),
        ("五轴 RTCP、非线性误差控制、刀具矢量平滑", 1),
        ("误差补偿：几何/热/空间误差，摩擦与反向间隙", 1),
        ("自适应控制与颤振抑制、伺服参数自整定", 1),
        "**格局**",
        ("FANUC、Siemens、Heidenhain、Mitsubishi 占高端；华中数控、广州数控、科德、沈阳高精等国产系统在五轴领域取得突破", 1),
        "**差距本质**",
        ("工艺知识积累与软件生态（后处理、仿真、工艺包）", 1),
        ("长期现场可靠性数据与用户信任", 1),
        "**前沿**：AI 数控（工艺参数自优化、数字孪生机床）、云—边协同、开放式数控与软件定义机床",
    ], size=13, para_space=4)

    # 2.4b 伺服
    s = content_slide("2.4 伺服电机与驱动器：高分辨率感知 + 高带宽控制", "决定运动系统的动态响应与跟踪精度",
                      notes=("1.5 分钟", "简要讲伺服指标和趋势：高分辨率编码器、高带宽、SiC 功率器件、多轴一体化。"))
    add_kpis(s, [("23–27 bit", "高端绝对式编码器单圈分辨率（800 万–1.3 亿脉冲/转）"),
                 ("> 2 kHz", "高端伺服驱动器速度环带宽量级"),
                 ("62.5 μs", "高端电流环更新周期量级（16 kHz 及以上）"),
                 ("SiC / GaN", "宽禁带功率器件：更高开关频率、更低损耗、更小体积")],
             height=Inches(1.4))
    add_cards(s, [
        ("电机本体", "高转矩密度永磁同步电机（钕铁硼、集中绕组）\n低齿槽转矩设计、低转动惯量\n高防护、高温等级、耐冷却液"),
        ("编码器", "光学/磁性/电感式绝对编码器\n多圈电池/齿轮/韦根方案\n高速下的信号细分精度与抗干扰"),
        ("驱动器与控制", "矢量控制、前馈、扰动观测器\n振动抑制（陷波、模型追随）\n自整定与机械谐振辨识"),
        ("系统集成", "多轴一体化驱动、共直流母线\n功能安全（STO/SS1/SLS）\n总线同步（分布式时钟 < 1 μs 抖动）"),
    ], cols=4, top=CONTENT_TOP + Inches(1.65), height=Inches(3.0), body_size=12)
    add_callout(s, "国内现状：中低端伺服已大规模替代；高端差距在编码器芯片/光栅、驱动算法成熟度、长期可靠性",
                top=Inches(6.05), height=Inches(0.6))

    # 2.5 液压气动
    s = content_slide("2.5 高端液压与气动元件", "功率密度最高的驱动方式，重型装备与航空航天不可替代",
                      notes=("2 分钟", "讲电液伺服阀、高压柱塞泵、精密气动元件的指标与难点。强调摩擦副材料、配磨工艺与污染控制是难点所在。"))
    add_cards(s, [
        ("电液伺服阀 / 比例阀", "喷嘴挡板、射流管两级伺服阀\n频宽 > 100–200 Hz，滞环 < 0.5%\n难点：阀芯阀套配磨（间隙 2–5 μm，棱边锐度）、力矩马达、抗污染"),
        ("高压柱塞泵 / 马达", "压力 35–42 MPa（工程机械）、更高（航空）\n难点：配流盘/缸体、滑靴/斜盘三大摩擦副材料与表面处理，噪声、效率、寿命"),
        ("精密气动元件", "高速电磁阀、精密减压阀、气浮导轨供气\n难点：响应时间（ms）、流量一致性、洁净度\n半导体与电子装配高度依赖"),
        ("电液作动器（EHA）", "泵控、集成化、机电液一体\n应用：飞行器舵机、重型机器人\n难点：高速电机泵、热管理、容错控制"),
    ], cols=4, body_size=12, height=Inches(3.1))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(3.3), CONTENT_W, Inches(2.4), [
        "**国内外格局**：Bosch Rexroth、Moog、Parker、Eaton、Kawasaki 领先；恒立液压、艾迪精密在工程机械液压件实现规模替代，航空航天伺服阀以航空工业 609 所等体系为主",
        "**研究方向**：摩擦副润滑—磨损—热耦合机理与新材料（陶瓷、DLC、PEEK 复合）；数字液压（高速开关阀阵列）；液压系统的智能诊断与节能控制；电静液作动器的高功率密度设计",
    ], size=13.5)
    add_callout(s, "共性挑战：液压元件寿命与可靠性对油液污染极度敏感 —— 研发不能脱离系统级的污染控制与应用工况",
                top=Inches(6.1), height=Inches(0.6))

    # 2.6 传感测量
    s = content_slide("2.6 精密测量与感知部件：装备的“眼睛”", "没有纳米级测量就没有纳米级制造",
                      notes=("2 分钟", "讲光栅、激光干涉、编码器、测头、力传感器。强调光栅制造（刻划、全息）与信号细分是难点，高端几乎被 Heidenhain、Renishaw 垄断。"))
    add_table(s, [
        ["部件", "原理", "典型性能", "应用", "难点"],
        ["直线光栅尺", "光学/磁栅莫尔条纹、干涉", "分辨率 1 nm–0.1 μm；精度 ±0.5–3 μm/m", "机床闭环、测量机", "基底刻划/全息制作、信号细分与误差"],
        ["圆光栅/角编码器", "增量/绝对码盘", "分辨率 < 0.01 arcsec（高端）；精度 ±1–5 arcsec", "转台、机器人关节", "码盘加工、偏心与读数头多头平均"],
        ["激光干涉仪", "He-Ne 激光干涉", "分辨率 0.1–1 nm，线性误差 ppm 级", "光刻机测量、机床检定", "稳频激光、环境补偿、光学元件"],
        ["接触/非接触测头", "触发/扫描/光学", "重复性 < 0.25–1 μm", "在机测量、CMM", "传感机构、温漂、与数控集成"],
        ["六维力/力矩传感器", "应变、电容、压电", "量程/精度/串扰 < 1–3%", "机器人装配、打磨、人机协作", "解耦结构设计、标定、温度补偿"],
    ], col_widths=[1.5, 1.9, 2.9, 1.8, 3.0], font_size=11, row_h=Inches(0.62))
    add_callout(s, "前沿：平面光栅（光刻机工件台 2D 位置测量）、绝对式纳米光栅、光纤与量子传感、测量—控制一体化（“测控融合”）",
                top=Inches(5.75), height=Inches(0.55), size=12.5)
    add_text(s, MARGIN, Inches(6.35), CONTENT_W, Inches(0.6),
             "国内：长春光机所体系、禹衡光学、中科院微电子所等在光栅/编码器方面持续突破；高端绝对式纳米光栅仍主要依赖进口。",
             size=12.5, color=GRAY)

    # 2.7 刀具
    s = content_slide("2.7 高性能刀具与超硬材料", "“工业的牙齿”：直接与工件对话的零部件",
                      notes=("1.5 分钟", "刀具是消耗品但技术含量极高。讲材料（硬质合金、PCD/CBN、陶瓷）、涂层、几何与国内外格局。"))
    add_cards(s, [
        ("刀具材料", "超细/纳米晶硬质合金（WC-Co，晶粒 < 0.5 μm）\nPCD、PCBN 超硬材料；陶瓷（Al₂O₃、Si₃N₄、SiAlON）\n粉末冶金高速钢\n难点：WC 粉末与钴粉纯度、烧结致密度"),
        ("涂层技术", "PVD（TiAlN、AlCrN、纳米多层/纳米复合）\nCVD（Al₂O₃ 厚膜）、金刚石涂层\n难点：膜基结合力、残余应力、涂层设备"),
        ("刀具几何与结构", "整体立铣刀槽型、变螺旋/不等分齿\n可转位刀片断屑槽、刀体结构\n难点：五轴磨削工艺、刀具—工艺匹配数据库"),
        ("格局与趋势", "Sandvik、Kennametal、Iscar、Mitsubishi、Kyocera 领先\n株洲钻石、厦门金鹭、欧科亿、华锐精密快速成长\n趋势：智能刀具（内置传感）、增材刀体、绿色干切"),
    ], cols=4, body_size=11.8, height=Inches(3.6))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(3.8), CONTENT_W, Inches(2.2), [
        "**科学问题**：难加工材料（钛合金、高温合金、CFRP）切削的刀—屑界面摩擦与磨损机理；涂层—基体—切削热耦合的失效机制",
        "**工程问题**：刀具寿命的离散度（批次一致性）；刀具—主轴—工艺参数的匹配数据库；在线刀具磨损监测与寿命预测",
    ], size=13.5)

    # 2.8 密封润滑冷却
    s = content_slide("2.8 密封、润滑与热管理单元", "常被忽视的“隐形”关键部件，却是现场可靠性的首要影响因素",
                      notes=("1.5 分钟", "简短但强调重要性。主轴密封失效、润滑不良、冷却不均是现场故障的高频原因。"))
    add_cards(s, [
        ("密封技术", "非接触：迷宫密封、气幕密封、磁流体密封（真空）\n接触：机械密封、唇形密封\n难点：高速下的发热与磨损、微间隙加工、半导体真空/洁净要求"),
        ("润滑技术", "油气润滑（高速主轴）、微量润滑（MQL）、脂润滑寿命\n固体润滑（真空/空间：MoS₂、DLC）\n难点：润滑剂寿命模型、极端环境润滑"),
        ("热管理", "主轴/电机水冷套、丝杠中空冷却\n热对称结构、低膨胀材料（殷钢、碳纤维）\n难点：多热源耦合的温度场预测与主动控制"),
    ], cols=3, body_size=12.5, height=Inches(2.8))
    add_callout(s, "“可靠性是设计出来的，也是润滑与密封保证出来的” —— 建议将摩擦学与热管理纳入零部件研发的早期阶段",
                top=CONTENT_TOP + Inches(3.0), height=Inches(0.6))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(3.8), CONTENT_W, Inches(1.9), [
        "**现场统计**：主轴密封失效导致的轴承进水/进屑、润滑不良引起的轴承早期失效、冷却不均导致的热变形，是机床现场故障中最常见的几类“非核心件”故障",
        "**研究方向**：极端工况（高速/高温/真空/洁净）润滑机理；润滑状态在线感知；基于相变/微通道的高功率密度冷却；热误差建模从“经验回归”走向“物理+数据”融合",
    ], size=13.5)

    # 2.9 机器人与半导体特殊部件
    s = content_slide("2.9 机器人与半导体装备的特殊关键部件", "新兴装备门类带来新的零部件需求",
                      notes=("2 分钟", "介绍半导体装备特有的零部件（工件台、真空机械手、静电卡盘等）和机器人新需求（关节模组、灵巧手、触觉）。这是未来十年需求增长最快的方向。"))
    add_cards(s, [
        ("超精密工件台（光刻/检测）", "气浮/磁浮 + 直线电机 + 平面光栅/干涉仪\n定位精度 nm 级，加速度数 g，同步误差 nm\n粗精双级、六自由度、主动隔振"),
        ("真空机械手与传输", "真空兼容材料与润滑、磁流体密封\n晶圆对中、低颗粒产生、高重复定位（±0.05 mm）\n直驱真空电机、非接触传感"),
        ("静电卡盘 / 加热盘", "陶瓷（AlN、Al₂O₃）基体、嵌入电极\n温度均匀性 ±1 ℃、平面度 μm 级\n材料、共烧、精密加工"),
        ("关节模组与执行器", "无框力矩电机 + 谐波/行星 + 双编码器 + 驱动器一体\n高扭矩密度（人形机器人需求）\n柔性关节、准直驱（QDD）"),
        ("力/触觉与灵巧手", "六维力传感器、关节力矩传感器\n电子皮肤、分布式触觉阵列\n微型减速器、腱驱动与微型伺服"),
        ("精密运动部件（泛半导体）", "振镜（激光加工）、精密直线模组\n晶圆切割/减薄主轴（> 60 000 r/min 气浮）\n高速贴片头（压电/音圈）"),
    ], cols=3, body_size=12, height=Inches(4.6))
    add_callout(s, "共同特征：极端精度 + 洁净/真空 + 高动态 + 多物理场集成 —— 零部件与整机的边界日益模糊，呼唤“部件级系统工程”",
                top=Inches(6.15), height=Inches(0.55), size=12.5)

    # 2.10 小结表
    s = content_slide("第二部分小结：关键零部件研发态势一览", "横向比较：瓶颈类型、国内水平与学术机会",
                      notes=("2 分钟", "总结第二部分。用表格引导听众看到：不同零部件的瓶颈类型不同（材料主导/工艺主导/软件生态主导），对应的学术介入方式也不同。"))
    add_table(s, [
        ["零部件", "瓶颈类型", "国内水平（高端段）", "学术研究切入点"],
        ["精密轴承", "材料 + 工艺 + 试验数据", "中档成熟，高端差距明显", "接触疲劳机理、多场耦合动力学、智能轴承"],
        ["丝杠导轨", "工艺 + 精度保持性", "中档有竞争力，高端追赶", "滚动体动力学、热误差、新原理丝杠"],
        ["精密减速器", "设计 + 工艺 + 装配", "谐波已替代，RV 追赶", "啮合理论、疲劳寿命、概率公差设计"],
        ["电主轴", "机电热耦合 + 轴承", "中低速成熟，高速大功率差距", "热—动力学耦合、智能主轴、磁悬浮"],
        ["数控/伺服", "软件生态 + 可靠性", "五轴突破，生态薄弱", "智能数控算法、开放式架构、AI 工艺优化"],
        ["液压元件", "材料 + 摩擦副工艺", "工程机械替代，航空差距", "摩擦学、数字液压、EHA"],
        ["测量感知", "光学制造 + 芯片", "中端可用，纳米级依赖进口", "新原理测量、测控融合、量子传感"],
        ["刀具", "材料 + 涂层 + 数据", "快速成长，高端追赶", "切削机理、涂层失效、智能刀具"],
    ], col_widths=[1.5, 2.6, 3.0, 4.6], font_size=11.5, row_h=Inches(0.52))
    add_callout(s, "共性规律：瓶颈从“原理”下沉到“材料—工艺—检测—数据”；学术界的价值在于机理阐释、模型工具与新原理探索",
                top=Inches(6.15), height=Inches(0.55), size=13)


# =====================================================================
# 第三部分 共性使能技术
# =====================================================================
def part3():
    section_slide(2, "本部分抽取第二部分反复出现的共性技术：超精密加工、精密装配、材料与热处理、表面工程、多场耦合设计、可靠性、测试评价、数字化研发。")

    s = content_slide("3.1 精密与超精密加工技术", "零部件精度的直接来源：加工精度从微米走向纳米",
                      notes=("2.5 分钟", "讲 Taniguchi 曲线式的精度演进，各类超精密加工方法的适用范围。强调“母机精度决定零件精度”的递归关系。"))
    # 精度阶梯
    levels = [("普通加工", "10–100 μm", STEEL), ("精密加工", "1–10 μm", GREEN), ("高精密加工", "0.1–1 μm", TEAL),
              ("超精密加工", "10–100 nm", NAVY_2), ("纳米/原子级", "< 10 nm", AMBER)]
    for i, (n, v, c) in enumerate(levels):
        x = MARGIN + i * Inches(1.25)
        yy = CONTENT_TOP + Inches(2.6) - i * Inches(0.5)
        hh = Inches(0.6) + i * Inches(0.5)
        b = add_rect(s, x, yy, Inches(1.15), hh, c)
        shape_text(b, f"{n}\n{v}", size=10.5, bold=True, color=WHITE if c != AMBER else NAVY, anchor=MSO_ANCHOR.TOP)
    add_text(s, MARGIN, CONTENT_TOP + Inches(3.3), Inches(6.3), Inches(0.4),
             "加工精度的演进（示意）：每一级跃迁都依赖新的机床、刀具、测量与环境控制", size=11.5, color=GRAY)
    add_table(s, [
        ["方法", "典型精度", "适用零部件"],
        ["单点金刚石车削（SPDT）", "形状 PV < 0.1 μm，Ra < 5 nm", "光学模具、气浮轴承、反射镜"],
        ["超精密磨削 / ELID 磨削", "Ra < 10 nm，形状 < 0.1 μm", "轴承沟道、硬脆材料、陶瓷"],
        ["精密成形磨削", "轮廓 ≤ 2 μm", "摆线轮、丝杠滚道、齿轮"],
        ["研磨 / 抛光（CMP、MRF、IBF）", "面形 < 1 nm RMS（光学）", "光刻物镜、EUV 反射镜、晶圆"],
        ["精密珩磨 / 超精研", "圆度 < 0.2 μm", "轴承套圈、液压阀套、缸体"],
        ["微细/特种加工（EDM、激光、FIB）", "μm–nm 特征", "喷嘴、节流器、微结构"],
    ], left=Inches(7.0), width=Inches(5.8), top=CONTENT_TOP, col_widths=[2.3, 2.0, 2.0], font_size=10.5, row_h=Inches(0.52))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(3.8), CONTENT_W, Inches(2.2), [
        "**“母机递归”问题**：超精密零部件需要更高精度的机床加工，而机床本身由零部件组成 —— 突破需依靠误差均化、误差分离与补偿等“以低制高”方法",
        "**科学问题**：纳米尺度材料去除机理（脆—塑转变）、亚表面损伤、加工表面完整性对零部件服役性能的影响",
    ], size=13.5)

    s = content_slide("3.2 精密装配与调试：性能实现的“最后一公里”", "同样的零件，不同的装配，性能可相差数倍",
                      notes=("2 分钟", "强调装配不是简单组合。讲选配、预紧控制、装配误差建模、装配过程监测。举例：成组轴承预紧、减速器回差装配、主轴整体动平衡。"))
    add_cards(s, [
        ("装配精度理论", "尺寸链/公差分析（最坏情况、统计、蒙特卡洛）\n小位移旋量（SDT）、雅可比—旋量模型\n接触变形与装配应力对精度的影响"),
        ("选配与成组技术", "轴承成组、滚珠分级选配、摆线轮—针齿配对\n基于测量数据的智能选配算法（组合优化）\n从“全互换”走向“选择装配 + 补偿”"),
        ("预紧与调整工艺", "轴承预紧力/预紧量的精确施加与在线监测\n丝杠预拉伸、导轨预压调整\n装配过程中的力—位移—温度多参数监控"),
        ("动平衡与整机调试", "高速转子装配后整体动平衡（G0.4 及以上）\n整机几何精度检测与误差补偿参数标定\n伺服参数整定与动态性能验证"),
    ], cols=2, body_size=12.5, height=Inches(4.2))
    add_callout(s, "研究方向：面向性能的装配（Performance-oriented assembly）、数字孪生驱动的装配工艺优化、装配知识的显性化与自动化",
                top=Inches(5.75), height=Inches(0.6), size=13)

    s = content_slide("3.3 材料与热处理：零部件性能的物质基础", "“材料先行”—— 多数高端零部件的差距可追溯至材料与热处理",
                      notes=("2.5 分钟", "讲几类关键材料体系与热处理工艺。重点：轴承钢纯净度，渗碳钢的组织控制，陶瓷材料，以及真空热处理与深冷处理。"))
    add_table(s, [
        ["材料体系", "代表牌号/类型", "关键控制指标", "典型应用"],
        ["高碳铬轴承钢", "GCr15、GCr15SiMn、100Cr6（ISO）", "氧含量 ≤ 5–8 ppm，Ti ≤ 15 ppm，夹杂物评级，碳化物均匀性", "轴承套圈、滚动体、丝杠"],
        ["渗碳轴承/齿轮钢", "G20CrNi2MoA、18CrNiMo7-6、M50NiL", "渗层深度与硬度梯度、心部韧性、淬透性带宽", "重载轴承、RV 摆线轮、齿轮"],
        ["高温轴承钢", "Cr4Mo4V（M50）、Cr15Mo4", "高温硬度（> 300 ℃）、尺寸稳定性", "航空发动机主轴承"],
        ["工程陶瓷", "Si₃N₄、ZrO₂、AlN、SiC", "致密度、断裂韧性、缺陷尺寸（< 10 μm）", "陶瓷球、静电卡盘、超精密导轨"],
        ["特殊合金/功能材料", "殷钢、钛合金、高强韧柔轮钢、钕铁硼", "热膨胀系数、疲劳极限、磁性能稳定性", "热稳定结构、柔轮、永磁电机"],
    ], col_widths=[1.7, 3.0, 4.2, 2.6], font_size=11, row_h=Inches(0.58))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(3.7), CONTENT_W, Inches(2.3), [
        "**热处理关键工艺**：真空/可控气氛淬火、等温淬火（贝氏体）、深冷处理（残余奥氏体 < 5–8%）、感应淬火；变形控制与尺寸稳定性（时效）",
        "**科学问题**：夹杂物—疲劳寿命的定量关系（极值统计）；组织—性能—工艺的多尺度模拟（ICME）；残余应力演化与尺寸稳定性机理",
        "**工程问题**：批次稳定性、热处理工艺窗口数据库、特种冶金装备（真空感应 + 电渣重熔双联工艺）",
    ], size=13)

    s = content_slide("3.4 表面工程与摩擦学", "零部件失效 80% 以上始于表面：磨损、疲劳、腐蚀",
                      notes=("1.5 分钟", "表面工程是提升零部件寿命的有效途径。讲涂层、表面强化和摩擦学设计。"))
    add_cards(s, [
        ("涂层与薄膜", "DLC、TiN/TiAlN、CrN、MoS₂\n热喷涂、PVD/CVD、ALD\n应用：液压摩擦副、刀具、真空轴承\n难点：膜基结合、厚度均匀性、内应力"),
        ("表面强化", "喷丸、激光冲击强化（LSP）、超声滚压、渗氮\n引入残余压应力、细化晶粒\n应用：齿轮、柔轮、丝杠滚道\n难点：工艺—残余应力—疲劳寿命定量关系"),
        ("表面织构与功能表面", "激光/光刻微织构（减摩、储油）\n超疏水/亲油表面\n应用：滑动轴承、密封、导轨\n难点：织构参数优化、耐久性"),
        ("摩擦学设计", "弹流润滑（EHL）、混合润滑、边界润滑\n润滑剂—材料—表面协同\n应用：几乎所有运动副\n难点：极端工况润滑、多尺度摩擦建模"),
    ], cols=4, body_size=12, height=Inches(3.6))
    add_callout(s, "表面完整性（Surface Integrity）：粗糙度 + 残余应力 + 显微组织 + 微裂纹 —— 应作为零部件制造的核心质量指标，而非仅看尺寸公差",
                top=CONTENT_TOP + Inches(3.8), height=Inches(0.65), size=13)
    add_text(s, MARGIN, CONTENT_TOP + Inches(4.6), CONTENT_W, Inches(0.9),
             "研究方向：原位摩擦学观测（同步辐射、原位 SEM）；表面完整性—服役性能的跨尺度模型；自润滑/自修复表面；面向极端环境（真空、低温、辐照）的摩擦学。",
             size=12.5, color=GRAY)

    s = content_slide("3.5 多物理场耦合设计与误差补偿", "从“经验设计 + 试错”走向“建模驱动 + 优化”",
                      notes=("2.5 分钟", "讲热—结构—动力学耦合建模、拓扑优化、误差建模与补偿方法。热误差是机床最大误差源，介绍补偿思路。"))
    add_bullets(s, MARGIN, CONTENT_TOP, Inches(6.3), Inches(5.5), [
        "**多场耦合建模**",
        ("热—结构：热源辨识（电机、轴承、切削）、对流边界、热变形", 1),
        ("结构—动力学：结合面刚度/阻尼辨识、模态、颤振稳定性叶瓣图", 1),
        ("电磁—热—结构（电主轴、直线电机）", 1),
        ("流—固—热（气浮/液压元件）", 1),
        "**结构优化**",
        ("拓扑/尺寸/形状优化：轻量化 + 高刚度 + 热对称", 1),
        ("点阵/仿生结构（结合增材制造）", 1),
        "**误差建模与补偿**",
        ("几何误差：多体系统理论/齐次坐标变换，21 项误差（三轴）", 1),
        ("热误差：温度传感 + 回归/神经网络/物理模型，补偿可降低热误差 50–80%", 1),
        ("动态误差：前馈、迭代学习、交叉耦合控制", 1),
    ], size=14, para_space=4)
    add_rect(s, Inches(7.1), CONTENT_TOP, Inches(5.7), Inches(5.5), GRAY_LIGHT)
    add_text(s, Inches(7.25), CONTENT_TOP + Inches(0.1), Inches(5.4), Inches(0.4), "热误差补偿技术路线（示意）", size=14, bold=True, color=NAVY_2)
    steps = ["温度/位移传感布点优化", "热特性试验与数据采集", "建模（回归 / 热模态 / 机器学习）", "数控系统实时补偿", "鲁棒性验证（工况、季节）"]
    for i, st in enumerate(steps):
        yy = CONTENT_TOP + Inches(0.6) + i * Inches(0.82)
        b = add_rect(s, Inches(7.3), yy, Inches(5.3), Inches(0.62), TEAL if i % 2 == 0 else NAVY_2)
        shape_text(b, f"{i+1}. {st}", size=12.5, bold=True, align=PP_ALIGN.LEFT, margin=0.15)
    add_text(s, Inches(7.25), CONTENT_TOP + Inches(4.75), Inches(5.4), Inches(0.7),
             "科学问题：热误差模型的泛化性与鲁棒性（工况变化下失效）；结合面特性辨识；“物理 + 数据”混合建模。",
             size=11.5, color=GRAY)

    s = content_slide("3.6 可靠性工程：从“能用”到“耐用”", "高端装备的本质竞争力：MTBF 与精度保持性",
                      notes=("2.5 分钟", "讲可靠性设计、试验与评估方法，以 04 专项推动国产数控机床 MTBF 提升为例。强调失效物理与加速试验的重要性。"))
    add_kpis(s, [("MTBF", "平均无故障时间：国产数控机床由 ~1000 h 提升至 > 2000 h（04 专项），国际先进 > 3000 h"),
                 ("L10", "轴承 90% 可靠度额定寿命；高端主轴轴承要求数万小时"),
                 ("Weibull", "寿命分布主流模型；形状参数 β 区分早期/随机/耗损失效"),
                 ("PoF", "失效物理：基于机理的寿命预测，替代纯统计方法")],
             height=Inches(1.45), big_size=24, label_size=10.5)
    add_cards(s, [
        ("可靠性设计", "可靠性分配与预计、FMEA/FTA\n降额设计、冗余、容错\n热设计、EMC 设计、防护设计"),
        ("可靠性试验", "加速寿命试验（ALT）、加速退化试验（ADT）\n强化试验（HALT）、环境应力筛选（ESS）\n专用台架：轴承、减速器、丝杠、电主轴寿命台"),
        ("可靠性评估与增长", "现场数据采集与故障统计、Weibull/贝叶斯评估\n可靠性增长（Duane 模型）、FRACAS 闭环\n精度保持性评价方法与标准"),
    ], cols=3, top=CONTENT_TOP + Inches(1.65), height=Inches(2.9), body_size=12.5)
    add_callout(s, "学术机会：多失效模式竞争下的寿命预测；小样本/无失效数据的可靠性评估；精度退化（而非功能失效）的建模与加速试验方法",
                top=Inches(6.0), height=Inches(0.65), size=12.5)

    s = content_slide("3.7 精密测试与性能评价", "标准化、可比较的测试是零部件研发的“度量衡”",
                      notes=("1.5 分钟", "讲零部件和整机测试手段与标准体系。强调测试能力建设和标准话语权。"))
    add_table(s, [
        ["对象", "测试内容", "仪器/方法", "相关标准"],
        ["轴承", "旋转精度、摩擦力矩、振动噪声、温升、寿命", "圆度仪、波纹度仪、振动测试机、寿命台", "ISO 492、ISO 15242、GB/T 307"],
        ["丝杠/导轨", "行程偏差、预紧力矩、刚度、噪声", "激光干涉仪、扭矩传感器、刚度试验台", "ISO 3408、JIS B1192、GB/T 17587"],
        ["减速器", "传动误差、回差、刚度、效率、寿命", "双读数头角度测量、扭矩加载台", "GB/T 37718、JB/T 系列"],
        ["电主轴", "回转精度、热漂移、动刚度、动平衡", "电容位移传感器、锤击/激振模态、动平衡机", "ISO 230-3、ISO 21940、GB/T 17421"],
        ["整机", "几何/定位/热/动态精度、圆度轨迹", "激光干涉仪、球杆仪、激光跟踪仪、R-test", "ISO 230 系列、ISO 10791、ISO 13041"],
    ], col_widths=[1.3, 3.2, 3.6, 3.4], font_size=11.5, row_h=Inches(0.6))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(3.85), CONTENT_W, Inches(2.0), [
        "**趋势**：从静态到动态、从离线到在线/在机、从单项指标到工况下的综合性能；测试数据标准化与共享",
        "**学术机会**：新型测试原理（多自由度误差同步测量）、精度退化的快速评估方法、参与国际标准制定（ISO/TC 39、TC 4、TC 213）",
    ], size=13.5)

    s = content_slide("3.8 数字化研发：仿真驱动、模型为中心", "MBSE、CAE、数字孪生与知识工程",
                      notes=("1.5 分钟", "简述数字化研发工具链。强调仿真可信度（V&V）和知识库是关键。"))
    add_chevrons(s, ["需求建模 MBSE", "多学科仿真 CAE", "代理模型与优化", "虚拟试验 V&V", "数字孪生反馈"],
                 height=Inches(0.85), size=12.5)
    add_cards(s, [
        ("工具链", "CAD/CAE/CAM 一体化；多物理场求解器（ANSYS、COMSOL、Abaqus）与国产 CAE 自主化\n专用工具：轴承（BEAST、ROMAX）、齿轮（KISSsoft、MASTA）\n仿真可信度评估（V&V）"),
        ("知识工程", "设计规则库、工艺参数库、失效案例库\n设计知识的显性化与复用\n大模型辅助的知识检索与设计助手（探索中）"),
        ("数字孪生", "零部件级：主轴/丝杠热态与寿命孪生\n装备级：几何—热—动态误差的实时镜像\n难点：模型降阶、实时性、模型更新（数据同化）"),
    ], cols=3, top=CONTENT_TOP + Inches(1.1), height=Inches(3.3), body_size=12.5)
    add_callout(s, "核心观点：数字化研发的瓶颈不是软件而是“模型 + 数据 + 知识”；国产 CAE 与专用零部件设计软件的自主化本身就是重要研究方向",
                top=Inches(6.0), height=Inches(0.65), size=12.5)


# =====================================================================
# 第四部分 前沿技术
# =====================================================================
def part4():
    section_slide(3, "本部分面向未来 5–15 年，介绍八大前沿方向。每个方向：技术内涵 → 典型进展 → 科学问题。")

    s = content_slide("前沿技术全景：八大方向", "“更智能、更极端、更集成、更绿色”",
                      notes=("1.5 分钟", "总览八大前沿方向，说明它们之间的关联：智能化与AI是横向贯穿，材料与增材是纵向支撑，极端制造是牵引。"))
    dirs = [("① 智能零部件", "自感知 · 自诊断 · 自适应"), ("② 增材制造使能", "拓扑优化 · 点阵 · 一体化"),
            ("③ 非接触与磁悬浮", "无磨损 · 超高速 · 主动控制"), ("④ 先进材料", "高熵合金 · 陶瓷基复合 · 功能材料"),
            ("⑤ AI 驱动研发", "生成式设计 · 代理模型 · PINN"), ("⑥ PHM 与数字孪生", "预测性维护 · 模型—数据融合"),
            ("⑦ 微纳与新型驱动", "压电 · 柔性机构 · MEMS"), ("⑧ 极端制造部件", "纳米定位 · 真空 · EUV 光学")]
    colors = [TEAL, NAVY_2, GREEN, AMBER, TEAL, NAVY_2, GREEN, AMBER]
    for i, (t, d) in enumerate(dirs):
        r, c = divmod(i, 4)
        x = MARGIN + c * (Inches(3.0) + Inches(0.08))
        y = CONTENT_TOP + r * Inches(2.0)
        b = add_rect(s, x, y, Inches(3.0), Inches(1.8), colors[i])
        shape_text(b, f"{t}\n\n{d}", size=14, bold=True, color=WHITE if colors[i] != AMBER else NAVY)
    add_callout(s, "贯穿主线：传感 + 计算 + 材料 的融合，使零部件从“被动的机械要素”转变为“主动的信息物理单元”",
                top=Inches(5.6), height=Inches(0.6))
    add_text(s, MARGIN, Inches(6.3), CONTENT_W, Inches(0.6),
             "另有两个横向主题：高速高精运动控制前沿、绿色低碳与可持续制造，将在本部分末尾讨论。", size=12.5, color=GRAY)

    s = content_slide("4.1 智能零部件：自感知、自诊断、自适应", "把传感器、计算与执行嵌入零部件本体",
                      notes=("2.5 分钟", "讲智能轴承、智能主轴、智能丝杠、智能刀具的概念与进展（如 SKF 传感轴承、Schaeffler 智能轴承、工业界的智能刀柄）。强调嵌入式传感的空间、供能、信号传输难题。"))
    add_cards(s, [
        ("智能轴承", "嵌入式温度/振动/载荷/转速传感（薄膜、光纤、压电）\n自供能（能量收集）与无线传输\n预紧主动调控（压电作动器）\n进展：SKF Insight、Schaeffler 传感轴承、国内多所高校样机"),
        ("智能主轴 / 刀柄", "内置位移、振动、温度、功率传感\n颤振在线识别与主动抑制（主动阻尼器、变转速）\n刀具磨损/破损监测\n进展：智能刀柄（无线应变/振动），商业化初期"),
        ("智能丝杠 / 导轨", "预压、润滑状态、温度场在线感知\n热伸长实时估计与补偿\n剩余寿命预测\n进展：集成温度传感的丝杠、润滑监测模块"),
        ("智能减速器 / 关节", "内置扭矩传感、温度、振动\n回差与刚度在线辨识\n面向协作机器人的碰撞检测与柔顺控制\n进展：一体化关节模组标配双编码器 + 扭矩传感"),
    ], cols=2, body_size=12, height=Inches(4.3))
    add_callout(s, "科学问题：微型传感器在高速/高温/强电磁环境下的可靠嵌入；弱信号提取与特征解耦；传感—结构一体化对零部件本体性能的影响",
                top=Inches(5.85), height=Inches(0.6), size=12.5)

    s = content_slide("4.2 增材制造使能的零部件创新", "设计自由度的释放：从“可制造性约束设计”到“性能驱动设计”",
                      notes=("2.5 分钟", "讲增材制造对零部件设计的改变：拓扑优化、点阵结构、随形冷却流道、一体化成形、多材料。举例：增材轻量化机器人臂、随形冷却主轴壳体、增材刀体。也要讲局限：精度、表面、疲劳性能、成本。"))
    add_bullets(s, MARGIN, CONTENT_TOP, Inches(6.4), Inches(5.5), [
        "**设计范式改变**",
        ("拓扑优化 + 点阵/TPMS 结构：刚度—质量比提升 30–50%", 1),
        ("随形冷却/润滑流道：主轴壳体、液压阀块、电机壳体热管理", 1),
        ("多零件一体化：减少装配接口与误差链（如液压阀块 70% 减重）", 1),
        ("多材料/梯度材料：耐磨表面 + 高韧心部，热膨胀匹配", 1),
        "**典型应用探索**",
        ("轻量化机器人臂、机床运动部件（龙门横梁、滑枕）", 1),
        ("增材刀体与随形冷却刀具；拓扑优化刀柄", 1),
        ("增材制造气浮节流器（多孔结构）与微流道", 1),
        ("航空发动机轴承座、整体叶盘等极端部件", 1),
        "**局限与挑战**",
        ("尺寸精度 ±0.1 mm、表面 Ra 5–15 μm，需后续精加工；疲劳性能离散；内部缺陷；成本", 1),
    ], size=14, para_space=4)
    add_rect(s, Inches(7.2), CONTENT_TOP, Inches(5.6), Inches(5.5), GRAY_LIGHT)
    add_text(s, Inches(7.35), CONTENT_TOP + Inches(0.1), Inches(5.3), Inches(0.4), "“增材 + 减材 + 智能”复合制造路线", size=14, bold=True, color=NAVY_2)
    chain = ["性能需求 → 拓扑/点阵生成式设计", "增材成形（SLM/EBM/DED）", "热处理 / HIP 消除缺陷", "精密减材加工关键配合面", "无损检测（CT）与性能验证"]
    for i, c in enumerate(chain):
        b = add_rect(s, Inches(7.4), CONTENT_TOP + Inches(0.6) + i * Inches(0.8), Inches(5.2), Inches(0.62),
                     TEAL if i % 2 == 0 else NAVY_2)
        shape_text(b, c, size=12.5, bold=True, align=PP_ALIGN.LEFT, margin=0.15)
    add_text(s, Inches(7.35), CONTENT_TOP + Inches(4.7), Inches(5.3), Inches(0.8),
             "科学问题：增材零部件的工艺—组织—缺陷—疲劳关系；面向增材的零部件设计理论（DfAM）；点阵结构的阻尼与动力学特性。",
             size=11.5, color=GRAY)

    s = content_slide("4.3 非接触支承与磁悬浮技术前沿", "“零磨损”带来的超高速、超洁净与主动可控",
                      notes=("2 分钟", "讲磁悬浮电主轴、磁悬浮转台、磁悬浮平面电机、超导磁悬浮以及气磁混合支承。重点是半导体与超高速领域的牵引。"))
    add_cards(s, [
        ("磁悬浮电主轴", "转速 > 60 000–100 000 r/min，无需润滑\n主动控制刚度/阻尼，可实现回转误差主动补偿、颤振主动抑制\n挑战：功放、控制带宽、断电保护、成本"),
        ("磁悬浮平面电机 / 工件台", "六自由度无接触运动，适合真空与超洁净\n半导体光刻、晶圆检测的下一代平台\n挑战：多自由度解耦控制、纳米级测量、热"),
        ("超导磁悬浮", "高温超导体的钉扎效应实现被动稳定悬浮\n零功耗、本征稳定\n应用探索：超高速飞轮、超精密隔振、磁悬浮轴承\n挑战：低温系统、承载与刚度"),
        ("气磁混合与新型气浮", "气浮刚度 + 磁力预载/主动调节\n多孔质石墨节流、微孔阵列（增材/激光制造）\n主动气浮（压电调节节流）实现纳米级误差补偿"),
    ], cols=4, body_size=12, height=Inches(3.8))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(4.0), CONTENT_W, Inches(1.9), [
        "**科学问题**：高速磁悬浮转子的非线性动力学与鲁棒控制；自传感（self-sensing）磁轴承的信号解耦；气膜微尺度流动的稳定性与极限刚度",
        "**工程趋势**：非接触支承 + 直驱 + 纳米测量 的“三位一体”成为极端精密装备的标准架构",
    ], size=13.5)

    s = content_slide("4.4 先进材料：从“选材”到“设计材料”", "新材料体系为零部件性能跃升提供可能",
                      notes=("2 分钟", "介绍几类有潜力的新材料及其在零部件中的应用前景。也提醒材料从实验室到工程应用的周期长（10–20 年）。"))
    add_table(s, [
        ["材料类别", "特点", "零部件应用前景", "成熟度"],
        ["高熵合金（HEA）", "多主元固溶、高强韧、耐磨耐蚀、耐高/低温", "极端环境轴承、耐磨涂层、紧固件", "实验室—小批量"],
        ["陶瓷基复合材料（CMC）", "耐高温（> 1200 ℃）、低密度、抗热震", "航空发动机热端部件、高温轴承座、刀具", "航空领域工程化"],
        ["金属玻璃（非晶合金）", "高强度、高弹性极限、低损耗", "精密弹性元件、齿轮（微型）、柔轮探索", "小尺寸工程化"],
        ["碳纤维复合材料", "高比刚度、低热膨胀、高阻尼", "机器人臂、机床运动部件、测量结构件", "工程应用中"],
        ["功能梯度材料（FGM）", "成分/组织连续过渡", "耐磨表面 + 韧性基体、热应力缓解", "增材推动中"],
        ("超硬/超润滑材料", "纳米孪晶金刚石、石墨烯、结构超滑", "超精密刀具、超低摩擦运动副", "基础研究"),
        ["宽禁带半导体（SiC/GaN）", "高频、高温、低损耗", "伺服驱动器、磁悬浮功放、电主轴驱动", "快速产业化"],
    ], col_widths=[2.0, 3.6, 3.6, 1.8], font_size=11.5, row_h=Inches(0.52))
    add_callout(s, "材料基因组（MGI）与 ICME：以计算 + 高通量实验 + 数据加速材料研发，是零部件材料“设计”的方法论基础",
                top=Inches(6.0), height=Inches(0.65), size=13)

    s = content_slide("4.5 AI 驱动的零部件设计与制造", "从“仿真驱动”到“数据 + 物理 + 智能”驱动",
                      notes=("2.5 分钟", "讲 AI 在设计（生成式、代理模型、PINN）、工艺（参数优化、过程监控）和运维中的应用。强调小样本、可解释、物理一致性是学术难点。谨慎提醒：AI 不能替代机理与试验。"))
    add_cards(s, [
        ("生成式设计与代理模型", "拓扑/形状生成（扩散模型、GAN）\n高保真仿真的代理模型（神经网络、高斯过程）加速多目标优化 10–100×\n案例：轴承保持架、齿形、结构件的 AI 辅助设计"),
        ("物理信息神经网络（PINN）", "将 PDE 约束嵌入网络，小样本下仍物理一致\n应用：热场重构与热误差预测、EHL 油膜求解、疲劳寿命\n挑战：训练难度、多尺度问题"),
        ("工艺智能", "磨削/热处理/装配参数优化（贝叶斯优化、强化学习）\n过程监控：声发射、振动、视觉的在线质量判别\n挑战：数据获取成本、跨机台迁移"),
        ("知识与大模型", "设计标准、失效案例、工艺知识的检索增强（RAG）\n多模态（图纸、仿真、报告）理解\n挑战：幻觉、专业可靠性、知识产权"),
    ], cols=4, body_size=12, height=Inches(3.8))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(4.0), CONTENT_W, Inches(1.9), [
        "**学术难点**：小样本与数据不平衡；可解释性与物理一致性；跨工况/跨设备泛化（迁移学习、域适应）；仿真—实测数据融合（多保真建模）",
        "**审慎观点**：AI 是加速器而非替代品 —— 机理模型与高质量试验数据仍是零部件研发的根基",
    ], size=13.5)

    s = content_slide("4.6 PHM 与数字孪生：从“定期维护”到“预测性维护”", "零部件健康状态的感知、评估与预测",
                      notes=("2 分钟", "讲 PHM 技术路线（数据驱动、模型驱动、混合）、数字孪生架构、边缘计算。零部件（轴承、丝杠、减速器）是 PHM 的主要对象。"))
    add_chevrons(s, ["状态感知（传感/边缘）", "特征提取与健康指标", "故障诊断（分类/定位）", "退化建模与 RUL 预测", "维护决策与闭环"], height=Inches(0.85))
    add_cards(s, [
        ("数据驱动方法", "深度学习（CNN、LSTM、Transformer）用于振动/电流信号诊断\n无监督/自监督学习应对缺少故障样本\n挑战：标注、泛化、可解释"),
        ("模型驱动与混合方法", "失效物理模型（Paris 裂纹扩展、Archard 磨损）+ 粒子滤波/贝叶斯更新\n“物理 + 数据”混合（灰箱）模型\n优势：外推能力、小样本"),
        ("数字孪生架构", "物理实体 ↔ 虚拟模型 ↔ 数据 ↔ 服务\n降阶模型（ROM）实现实时性\n模型更新（数据同化）保证长期一致性"),
    ], cols=3, top=CONTENT_TOP + Inches(1.1), height=Inches(3.3), body_size=12.5)
    add_callout(s, "科学问题：精度退化（而非功能失效）的健康指标构建；多部件耦合退化建模；少样本/零样本故障诊断；孪生模型的可信度评估",
                top=Inches(6.0), height=Inches(0.65), size=12.5)

    s = content_slide("4.7 微纳制造与新型驱动：超越电磁与滚动的新原理", "压电、柔性机构、MEMS 与智能材料驱动",
                      notes=("2 分钟", "讲纳米定位台（压电 + 柔性铰链）、音圈电机、磁致伸缩、形状记忆合金、介电弹性体等新型驱动，以及 MEMS 传感器与微执行器。"))
    add_table(s, [
        ["驱动原理", "特点", "典型性能", "零部件应用"],
        ["压电（PZT）", "高分辨率、高带宽、大力、小行程", "分辨率 sub-nm；行程 μm–百 μm（放大后 mm）", "纳米定位台、快刀伺服（FTS）、主动阻尼、微量进给"],
        ["柔性机构（柔性铰链）", "无摩擦、无间隙、免润滑", "重复性 nm 级，寿命长", "精密定位、微夹持、精密对准"],
        ["音圈电机（VCM）", "直驱、高动态、行程 mm 级", "加速度 > 10 g、带宽 kHz", "高速贴片、光学调焦、主动隔振"],
        ["磁致伸缩（Terfenol-D、Galfenol）", "大应变、大力、低电压", "应变 ~1000 ppm", "精密作动器、振动主动控制"],
        ["形状记忆合金 / 介电弹性体", "大变形、轻质、柔顺", "应变 > 5%（SMA）、> 100%（DE）", "软体机器人、仿生执行器、可变刚度"],
        ["MEMS 传感与执行", "微型化、批量化、低成本", "加速度计、陀螺、压力、微镜", "机器人惯性测量、智能零部件嵌入式传感"],
    ], col_widths=[2.2, 2.8, 3.0, 3.0], font_size=11.5, row_h=Inches(0.55))
    add_callout(s, "趋势：宏—微复合驱动（粗精双级）；柔性机构的拓扑优化设计；超精密装备中压电快刀伺服、主动隔振平台已成标配",
                top=Inches(5.75), height=Inches(0.55), size=12.5)
    add_text(s, MARGIN, Inches(6.35), CONTENT_W, Inches(0.6),
             "科学问题：压电迟滞/蠕变的建模与补偿；柔性机构大变形非线性动力学；智能材料驱动的疲劳与可靠性。", size=12.5, color=GRAY)

    s = content_slide("4.8 极端制造中的零部件：精度与环境的双重极限", "以光刻机为代表的极端装备，牵引零部件技术的天花板",
                      notes=("2.5 分钟", "讲光刻机双工件台、EUV 光学、极端环境（真空、低温、辐射）部件。这是最能体现“零部件决定装备”的例子，也是前沿研究的焦点。"))
    add_cards(s, [
        ("光刻机双工件台", "定位精度亚纳米—纳米级，加速度数 g，同步误差 nm\n气浮/磁浮 + 直线/平面电机 + 平面光栅/干涉仪\n粗精双级、主动隔振、热控制（mK 级）\n涉及：超精密导向、直驱、测量、控制的极限集成"),
        ("EUV / DUV 光学部件", "EUV 反射镜面形误差 < 0.1 nm RMS（原子级）\n低热膨胀玻璃陶瓷（Zerodur/ULE）、多层膜\n超精密抛光（IBF、MRF）与纳米检测\n涉及：超精密加工测量的极限"),
        ("真空 / 洁净环境部件", "真空兼容轴承（固体润滑、陶瓷）、磁流体密封\n低放气材料、低颗粒运动副\n真空直驱电机、非接触传感\n涉及：材料、润滑、密封的极限"),
        ("低温 / 高温 / 辐照部件", "液氢/液氧泵轴承（航天）、超导低温运动机构\n航空发动机 > 300 ℃ 主轴承\n核环境机器人（辐照老化）\n涉及：材料与润滑在极端温度/辐照下的性能"),
    ], cols=2, body_size=12, height=Inches(4.4))
    add_callout(s, "启示：极端制造装备是零部件技术的“珠峰”，攻克过程中形成的材料、工艺、测量与控制能力将反哺整个装备产业",
                top=Inches(5.95), height=Inches(0.6), size=12.5)

    s = content_slide("4.9 高速高精运动控制前沿", "控制算法是“软零部件”：以算法弥补硬件、释放硬件潜能",
                      notes=("1.5 分钟", "讲先进运动控制方法及其在零部件/装备中的应用。强调控制与机械设计协同（机电一体化设计）。"))
    add_cards(s, [
        ("迭代学习与重复控制", "利用重复轨迹的误差信息逐次修正\n光刻工件台、重复加工轨迹\n跟踪误差降低一个数量级以上"),
        ("模型预测控制（MPC）", "显式处理约束（加速度、电流、热）\n多轴协调、磁悬浮轴承控制\n挑战：实时计算（μs 级周期）"),
        ("主动振动抑制", "输入整形、陷波、主动阻尼\n颤振抑制（变主轴转速、主动刀柄）\n柔性结构的模态控制"),
        ("数据驱动与学习控制", "摩擦/齿槽力等非线性的学习前馈\n参数自整定、强化学习调参\n挑战：安全性与稳定性保证"),
    ], cols=4, body_size=12.5, height=Inches(3.5))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(3.7), CONTENT_W, Inches(2.2), [
        "**机电一体化协同设计**：结构动力学（固有频率、模态）与控制带宽的匹配是高动态系统设计的核心；控制可“软化”机械缺陷，但不能无中生有",
        "**趋势**：计算能力提升（FPGA、多核）使复杂算法进入 μs 级控制周期；控制—测量—结构的一体化设计（“控制可设计性”）",
    ], size=13.5)

    s = content_slide("4.10 绿色低碳与可持续：零部件的全生命周期视角", "“双碳”目标下，高效、长寿命、可再制造成为新指标",
                      notes=("1.5 分钟", "简述绿色制造对零部件的新要求：效率、寿命、再制造、材料回收。"))
    add_cards(s, [
        ("高效与节能", "低摩擦轴承与密封（降低摩擦损耗 20–30%）\n高效传动（行星 vs 蜗轮）、高效电机（IE4/IE5）\n液压系统节能（变量泵、数字液压）"),
        ("长寿命与再制造", "延寿设计、表面再制造（激光熔覆、热喷涂）\n主轴、丝杠、液压件再制造产业\n再制造性能评价与认证"),
        ("绿色工艺与材料", "干切/微量润滑，减少切削液\n无钴硬质合金、稀土减量永磁\n可回收设计与材料流分析"),
        ("全生命周期评价（LCA）", "从材料—制造—使用—回收的碳足迹核算\n零部件碳标签与绿色供应链\n学术机会：零部件 LCA 数据库与方法"),
    ], cols=4, body_size=12.5, height=Inches(3.0))
    add_callout(s, "观点：绿色不只是约束，也是差异化竞争力 —— 低摩擦、长寿命、可再制造本身就是高端零部件的核心指标",
                top=CONTENT_TOP + Inches(3.2), height=Inches(0.6))
    add_bullets(s, MARGIN, CONTENT_TOP + Inches(4.0), CONTENT_W, Inches(1.7), [
        "**量级参考**：工业电机消耗全球电力约 40–45%；机床主轴与进给系统摩擦损耗、液压系统节流损失是装备运行能耗的主要组成",
        "**研究方向**：低摩擦运动副设计与润滑；再制造零部件的性能恢复与寿命评估；零部件级 LCA 方法与数据库；面向回收的材料与结构设计",
    ], size=13.5)

    s = content_slide("4.11 颠覆性技术展望", "可能重塑零部件形态与制造方式的远期方向（10–20 年）",
                      notes=("1.5 分钟", "简要展望，明确这些属于远期、不确定性大的方向。鼓励学术界关注。"))
    add_cards(s, [
        ("原子级制造", "单原子操纵、原子层去除/沉积（ALE/ALD）\n面向 EUV 光学、量子器件的原子级精度\n从“近原子尺度制造”到“原子制造”"),
        ("量子传感与计量", "原子干涉、NV 色心磁强计、光钟\n超越经典极限的位移、加速度、角速度测量\n未来的纳米定位与导航基准"),
        ("自修复与自适应材料", "微胶囊/血管网络自修复、可逆共价键\n可变刚度/可变阻尼材料\n零部件“自愈”延寿的可能性"),
        ("软体与仿生零部件", "软体执行器、仿生关节与肌腱\n人形机器人与医疗装备驱动\n传统“刚性精密”范式之外的新路径"),
        ("数字材料与可编程物质", "体素化多材料增材、可编程机械超材料\n按需设计刚度、泊松比、热膨胀\n“结构即功能”的零部件"),
        ("全自主研发系统", "AI + 机器人实验室的闭环材料/工艺研发\n自主设计—制造—测试—迭代\n零部件研发周期的数量级缩短"),
    ], cols=3, body_size=12, height=Inches(4.5))
    add_text(s, MARGIN, Inches(6.0), CONTENT_W, Inches(0.8),
             "注：以上方向多处于基础研究或概念验证阶段，不确定性高。学术界的价值恰在于提前布局高风险、高回报方向，为 10–20 年后的产业变革储备知识与人才。",
             size=12.5, color=GRAY, italic=True)


# =====================================================================
# 第五部分 趋势与建议
# =====================================================================
def part5():
    section_slide(4, "收束全篇：趋势总结、技术路线图、科学问题清单、产学研建议、总结与讨论。")

    s = content_slide("5.1 技术发展趋势总览", "六化并进：精密化、高速化、智能化、集成化、极端化、绿色化",
                      notes=("2 分钟", "总结六大趋势，并说明它们如何相互强化（如智能化支撑精密化与可靠性；极端化牵引材料与工艺）。"))
    trends = [("精密化", "μm → nm → 原子级；精度保持性成为核心", TEAL),
              ("高速化", "更高 DN 值、更高加速度、更高带宽；热与动力学的极限", NAVY_2),
              ("智能化", "自感知—自诊断—自适应；AI 融入设计、制造与运维", GREEN),
              ("集成化", "机电液控一体、部件即系统（关节模组、智能主轴）", AMBER),
              ("极端化", "真空、洁净、超高/低温、强辐射；极端制造牵引", TEAL),
              ("绿色化", "高效、长寿命、可再制造、全生命周期低碳", NAVY_2)]
    for i, (t, d, c) in enumerate(trends):
        r, col = divmod(i, 3)
        x = MARGIN + col * (Inches(4.0) + Inches(0.1))
        y = CONTENT_TOP + r * Inches(2.2)
        add_rect(s, x, y, Inches(4.0), Inches(2.0), WHITE, line=GRAY_MID)
        add_rect(s, x, y, Inches(0.12), Inches(2.0), c)
        add_text(s, x + Inches(0.3), y + Inches(0.15), Inches(3.5), Inches(0.6), t, size=22, bold=True, color=NAVY_2)
        add_text(s, x + Inches(0.3), y + Inches(0.8), Inches(3.5), Inches(1.1), d, size=13.5, color=BLACK, line_spacing=1.15)
    add_callout(s, "底层驱动力：下游装备（半导体、新能源、航空航天、机器人）需求升级 + 材料/传感/计算技术进步 + 自主可控的战略需求",
                top=Inches(5.95), height=Inches(0.6), size=12.5)

    s = content_slide("5.2 技术路线图（2025—2035，示意）", "分阶段目标：补短板 → 锻长板 → 引领",
                      notes=("2 分钟", "路线图为报告人基于公开资料的综合判断，供讨论。强调三阶段的不同重点。"))
    phases = [("2025–2027  补短板", TEAL, ["高端轴承、RV 减速器、高端数控系统的批量一致性与可靠性达到国际主流水平",
                                          "建成零部件试验验证与数据平台；标准体系完善", "智能零部件商业化初期；增材零部件工程应用"]),
              ("2028–2030  锻长板", NAVY_2, ["半导体装备核心运动部件（工件台、真空传输）自主可控",
                                           "智能主轴/轴承/关节模组规模应用；PHM 与数字孪生成熟", "新材料（陶瓷、复合材料）零部件工程化；AI 驱动研发普及"]),
              ("2031–2035  引领", AMBER, ["在若干细分领域（磁悬浮、超精密、机器人关节）形成国际领先产品与标准",
                                        "原子级制造、量子计量等颠覆性技术开始工程验证", "零部件研发范式转向“自主智能研发系统”"])]
    for i, (name, c, items) in enumerate(phases):
        x = MARGIN + i * (Inches(4.0) + Inches(0.1))
        hdr = add_rect(s, x, CONTENT_TOP, Inches(4.0), Inches(0.6), c)
        shape_text(hdr, name, size=15, bold=True, color=WHITE if c != AMBER else NAVY)
        add_rect(s, x, CONTENT_TOP + Inches(0.6), Inches(4.0), Inches(3.9), GRAY_LIGHT)
        add_bullets(s, x + Inches(0.1), CONTENT_TOP + Inches(0.75), Inches(3.8), Inches(3.7), items, size=12.5, para_space=8)
    add_line(s, MARGIN, CONTENT_TOP + Inches(4.7), W - MARGIN, CONTENT_TOP + Inches(4.7), color=TEAL, width=3)
    add_rect(s, W - MARGIN - Inches(0.3), CONTENT_TOP + Inches(4.55), Inches(0.3), Inches(0.3), TEAL, shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
    add_text(s, MARGIN, CONTENT_TOP + Inches(4.85), CONTENT_W, Inches(0.7),
             "贯穿始终：基础研究（机理、材料、测量）持续投入；人才培养；国际合作与标准参与。路线图为示意，具体节点需结合产业实际动态调整。",
             size=12.5, color=GRAY, italic=True)

    s = content_slide("5.3 值得关注的基础科学问题", "面向学术界：零部件研发背后的“硬核”科学问题清单",
                      notes=("2.5 分钟", "这是面向学术听众最重要的一页。逐条简述，鼓励听众从中寻找与自身方向的结合点。"))
    add_table(s, [
        ["领域", "科学问题", "涉及零部件"],
        ["接触力学与摩擦学", "滚动接触疲劳的多尺度起裂与扩展机理；极端工况（高速/真空/低温）润滑失效机理；表面完整性—服役性能的定量关系", "轴承、丝杠、减速器、液压副"],
        ["多场耦合动力学", "热—结构—电磁—流体耦合下高速转子/运动系统的非线性动力学与稳定性；结合面特性辨识", "电主轴、磁悬浮、直驱系统"],
        ["精度理论", "误差生成—传递—演化（退化）的统一建模；误差均化与误差分离的极限；精度保持性的预测", "所有精密零部件与整机"],
        ["材料科学", "夹杂物/缺陷—疲劳寿命的概率关系；组织—性能—工艺的 ICME；新材料（HEA、CMC）的服役行为", "轴承钢、齿轮钢、陶瓷"],
        ["测量与感知", "纳米/亚纳米测量的新原理与溯源；嵌入式弱信号传感与解耦；多自由度同步测量", "光栅、智能零部件、工件台"],
        ["智能与信息", "物理—数据融合建模的理论框架；小样本可靠性/寿命推断；可解释、可泛化的工业 AI", "PHM、数字孪生、AI 设计"],
        ["制造科学", "纳米尺度材料去除与亚表面损伤机理；增材工艺—组织—缺陷—性能链；面向性能的装配理论", "超精密零件、增材零部件"],
    ], col_widths=[1.7, 7.0, 3.0], font_size=11, row_h=Inches(0.62))
    add_callout(s, "这些问题的共同特征：跨尺度（nm–m）、跨学科（力/热/电磁/材料/信息）、跨阶段（设计—制造—服役）",
                top=Inches(6.2), height=Inches(0.5), size=12.5)

    s = content_slide("5.4 产学研协同与成果转化模式", "零部件研发的特殊性决定了需要长期、深度、互信的合作机制",
                      notes=("2 分钟", "讨论产学研合作的几种有效模式和常见问题，结合国内外经验（如德国 Fraunhofer、WZL 模式；国内创新中心、04 专项“一条龙”）。"))
    add_cards(s, [
        ("联合实验室 / 创新中心", "高校机理研究 + 企业工程化 + 整机厂应用验证\n共享试验平台与数据库（如轴承、减速器寿命台）\n国内：国家制造业创新中心、省级共性技术平台\n国际参照：Fraunhofer IPT/IWU、RWTH WZL、Berkeley/Michigan 联盟"),
        ("“一条龙”应用示范", "整机厂承诺装机验证 → 零部件厂迭代 → 高校提供测试与分析\n破解“首台套”“不敢用”难题\n04 专项经验：应用牵引下的技术成熟度提升"),
        ("人才双向流动", "工程博士、联合培养研究生、企业导师\n高校教师企业挂职、企业工程师进实验室\n零部件研发经验高度依赖“人”的隐性知识"),
        ("标准、专利与数据", "共同参与标准制定，形成技术话语权\n专利池与交叉许可\n预竞争阶段的数据共享机制（试验数据、失效案例）"),
    ], cols=2, body_size=12, height=Inches(4.3))
    add_callout(s, "关键认识：零部件研发回报周期长（5–10 年），需要“耐心资本”与长期考核机制；短期项目制难以积累工程数据与隐性知识",
                top=Inches(5.85), height=Inches(0.6), size=12.5)

    s = content_slide("5.5 对研究团队的几点建议", "如何在零部件研究中做出有学术深度又有产业价值的工作",
                      notes=("2 分钟", "以同行视角分享建议，语气平实。强调“顶天立地”：既要科学问题，也要工程落地；重视测试平台与数据积累；重视交叉。"))
    tips = [
        ("选题：从真实瓶颈中提炼科学问题", "深入企业现场，理解失效模式与工艺难点；将“做不好”转化为可研究的机理/模型/方法问题，避免“伪问题”"),
        ("方法：机理 + 数据 + 试验三位一体", "不迷信仿真或 AI；建立自己的试验能力（哪怕是小型台架）；重视测量不确定度与数据质量"),
        ("积累：建设平台与数据库", "零部件研究的壁垒是数据与经验；坚持方向 5–10 年，形成可复用的试验平台、模型工具与数据资产"),
        ("交叉：主动跨学科", "摩擦学 × 控制 × 材料 × 信息；与数学、物理、材料学院合作；关注半导体、机器人等新需求"),
        ("转化：以“可用”为目标", "成果形态可以是模型/软件/标准/测试方法，不必都是“产品”；与企业共同定义验收指标"),
        ("视野：关注国际前沿与标准", "跟踪 CIRP、ASPE、euspen、ISO/TC 等；积极参与国际学术共同体与标准组织"),
    ]
    for i, (t, d) in enumerate(tips):
        r, c = divmod(i, 2)
        x = MARGIN + c * (Inches(6.05) + Inches(0.15))
        y = CONTENT_TOP + r * Inches(1.75)
        add_rect(s, x, y, Inches(6.05), Inches(1.6), GRAY_LIGHT)
        num = add_rect(s, x + Inches(0.12), y + Inches(0.15), Inches(0.5), Inches(0.5), TEAL if i % 2 == 0 else AMBER, shape=MSO_SHAPE.OVAL)
        shape_text(num, str(i + 1), size=14, bold=True, color=WHITE if i % 2 == 0 else NAVY)
        add_text(s, x + Inches(0.75), y + Inches(0.1), Inches(5.2), Inches(0.5), t, size=14.5, bold=True, color=NAVY_2)
        add_text(s, x + Inches(0.75), y + Inches(0.6), Inches(5.2), Inches(1.0), d, size=12, color=BLACK, line_spacing=1.1)

    s = content_slide("总结", "三句话带走",
                      notes=("2 分钟", "总结三个核心观点。给听众留下清晰印象。"))
    msgs = [("1", "关键零部件是先进制造装备的“性能之源、可靠之本”",
             "装备的精度、速度、可靠性最终由轴承、传动、驱动、控制、感知等基础零部件决定；差距的本质是材料—工艺—检测—数据的工程积累，而非单点原理。"),
            ("2", "共性使能技术是突破口，也是学术界的主战场",
             "超精密加工、精密装配、材料热处理、表面工程、多场耦合设计、可靠性工程、测试评价与数字化研发贯穿所有零部件，蕴含大量跨尺度、跨学科的科学问题。"),
            ("3", "前沿技术正在重塑零部件的形态与研发范式",
             "智能化、增材、非接触支承、新材料、AI、PHM、新型驱动与极端制造，使零部件从“被动机械要素”走向“主动信息物理单元”，研发从“经验试错”走向“模型与数据驱动”。")]
    for i, (n, t, d) in enumerate(msgs):
        y = CONTENT_TOP + i * Inches(1.65)
        add_rect(s, MARGIN, y, Inches(0.9), Inches(1.45), [TEAL, NAVY_2, AMBER][i])
        add_text(s, MARGIN, y, Inches(0.9), Inches(1.45), n, size=36, bold=True,
                 color=WHITE if i != 2 else NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, MARGIN + Inches(0.9), y, CONTENT_W - Inches(0.9), Inches(1.45), GRAY_LIGHT)
        add_text(s, MARGIN + Inches(1.1), y + Inches(0.1), CONTENT_W - Inches(1.3), Inches(0.5), t, size=17, bold=True, color=NAVY_2)
        add_text(s, MARGIN + Inches(1.1), y + Inches(0.6), CONTENT_W - Inches(1.3), Inches(0.85), d, size=12.5, color=BLACK, line_spacing=1.12)
    add_callout(s, "“把简单的事情做到极致，就是高端” —— 零部件研发没有捷径，但有方法：机理、数据、平台与耐心",
                top=Inches(6.35), height=Inches(0.5), size=13)

    s = content_slide("参考文献与推荐阅读（示例）", "报告中数据与观点的主要来源类型；请以最新版本为准",
                      notes=("0.5 分钟", "不必逐条讲，告知听众资料类型与获取途径。"))
    add_bullets(s, MARGIN, CONTENT_TOP, CONTENT_W, Inches(5.5), [
        "**政策与规划**：《中国制造 2025》（2015）；《工业强基工程实施指南（2016—2020）》；《“十四五”智能制造发展规划》（2021）；“高档数控机床与基础制造装备”科技重大专项相关总结材料",
        "**学术期刊**：CIRP Annals – Manufacturing Technology；Precision Engineering；International Journal of Machine Tools and Manufacture；Tribology International；Mechanical Systems and Signal Processing；机械工程学报；中国机械工程",
        "**学术组织与会议**：CIRP（国际生产工程科学院）；ASPE / euspen（精密工程）；ISO/TC 39（机床）、TC 4（滚动轴承）、TC 60（齿轮）、TC 213（产品几何技术规范）",
        "**行业报告**：中国机床工具工业协会、中国轴承工业协会、中国机器人产业联盟年度报告；国际机器人联合会（IFR）World Robotics",
        "**经典著作**：Harris & Kotzalas, Rolling Bearing Analysis；Slocum, Precision Machine Design；Altintas, Manufacturing Automation；Schweitzer & Maslen, Magnetic Bearings",
        "**说明**：报告中的国产化率、成本占比、性能指标等为公开资料整理的大致区间，用于说明趋势，具体数值请以权威统计与最新文献为准",
    ], size=13.5, para_space=9)

    # 致谢 / Q&A
    _state["page"] += 1
    s = prs.slides.add_slide(BLANK)
    add_rect(s, 0, 0, W, H, NAVY)
    add_rect(s, 0, Inches(5.2), W, Inches(2.3), NAVY_2)
    for i in range(6):
        add_rect(s, W - Inches(3.2) + i * Inches(0.45), Inches(0.5) + i * Inches(0.3),
                 Inches(0.3), Inches(0.3), TEAL if i % 2 == 0 else AMBER, shape=MSO_SHAPE.OVAL)
    add_text(s, MARGIN, Inches(2.0), Inches(12), Inches(1.2), "谢谢！欢迎讨论", size=48, bold=True, color=WHITE)
    add_text(s, MARGIN, Inches(3.3), Inches(12), Inches(0.6), "Thank you — Questions & Discussion", size=20, color=STEEL, italic=True)
    add_rect(s, MARGIN, Inches(4.1), Inches(1.6), Inches(0.06), AMBER)
    add_text(s, MARGIN, Inches(5.55), Inches(12), Inches(0.5), "讨论议题建议：", size=15, bold=True, color=TEAL_LIGHT)
    add_text(s, MARGIN, Inches(6.0), Inches(12), Inches(1.2),
             "① 您所在领域中，哪些零部件瓶颈最亟待学术界介入？    ② 产学研长期合作机制如何设计？\n③ AI 与数字孪生在零部件研发中的边界在哪里？    ④ 极端制造需求对基础研究提出了哪些新问题？",
             size=13, color=WHITE, line_spacing=1.2)
    add_notes(s, "10 分钟", "进入讨论环节。可用上述四个议题引导。总结讨论要点并致谢。")


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "先进制造装备关键零部件研发及前沿技术.pptx"
    cover()
    agenda()
    part1()
    part2()
    part3()
    part4()
    part5()
    prs.save(str(out))
    print(f"saved: {out}  slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
