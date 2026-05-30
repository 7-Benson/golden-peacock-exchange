#!/usr/bin/env python3
"""Generate Golden Peacock Exchange 16-slide PPT"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

prs = Presentation()
prs.slide_width = Inches(13.333)  # 16:9
prs.slide_height = Inches(7.5)

# ── Color Palette ──
GOLD = RGBColor(0xD4, 0xA5, 0x37)       # 金色主色
GOLD_LIGHT = RGBColor(0xF0, 0xD0, 0x60)  # 浅金
GOLD_DARK = RGBColor(0xB8, 0x86, 0x0B)   # 暗金
DARK = RGBColor(0x1A, 0x1A, 0x2E)        # 深蓝黑
DARK2 = RGBColor(0x16, 0x21, 0x38)       # 更深的背景
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xCC, 0xCC, 0xCC)
ACCENT_GREEN = RGBColor(0x2E, 0xCC, 0x71)
ACCENT_BLUE = RGBColor(0x54, 0x9B, 0xFF)
ACCENT_RED = RGBColor(0xE7, 0x4C, 0x3C)


def set_slide_bg(slide, color=DARK2):
    """Set slide background to solid color."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_shape_bg(slide, left, top, width, height, color, alpha=None):
    """Add a colored rectangle shape as background element."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    if alpha is not None:
        # python-pptx doesn't support alpha directly; skip
        pass
    return shape


def add_text_box(slide, left, top, width, height, text, font_size=18, color=WHITE,
                 bold=False, alignment=PP_ALIGN.LEFT, font_name='Microsoft YaHei'):
    """Add a text box with single paragraph."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def add_rich_text_box(slide, left, top, width, height, lines, font_name='Microsoft YaHei'):
    """Add a text box with multiple paragraphs of different styles.
    lines: list of (text, font_size, color, bold, alignment) tuples
    """
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line_data in enumerate(lines):
        text = line_data[0]
        font_size = line_data[1] if len(line_data) > 1 else 18
        color = line_data[2] if len(line_data) > 2 else WHITE
        bold = line_data[3] if len(line_data) > 3 else False
        alignment = line_data[4] if len(line_data) > 4 else PP_ALIGN.LEFT

        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = text
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.font.name = font_name
        p.alignment = alignment
        p.space_after = Pt(6)
    return txBox


def add_gold_accent_line(slide, left, top, width, height=Pt(3)):
    """Add a thin gold line."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = GOLD
    shape.line.fill.background()
    return shape


def add_card(slide, left, top, width, height, title, body, title_color=GOLD, bg_color=RGBColor(0x22, 0x2E, 0x44)):
    """Add a card with title and body text."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = GOLD_DARK
    card.line.width = Pt(1)

    # Title
    add_text_box(slide, left + Inches(0.3), top + Inches(0.2), width - Inches(0.6), Inches(0.5),
                 title, font_size=16, color=title_color, bold=True)
    # Body
    add_text_box(slide, left + Inches(0.3), top + Inches(0.7), width - Inches(0.6), height - Inches(0.9),
                 body, font_size=13, color=LIGHT_GRAY)
    return card


def add_bullet_list(slide, left, top, width, height, items, font_size=16, color=LIGHT_GRAY, bullet_char="▸"):
    """Add a bulleted list."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = f"{bullet_char} {item}"
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = 'Microsoft YaHei'
        p.space_after = Pt(8)
    return tb


def add_section_number(slide, num, total=16):
    """Add page number at bottom right."""
    add_text_box(slide, Inches(12.0), Inches(7.0), Inches(1.2), Inches(0.4),
                 f"{num}/{total}", font_size=11, color=GOLD_DARK, alignment=PP_ALIGN.RIGHT)


def add_header_bar(slide, title, subtitle=None, page_num=None):
    """Add a standard header with gold accent line."""
    # Top gold line
    add_shape_bg(slide, Inches(0), Inches(0), Inches(13.333), Pt(4), GOLD)
    # Title
    add_text_box(slide, Inches(0.8), Inches(0.4), Inches(10), Inches(0.7),
                 title, font_size=32, color=GOLD, bold=True)
    if subtitle:
        add_text_box(slide, Inches(0.8), Inches(1.0), Inches(10), Inches(0.5),
                     subtitle, font_size=16, color=LIGHT_GRAY)
    # Bottom accent
    add_gold_accent_line(slide, Inches(0.8), Inches(1.4), Inches(3), Pt(2))
    if page_num:
        add_section_number(slide, page_num)


# ═══════════════════════════════════════════════
# SLIDE 1: COVER
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
set_slide_bg(slide, DARK)

# Decorative gold rectangle top
add_shape_bg(slide, Inches(0), Inches(0), Inches(13.333), Inches(2.8), RGBColor(0x12, 0x1A, 0x30))

# Big gold accent bar
add_shape_bg(slide, Inches(0), Inches(2.8), Inches(13.333), Pt(5), GOLD)

# Gold decorative circle (abstract peacock)
circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.5), Inches(1.0), Inches(2.3), Inches(2.3))
circ.fill.solid()
circ.fill.fore_color.rgb = GOLD
circ.line.fill.background()
# Inner circle
circ2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.9), Inches(1.4), Inches(1.5), Inches(1.5))
circ2.fill.solid()
circ2.fill.fore_color.rgb = GOLD_DARK
circ2.line.fill.background()
# Peacock "eye" dot
circ3 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.4), Inches(1.9), Inches(0.5), Inches(0.5))
circ3.fill.solid()
circ3.fill.fore_color.rgb = WHITE
circ3.line.fill.background()

# Title
add_text_box(slide, Inches(1), Inches(3.5), Inches(11.3), Inches(1.2),
             "金孔雀交易所", font_size=54, color=GOLD, bold=True,
             alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(4.5), Inches(11.3), Inches(0.8),
             "GOLDEN PEACOCK EXCHANGE", font_size=28, color=WHITE, bold=False,
             alignment=PP_ALIGN.CENTER)

# Gold separator
add_shape_bg(slide, Inches(5), Inches(5.3), Inches(3.333), Pt(3), GOLD)

# Subtitle
add_text_box(slide, Inches(1), Inches(5.6), Inches(11.3), Inches(0.6),
             "东南亚RWA资产发行平台  —  泰国房地产RWA首发", font_size=20, color=LIGHT_GRAY,
             alignment=PP_ALIGN.CENTER)

# Tagline
add_text_box(slide, Inches(1), Inches(6.3), Inches(11.3), Inches(0.5),
             "链接实体资产 · 开启RWA新纪元", font_size=16, color=GOLD_LIGHT,
             alignment=PP_ALIGN.CENTER)

# Bottom gold bar
add_shape_bg(slide, Inches(0), Inches(7.1), Inches(13.333), Pt(4), GOLD)

add_section_number(slide, 1)

# ═══════════════════════════════════════════════
# SLIDE 2: PROJECT OVERVIEW
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "项目概览", "Project Overview", 2)

# Positioning statement
add_shape_bg(slide, Inches(0.8), Inches(1.7), Inches(11.7), Inches(0.8), RGBColor(0x22, 0x2E, 0x44))
add_text_box(slide, Inches(1.0), Inches(1.8), Inches(11.3), Inches(0.6),
             "东南亚首个专注房地产RWA的合规资产发行与交易平台", font_size=20, color=GOLD_LIGHT, bold=True)

# Three key points
card_data = [
    ("🎯 首个标的", "泰国芭提雅核心区\n房地产RWA项目", Inches(0.8)),
    ("💰 发行规模", "具体金额待定\n（根据资产估值确定）", Inches(4.8)),
    ("🏢 资产方", "环球国际资产管理公司\n优质底层资产提供方", Inches(8.8)),
]
for title, body, left in card_data:
    add_card(slide, left, Inches(2.8), Inches(3.6), Inches(1.8), title, body)

# Vision
add_text_box(slide, Inches(0.8), Inches(5.0), Inches(11.7), Inches(0.4),
             "项目愿景", font_size=20, color=GOLD, bold=True)
add_text_box(slide, Inches(0.8), Inches(5.5), Inches(11.7), Inches(0.8),
             "让全球投资者低门槛参与东南亚优质不动产投资，实现资产跨境自由流动",
             font_size=17, color=LIGHT_GRAY)


# ═══════════════════════════════════════════════
# SLIDE 3: RWA MARKET BACKGROUND
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "行业背景 — RWA市场爆发", "Real World Assets: The Next Crypto Frontier", 3)

# RWA Definition
add_card(slide, Inches(0.8), Inches(1.8), Inches(5.5), Inches(2.0),
         "什么是RWA？",
         "Real World Assets（真实世界资产）\n将房地产、债券、大宗商品等实体\n资产进行代币化并上链流通")

# Data card
add_card(slide, Inches(6.8), Inches(1.8), Inches(5.5), Inches(2.0),
         "市场数据",
         "• RWA市场2025年突破多亿美元规模\n• 全球代币化资产预计2030年\n  达数万亿至十万亿美元级别\n• 头部机构纷纷布局")

# Why RWA
add_card(slide, Inches(0.8), Inches(4.2), Inches(5.5), Inches(2.5),
         "为什么是RWA？",
         "✅ 传统资产流动性差 → 代币化提升流动性\n✅ 投资门槛高 → 碎片化降低门槛\n✅ 跨境交易难 → 链上全球自由流通\n✅ 透明度低 → 智能合约自动执行")

# Drivers
add_card(slide, Inches(6.8), Inches(4.2), Inches(5.5), Inches(2.5),
         "核心驱动力",
         "🚀 区块链技术成熟\n📋 各国监管政策清晰化\n🌍 全球通胀下的资产配置需求\n🏦 机构资金大举入场")


# ═══════════════════════════════════════════════
# SLIDE 4: MARKET OPPORTUNITY - SEA
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "市场机遇 — 东南亚RWA蓝海", "Southeast Asia: The Untapped RWA Market", 4)

# SEA advantages
add_card(slide, Inches(0.8), Inches(1.8), Inches(5.5), Inches(2.3),
         "东南亚独特优势",
         "📈 高经济增长 + 持续外资流入\n✈️ 旅游业强劲复苏\n🏠 泰国房地产对外国人购买有限制\n   → 代币化提供合规投资通道")

# Pattaya basics
add_card(slide, Inches(6.8), Inches(1.8), Inches(5.5), Inches(2.3),
         "芭提雅房产基本面",
         "🌊 年游客超千万人次\n💰 租金回报率 4%-8%\n🏖️ 一线海景公寓需求旺盛\n🛣️ 基建持续升级中")

# Market gap
add_shape_bg(slide, Inches(0.8), Inches(4.5), Inches(11.7), Inches(1.2), RGBColor(0x22, 0x2E, 0x44))
add_text_box(slide, Inches(1.0), Inches(4.6), Inches(11.3), Inches(0.4),
             "空白市场机会", font_size=20, color=GOLD, bold=True)
add_text_box(slide, Inches(1.0), Inches(5.0), Inches(11.3), Inches(0.5),
             "东南亚专业RWA交易所稀缺 → 金孔雀交易所抢占先机，成为区域首家专注房地产RWA的合规平台",
             font_size=17, color=LIGHT_GRAY)

# Opportunity size callout
add_shape_bg(slide, Inches(4.0), Inches(6.0), Inches(5.3), Inches(0.7), GOLD_DARK)
add_text_box(slide, Inches(4.0), Inches(6.05), Inches(5.3), Inches(0.6),
             "先发优势 — 蓝海市场第一梯队",
             font_size=18, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════
# SLIDE 5: PLATFORM POSITIONING
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "平台定位与核心价值", "Platform Positioning & Value Proposition", 5)

# Positioning
add_card(slide, Inches(0.8), Inches(1.8), Inches(11.7), Inches(1.3),
         "金孔雀交易所定位",
         "RWA资产的「发行 + 交易 + 管理」一体化平台 — 专注东南亚房地产RWA赛道")

# Value props
props = [
    ("🔑", "合规", "持牌运营\n严格KYC/AML体系", Inches(0.8)),
    ("🔑", "真实", "资产严格尽调\n链上锚定线下产权", Inches(3.8)),
    ("🔑", "低门槛", "最小投资单位大幅降低\n碎片化普惠投资", Inches(6.8)),
    ("🔑", "高流动性", "二级市场交易\nT+0灵活退出机制", Inches(9.8)),
]
for icon, title, desc, left in props:
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(3.5), Inches(2.7), Inches(2.5))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0x22, 0x2E, 0x44)
    card.line.color.rgb = GOLD_DARK
    card.line.width = Pt(1)
    add_text_box(slide, left + Inches(0.3), Inches(3.7), Inches(2.1), Inches(0.5),
                 f"{icon} {title}", font_size=20, color=GOLD, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, left + Inches(0.2), Inches(4.3), Inches(2.3), Inches(1.5),
                 desc, font_size=14, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════
# SLIDE 6: FIRST ASSET
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "首个标的 — 芭提雅房地产RWA", "First RWA Asset: Pattaya Real Estate", 6)

# Asset details
add_card(slide, Inches(0.8), Inches(1.8), Inches(5.8), Inches(1.8),
         "标的资产信息",
         "📍 位置：芭提雅核心区（中天海滩 / 帕山区 / 纳哥路）\n🏠 类型：高端海景公寓 / 度假酒店综合项目\n📐 规模：待具体确认（建筑面积 / 单位数量 / 估值）")

add_card(slide, Inches(7.0), Inches(1.8), Inches(5.5), Inches(1.8),
         "芭提雅区位优势",
         "🚄 曼谷EEC高铁1小时直达\n✈️ 乌塔堡机场扩建中\n🏗️ 大型基建项目密集落地\n🛍️ 商业配套持续升级")

# Asset highlights
add_text_box(slide, Inches(0.8), Inches(4.0), Inches(11.7), Inches(0.4),
             "资产亮点", font_size=20, color=GOLD, bold=True)

highlights = [
    ("💰", "稳定租金现金流", "旅游复苏带动入住率\n预期年化租金回报 5-8%"),
    ("📈", "长期增值潜力", "芭提雅升级为城市\n土地价值持续攀升"),
    ("⚖️", "产权清晰", "法律结构完善\nSPV持有产权"),
    ("🛡️", "专业管理", "环球国际资产管理\n全流程资产服务"),
]
for i, (icon, title, desc, ) in enumerate(highlights):
    left = Inches(0.8 + i * 3.1)
    add_card(slide, left, Inches(4.5), Inches(2.8), Inches(2.2), f"{icon} {title}", desc)


# ═══════════════════════════════════════════════
# SLIDE 7: ASSET PARTNER
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "资产方 — 环球国际资产管理公司", "Asset Partner: Global International Asset Management", 7)

# Company intro
add_card(slide, Inches(0.8), Inches(1.8), Inches(5.5), Inches(1.8),
         "公司简介",
         "环球国际资产管理公司\n深耕泰国不动产领域多年\n拥有丰富的本地资源和项目经验")

# Core capabilities
add_card(slide, Inches(6.8), Inches(1.8), Inches(5.5), Inches(1.8),
         "核心能力",
         "🌐 本地资源：泰国房地产收购与开发经验丰富\n✅ 合规能力：熟悉BOI、土地法、外商投资法规\n🏗️ 项目储备：芭提雅多个优质项目资源")

# Cooperation model
add_card(slide, Inches(0.8), Inches(4.0), Inches(5.5), Inches(1.5),
         "合作模式",
         "作为RWA的底层资产提供方与资产服务方\n负责标的筛选、收购及后续管理")

# Endorsements
add_card(slide, Inches(6.8), Inches(4.0), Inches(5.5), Inches(1.5),
         "资质背书",
         "• 泰国本地相关运营牌照\n• 业内知名合作伙伴\n• 成熟的项目管理团队")


# ═══════════════════════════════════════════════
# SLIDE 8: RWA PRODUCT STRUCTURE
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "RWA产品结构", "RWA Product Structure", 8)

# Flow steps
steps = [
    ("1", "资产收购", "筛选并收购\n优质标的资产"),
    ("2", "法律确权", "确权+法律意见书\n产权清晰"),
    ("3", "资产估值", "第三方独立\n专业估值"),
    ("4", "SPV架构", "SPV持有资产\n隔离风险"),
    ("5", "代币发行", "RWA代币化\n合规发行"),
    ("6", "上架交易", "二级市场\nT+0交易"),
]
for i, (num, title, desc) in enumerate(steps):
    left = Inches(0.5 + i * 2.1)
    # Step circle
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(0.5), Inches(1.8), Inches(0.7), Inches(0.7))
    circ.fill.solid()
    circ.fill.fore_color.rgb = GOLD if int(num) <= 2 else GOLD_DARK
    circ.line.fill.background()
    add_text_box(slide, left + Inches(0.5), Inches(1.85), Inches(0.7), Inches(0.6),
                 num, font_size=22, color=DARK, bold=True, alignment=PP_ALIGN.CENTER)
    # Arrow (except last)
    if i < len(steps) - 1:
        add_text_box(slide, left + Inches(1.3), Inches(1.9), Inches(0.5), Inches(0.5),
                     "→", font_size=24, color=GOLD_LIGHT, alignment=PP_ALIGN.CENTER)
    # Title
    add_text_box(slide, left, Inches(2.7), Inches(1.9), Inches(0.4), title,
                 font_size=14, color=GOLD, bold=True, alignment=PP_ALIGN.CENTER)
    # Desc
    add_text_box(slide, left, Inches(3.1), Inches(1.9), Inches(0.7), desc,
                 font_size=11, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

# Token design
add_shape_bg(slide, Inches(0.8), Inches(4.0), Inches(11.7), Inches(0.6), RGBColor(0x22, 0x2E, 0x44))
add_text_box(slide, Inches(1.0), Inches(4.05), Inches(11.3), Inches(0.5),
             "代币设计", font_size=20, color=GOLD, bold=True)

token_features = [
    "每枚代币对应实物资产的一定份额（如1㎡房产权益）",
    "收益分为：租金分红 + 资产增值分红",
    "基于币安链智能合约自动分配收益，透明可查",
    "法律结构：SPV持有资产，代币代表SPV权益份额",
    "BEP-20代币标准，兼容主流DEX/钱包生态",
]
add_bullet_list(slide, Inches(1.0), Inches(4.7), Inches(11.3), Inches(2.0), token_features, font_size=15)


# ═══════════════════════════════════════════════
# SLIDE 9: TECH ARCHITECTURE
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "技术架构", "Technology Architecture", 9)

# Blockchain
add_card(slide, Inches(0.8), Inches(1.8), Inches(11.7), Inches(1.5),
         "底层公链 — 币安链（BNB Chain）",
         "采用BNB Chain（币安智能链）作为底层基础设施\n"
         "• BEP-20代币标准，高吞吐量（~300 TPS），低交易费（<0.1$）\n"
         "• 成熟的DeFi生态与DEX基础设施，支持PancakeSwap等主流协议集成\n"
         "• 全球节点数最多（40+验证节点），社群活跃，开发者生态完善")

# Core modules
modules = [
    ("📦", "资产发行模块", "RWA代币化发行工具\nBEP-20合规代币铸造"),
    ("🔄", "交易模块", "订单簿 + DEX混合撮合\n支持PancakeSwap流动性"),
    ("🔐", "资产管理模块", "KYC / 合规审查\n链上资产存证"),
    ("💸", "收益分发模块", "智能合约自动分红\n链上实时可查"),
]
for i, (icon, title, desc) in enumerate(modules):
    left = Inches(0.8 + i * 3.1)
    add_card(slide, left, Inches(3.6), Inches(2.8), Inches(2.0), f"{icon} {title}", desc)

# Security
add_card(slide, Inches(0.8), Inches(5.8), Inches(11.7), Inches(0.8),
         "安全审计",
         "第三方智能合约代码审计  |  币安链生态安全标准  |  多重签名钱包  |  定期安全巡检")


# ═══════════════════════════════════════════════
# SLIDE 10: COMPLIANCE
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "合规与监管框架", "Regulatory & Compliance Framework", 10)

# Compliance strategy
add_card(slide, Inches(0.8), Inches(1.8), Inches(5.8), Inches(2.2),
         "监管合规策略",
         "🛡️ 持有或申请目标地区合规牌照\n🔍 严格KYC/AML体系（三级身份认证）\n📊 每季度资产审计与透明度报告\n👮 反洗钱内控制度完善")

# Legal
add_card(slide, Inches(6.8), Inches(1.8), Inches(5.8), Inches(2.2),
         "法律保障",
         "⚖️ 泰国本地律所出具法律意见书\n📋 SPV架构满足外商投资法规\n👥 投资者权益保护机制完善\n🌏 符合国际合规标准")

# Certification badges
badges = ["KYC/AML合规", "季度审计", "法律意见书", "投资者保护"]
for i, badge in enumerate(badges):
    left = Inches(0.8 + i * 3.1)
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(4.5), Inches(2.8), Inches(0.8))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0x1A, 0x2A, 0x3E)
    box.line.color.rgb = GOLD
    box.line.width = Pt(1)
    add_text_box(slide, left, Inches(4.6), Inches(2.8), Inches(0.6),
                 f"✅ {badge}", font_size=16, color=ACCENT_GREEN, bold=True, alignment=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════
# SLIDE 11: TOKENOMICS
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "代币经济模型", "Tokenomics", 11)

# Token basic info
add_card(slide, Inches(0.8), Inches(1.8), Inches(11.7), Inches(0.8),
         "代币符号：待定（如 GPC / GPE）",
         "平台通证，承载交易、治理与生态价值")

# Allocation
allocations = [
    ("房地产RWA发行池", "60%", Inches(0.8)),
    ("流动性储备", "15%", Inches(4.5)),
    ("团队/运营", "10%", Inches(8.2)),
]
for name, pct, left in allocations:
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(2.9), Inches(3.3), Inches(1.3))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0x1A, 0x2A, 0x3E)
    card.line.color.rgb = GOLD
    card.line.width = Pt(1)
    add_text_box(slide, left, Inches(3.0), Inches(3.3), Inches(0.4),
                 pct, font_size=28, color=GOLD, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, left, Inches(3.5), Inches(3.3), Inches(0.4),
                 name, font_size=14, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

# Remaining allocations
add_text_box(slide, Inches(0.8), Inches(4.5), Inches(5.5), Inches(0.3),
             "其他分配", font_size=16, color=GOLD, bold=True)
other_items = [
    "市场推广：10%",
    "社区激励：5%",
]
add_bullet_list(slide, Inches(0.8), Inches(4.9), Inches(5.5), Inches(1.0), other_items, font_size=14)


# Value capture
add_text_box(slide, Inches(6.5), Inches(4.5), Inches(6.0), Inches(0.3),
             "代币价值捕获", font_size=16, color=GOLD, bold=True)
value_items = [
    "交易手续费折扣",
    "平台治理投票权",
    "新RWA项目优先认购权",
]
add_bullet_list(slide, Inches(6.5), Inches(4.9), Inches(6.0), Inches(1.2), value_items, font_size=14)


# ═══════════════════════════════════════════════
# SLIDE 12: INVESTOR RETURNS
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "投资者收益模型", "Investor Returns Model", 12)

# BNB Chain advantage note
add_shape_bg(slide, Inches(0.8), Inches(1.8), Inches(11.7), Inches(0.6), RGBColor(0x1A, 0x2A, 0x3E))
add_text_box(slide, Inches(1.0), Inches(1.85), Inches(11.3), Inches(0.5),
             "基于币安链（BNB Chain）的RWA生态，实现链上收益自动分发与二级市场无缝交易",
             font_size=14, color=LIGHT_GRAY)

# Return sources
sources = [
    ("🏠", "租金分红", "季度发放\n预期年化 5-8%", Inches(0.8)),
    ("📈", "资产增值", "标的升值带来的\n代币价格增长", Inches(4.8)),
    ("💹", "二级交易", "流动性溢价\n折价套利机会", Inches(8.8)),
]
for icon, title, desc, left in sources:
    add_card(slide, left, Inches(1.8), Inches(3.6), Inches(1.8), f"{icon} {title}", desc)

# Exit mechanism
add_text_box(slide, Inches(0.8), Inches(4.0), Inches(11.7), Inches(0.3),
             "退出机制", font_size=20, color=GOLD, bold=True)

exits = [
    "🔄 二级市场直接卖出 — T+0灵活退出",
    "🛒 定期回购 — 资产方承诺回购机制（具体条件以协议为准）",
]
add_bullet_list(slide, Inches(0.8), Inches(4.4), Inches(11.7), Inches(1.0), exits, font_size=16)

# Risk disclaimer
add_shape_bg(slide, Inches(0.8), Inches(5.5), Inches(11.7), Inches(1.0), RGBColor(0x30, 0x20, 0x20))
add_text_box(slide, Inches(1.0), Inches(5.6), Inches(11.3), Inches(0.3),
             "⚠️ 风险提示", font_size=16, color=ACCENT_RED, bold=True)
add_text_box(slide, Inches(1.0), Inches(5.95), Inches(11.3), Inches(0.4),
             "市场波动风险 | 汇率风险 | 资产流动性风险 | 政策监管风险 | 投资有风险，决策需谨慎",
             font_size=13, color=LIGHT_GRAY)


# ═══════════════════════════════════════════════
# SLIDE 13: RISK CONTROL
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "风控与资产安全", "Risk Control & Asset Security", 13)

# Risk layers
layers = [
    ("第一层", "底层资产尽调", "第三方专业评估机构\n实地尽调 + 独立估值", ACCENT_GREEN),
    ("第二层", "法律合规审查", "知名律所出具\n法律意见书", ACCENT_BLUE),
    ("第三层", "智能合约审计", "安全公司审计\n代码漏洞检测", GOLD),
    ("第四层", "资产托管", "独立托管机构\n资产隔离保管", ACCENT_RED),
]
for i, (layer, title, desc, color) in enumerate(layers):
    left = Inches(0.8 + i * 3.1)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.8), Inches(2.8), Inches(2.3))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0x1A, 0x2A, 0x3E)
    card.line.color.rgb = color
    card.line.width = Pt(2)
    add_text_box(slide, left, Inches(1.9), Inches(2.8), Inches(0.4),
                 layer, font_size=14, color=color, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, left, Inches(2.3), Inches(2.8), Inches(0.4),
                 title, font_size=16, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, left + Inches(0.1), Inches(2.8), Inches(2.6), Inches(1.0),
                 desc, font_size=13, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

# Transparency
add_card(slide, Inches(0.8), Inches(4.5), Inches(5.5), Inches(1.3),
         "资产透明",
         "🔍 链上资产数据实时可查\n📋 定期第三方审计报告公开\n📊 所有交易链上可追溯")

# Insurance
add_card(slide, Inches(6.8), Inches(4.5), Inches(5.5), Inches(1.3),
         "保险机制",
         "🛡️ 底层资产投保\n🔒 防范意外风险\n✅ 全方位安全保障")


# ═══════════════════════════════════════════════
# SLIDE 14: ROADMAP
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "发展路线图", "Roadmap", 14)

phases = [
    ("Phase 1", "Q2-Q3 2026", "平台搭建 + 首个标的尽调",
     [
         "完成平台技术开发",
         "芭提雅标的完成法律确权",
         "获取合规资质",
     ]),
    ("Phase 2", "Q4 2026", "首个RWA发行 + 上线交易",
     [
         "芭提雅RWA正式发行",
         "开放二级市场交易",
         "启动全球社区建设",
     ]),
    ("Phase 3", "2027", "多资产上线 + 国际化",
     [
         "拓展至曼谷、普吉岛",
         "引入商业地产等新RWA类型",
         "申请更多地区合规牌照",
     ]),
    ("Phase 4", "2028+", "生态扩张",
     [
         "东南亚最大RWA发行交易平台",
         "建立RWA行业标准",
     ]),
]
for i, (phase, time, title, items) in enumerate(phases):
    left = Inches(0.5 + i * 3.2)
    # Phase header
    header = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.8), Inches(2.9), Inches(0.9))
    header.fill.solid()
    header.fill.fore_color.rgb = GOLD_DARK if i % 2 == 0 else RGBColor(0x2A, 0x3A, 0x50)
    header.line.fill.background()
    add_text_box(slide, left, Inches(1.85), Inches(2.9), Inches(0.3),
                 f"{phase} | {time}", font_size=13, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, left, Inches(2.2), Inches(2.9), Inches(0.3),
                 title, font_size=12, color=GOLD_LIGHT, alignment=PP_ALIGN.CENTER)
    # Items
    add_bullet_list(slide, left + Inches(0.2), Inches(2.9), Inches(2.6), Inches(2.5),
                    items, font_size=12, color=LIGHT_GRAY, bullet_char="•")
    # Vertical connecting line
    if i < len(phases) - 1:
        add_text_box(slide, left + Inches(3.0), Inches(2.0), Inches(0.3), Inches(0.5),
                     "→", font_size=20, color=GOLD, alignment=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════
# SLIDE 15: CORE TEAM
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK2)
add_header_bar(slide, "核心团队", "Core Team", 15)

# Team members
team = [
    ("👤", "CEO / 创始人", "待定\n行业背景丰富"),
    ("👤", "CTO", "待定\n技术负责人"),
    ("👤", "首席合规官", "待定\n法律/合规背景"),
    ("👤", "首席运营官", "待定\n运营/市场经验"),
]
for i, (icon, title, desc) in enumerate(team):
    left = Inches(0.8 + i * 3.1)
    add_card(slide, left, Inches(1.8), Inches(2.8), Inches(2.2), f"{icon} {title}", desc)

# Advisors
add_card(slide, Inches(0.8), Inches(4.3), Inches(5.8), Inches(1.5),
         "资产顾问",
         "泰国房地产专家团队\n提供本地市场洞察与项目评估支持")

add_card(slide, Inches(7.0), Inches(4.3), Inches(5.5), Inches(1.5),
         "顾问团队（拟邀）",
         "区块链 / 金融 / 法律领域\n知名专家与机构顾问")


# ═══════════════════════════════════════════════
# SLIDE 16: CLOSING
# ═══════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK)

# Decorative top
add_shape_bg(slide, Inches(0), Inches(0), Inches(13.333), Inches(3.0), RGBColor(0x12, 0x1A, 0x30))

# Gold accent
add_shape_bg(slide, Inches(0), Inches(3.0), Inches(13.333), Pt(5), GOLD)

# Summary
add_text_box(slide, Inches(1), Inches(1.0), Inches(11.3), Inches(0.5),
             "金孔雀交易所  |  GOLDEN PEACOCK EXCHANGE",
             font_size=22, color=GOLD, bold=True, alignment=PP_ALIGN.CENTER)

add_text_box(slide, Inches(1), Inches(1.6), Inches(11.3), Inches(0.6),
             "核心信息回顾",
             font_size=20, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)

recap = [
    "✅ 专注东南亚房地产RWA的合规资产发行与交易平台",
    "✅ 首个标的：芭提雅核心区高端房产RWA",
    "✅ 资产方：环球国际资产管理公司",
    "✅ 投资门槛低，流动性强，合规透明",
]
add_bullet_list(slide, Inches(2), Inches(2.1), Inches(9.3), Inches(1.0),
                recap, font_size=14, color=LIGHT_GRAY)

# CTA
add_text_box(slide, Inches(1), Inches(3.8), Inches(11.3), Inches(0.5),
             "合作邀约",
             font_size=24, color=GOLD, bold=True, alignment=PP_ALIGN.CENTER)

add_text_box(slide, Inches(1), Inches(4.4), Inches(11.3), Inches(0.8),
             "欢迎机构投资者 · 资产方 · 合作伙伴 洽谈交流",
             font_size=18, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

# Contact
add_shape_bg(slide, Inches(3), Inches(5.3), Inches(7.333), Inches(1.0), RGBColor(0x22, 0x2E, 0x44))
add_text_box(slide, Inches(3.5), Inches(5.4), Inches(6.333), Inches(0.3),
             "联系方式（待定）", font_size=15, color=GOLD, bold=True, alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(3.5), Inches(5.75), Inches(6.333), Inches(0.4),
             "官网 / 邮箱 / Telegram / 微信商务",
             font_size=14, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

# Tagline
add_text_box(slide, Inches(1), Inches(6.5), Inches(11.3), Inches(0.6),
             "「让世界资产自由流动」",
             font_size=24, color=GOLD_LIGHT, bold=True, alignment=PP_ALIGN.CENTER)

# Bottom bar
add_shape_bg(slide, Inches(0), Inches(7.1), Inches(13.333), Pt(4), GOLD)

add_section_number(slide, 16)


# ── SAVE ──
output_dir = "/Users/mac/.openclaw/workspace/projects/金孔雀交易所PPT"
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, "金孔雀交易所_Golden_Peacock_Exchange.pptx")
prs.save(output_path)
print(f"✅ PPT saved: {output_path}")
print(f"📊 Total slides: {len(prs.slides)}")
