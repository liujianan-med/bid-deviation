# -*- coding: utf-8 -*-
"""刘佳男 · 单页简历（复刻 2024 版斜切角标签条排版）"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = (0x2A, 0x40, 0x58)          # 2024 版主色 #2A4058
NAVY_HEX = "2A4058"
NAVY_LIGHT = (0xB9, 0xC4, 0xD0)
GREY_BLUE_HEX = "758495"
ACCENT = (0x1F, 0x6F, 0xB2)
DARK = (0x2B, 0x2B, 0x2B)
GREY = (0x59, 0x57, 0x57)
LIGHT = (0x9A, 0xA3, 0xAE)

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin = Cm(1.0)
sec.bottom_margin = Cm(1.0)
sec.left_margin = Cm(1.5)
sec.right_margin = Cm(1.5)

st = doc.styles["Normal"]
st.font.name = "Microsoft YaHei"
st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
st.paragraph_format.line_spacing = 1.08
st.paragraph_format.space_after = Pt(0)


def R(p, text, size=10, bold=False, color=None, font="微软雅黑", spacing=None):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = font
    r.element.rPr.rFonts.set(qn("w:eastAsia"), font)
    if color:
        r.font.color.rgb = RGBColor(*color)
    if spacing:
        rpr = r._element.get_or_add_rPr()
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(int(spacing * 20)))
        rpr.append(sp)
    return r


def P(container=None, align=None, before=0, after=0, indent=0, line=None):
    p = container.add_paragraph() if container is not None else doc.add_paragraph()
    if align:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if indent:
        pf.left_indent = Cm(indent)
    if line:
        pf.line_spacing = line
    return p


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), hexcolor)
    tcPr.append(sh)


def cell_valign(cell, val="top"):
    tcPr = cell._tc.get_or_add_tcPr()
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), val)
    tcPr.append(va)


def no_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "none")
        e.set(qn("w:sz"), "0")
        borders.append(e)
    tblPr.append(borders)


def cell_margins(table, top=0, bottom=0, left=0, right=0):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement("w:" + k)
        e.set(qn("w:w"), str(int(v * 567)))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tblPr.append(mar)


def fixed_layout(table):
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)


def bar(text):
    """斜切角标签条：深藏青底 + 白字 + 灰色斜角块"""
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    no_borders(t)
    fixed_layout(t)
    cell_margins(t, top=0.09, bottom=0.09, left=0.34, right=0)
    a, b = t.cell(0, 0), t.cell(0, 1)
    a.width = Cm(3.3)
    b.width = Cm(14.7)
    shade(a, NAVY_HEX)
    shade(b, GREY_BLUE_HEX)
    cell_valign(a, "center")
    cell_valign(b, "center")
    pa = a.paragraphs[0]
    pa.paragraph_format.space_after = Pt(0)
    R(pa, text, size=10, bold=True, color=(0xFF, 0xFF, 0xFF), spacing=1.5)
    pb = b.paragraphs[0]
    pb.paragraph_format.space_after = Pt(0)
    R(pb, "", size=6)
    return t


def item(parts, after=3, indent=0.1):
    p = P(after=after, indent=indent, line=1.22)
    R(p, "◆ ", size=7.5, color=(0x1F, 0x6F, 0xB2))
    for text, bold, color in parts:
        R(p, text, size=10, bold=bold, color=color or DARK)
    return p


# ==================== 顶部页眉带 ====================
top = doc.add_table(rows=1, cols=1)
no_borders(top)
fixed_layout(top)
cell_margins(top, top=0.1, bottom=0.1, left=0, right=0)
tc = top.cell(0, 0)
shade(tc, NAVY_HEX)
tp = tc.paragraphs[0]
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
tp.paragraph_format.space_after = Pt(0)
R(tp, "医 疗 器 械 行 业 专 家　·　A I　 应 用 实 践 者", size=7.5, color=(0xD5, 0xDE, 0xE7), spacing=2.0)

P(after=0, before=6)

# ==================== 姓名 + 照片 ====================
ht = doc.add_table(rows=1, cols=2)
no_borders(ht)
fixed_layout(ht)
cell_margins(ht, 0, 0, 0, 0)
hc_l, hc_r = ht.cell(0, 0), ht.cell(0, 1)
hc_l.width = Cm(14.2)
hc_r.width = Cm(3.8)
cell_valign(hc_l, "top")
cell_valign(hc_r, "top")

p = hc_l.paragraphs[0]
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(3)
R(p, "刘　佳　男", size=22, bold=True, color=NAVY, spacing=2)

p = hc_l.add_paragraph()
p.paragraph_format.space_after = Pt(4)
R(p, "医疗美容设备行业专家　|　AI 应用实践者", size=10.5, color=ACCENT, spacing=1.0)

# 信息两列
info_t = hc_l.add_table(rows=2, cols=4)
no_borders(info_t)
fixed_layout(info_t)
cell_margins(info_t, 0, 0, 0, 0.1)
cells = [("电　话", "15910885356"), ("年　龄", "29 岁"), ("邮　箱", "William_liujianan@163.com"), ("常　驻", "成都（覆盖西南/西北）")]
for i, (k, v) in enumerate(cells):
    r_, c_ = i // 2, (i % 2) * 2
    kc, vc = info_t.cell(r_, c_), info_t.cell(r_, c_ + 1)
    kc.width = Cm(1.3)
    vc.width = Cm(5.5)
    kp = kc.paragraphs[0]
    kp.paragraph_format.space_after = Pt(1)
    R(kp, k, size=8, color=GREY, spacing=0.5)
    vp = vc.paragraphs[0]
    vp.paragraph_format.space_after = Pt(1)
    R(vp, v, size=9, color=DARK, font="Calibri" if i in (0, 2) else "微软雅黑")

rp = hc_r.paragraphs[0]
rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
rp.paragraph_format.space_after = Pt(0)
rp.add_run().add_picture(r"C:\Users\Administrator\code\bid-deviation\data\photo_hd.png", width=Cm(3.0))

P(after=0, before=10)

# ==================== 个人定位 ====================
bar("个人概述")
p = P(before=3, after=0, line=1.35)
R(p, "6 年医疗器械行业经验", size=9.5, bold=True, color=NAVY)
R(p, "，路径贯穿临床跟台 → 设备销售 → 招投标全链路：从强生心血管介入导管室技术支持起步，深耕医美设备行业，熟悉皮秒激光、强脉冲光等产品的技术参数、竞品格局、公立医院采购流程与经销商体系。当前聚焦用大语言模型重构医疗投标与商务流程，", size=9.5, color=DARK)
R(p, "已独立建成并在真实政采项目投产", size=9.5, bold=True, color=NAVY)
R(p, "。", size=9.5, color=DARK)

# ==================== 工作经历 ====================
P(after=0, before=10)
bar("工作经历")

p = P(before=3, after=1)
R(p, "南京伟思医疗科技股份有限公司　|　医美事业部 · 区域销售经理", size=9.5, bold=True, color=NAVY)
p = P(after=2)
R(p, "2024 – 至今　|　成都", size=8.5, color=GREY, font="Calibri")
item([("负责西南、西北区域皮秒激光、强脉冲光等医美设备的销售与经销商渠道管理，覆盖四川、新疆等地公立医院与民营机构", False, None)])
item([
    ("独立开发技术偏离表智能生成系统", True, NAVY),
    ("：输入招标文件自动完成参数提取、三态判定与投标文档生成，单份标书 ", False, None),
    ("数小时 → 分钟级", True, ACCENT),
    ("，已在真实政府采购项目中端到端使用", False, None),
])

p = P(before=5, after=1)
R(p, "深圳半岛医疗科技有限责任公司　|　销售经理", size=9.5, bold=True, color=NAVY)
p = P(after=2)
R(p, "2022.08 – 2023.09　|　成都", size=8.5, color=GREY, font="Calibri")
item([
    ("负责皮肤科治疗设备在华西本部及分院、成都市一院/六院及眉山、雅安、乐山等川内地级市三甲三乙医院的招标、销售与售后", False, None),
])
item([
    ("年度业绩任务 ", False, None),
    ("300 万、完成率约 85%", True, ACCENT),
    ("；成都公立医院板块份额由 ", False, None),
    ("15% 提升至约 25%", True, ACCENT),
    ("；雅安市第二人民医院科室采购占比实现 ", False, None),
    ("0 → 50%", True, ACCENT),
])
item([
    ("推动成都市金牛区妇幼保健院成立以公司全系列设备为治疗闭环的研究治疗中心", False, None),
])

p = P(before=5, after=1)
R(p, "强生（中国）医疗器材有限责任公司　|　产品技术专员", size=9.5, bold=True, color=NAVY)
p = P(after=2)
R(p, "2020.07 – 2022.06　|　北京 / 福建", size=8.5, color=GREY, font="Calibri")
item([
    ("于北京协和等医院导管室内配合术者完成复杂心脏介入手术，独立跟台 ", False, None),
    ("101 台", True, ACCENT),
    ("；负责导管室技术支持、新产品推广与配送协调", False, None),
])
item([
    ("推动公司最新 8F 超声导管引进，完成", False, None),
    ("福建省首例", True, ACCENT),
    (" 8F 超声导管联合射频消融手术；全年组织行业中小型学术会议 ", False, None),
    ("20 场", True, ACCENT),
])

# ==================== AI 项目经验 ====================
P(after=0, before=10)
bar("AI 项目经验")

p = P(before=3, after=1)
R(p, "技术偏离表智能生成系统", size=9.5, bold=True, color=NAVY)
R(p, "　独立主导 · 2026", size=8.5, color=GREY, font="Calibri")
item([("针对医疗设备投标中偏离表编制耗时、易漏项、易废标的行业痛点", False, None)])
item([
    ("大语言模型负责招标文件语义解析与参数抽取，规则引擎负责数值三态判定（响应 / 正偏离 / 负偏离），自动输出带证据页码出处的可投标 Word 文档", False, None),
])
item([
    ("真实地市级医院政采项目实测：", False, None),
    ("18 条技术参数 33 秒", True, ACCENT),
    ("完成判定与成表；投标资料包整理由 ", False, None),
    ("3 天压缩至半天", True, ACCENT),
])

# ==================== 核心专长 ====================
P(after=0, before=10)
bar("核心专长")
for k, v in [
    ("医美设备行业", "皮秒激光、强脉冲光（IPL）、射频、磁刺激等产品的技术、临床应用与竞品对比"),
    ("器械采购链路", "政府采购招投标全流程、技术参数响应与偏离判定、厂家授权与经销商生态"),
    ("临床端认知", "心内科导管室 100+ 台跟台经验，理解科室决策链与学术会议推广模式"),
    ("AI 落地实践", "大语言模型文档自动化：招标解析、参数比对、投标文档生成"),
]:
    item([(k + "　", True, NAVY), (v, False, None)], after=1.5)

# ==================== 教育背景 ====================
P(after=0, before=10)
bar("教育背景")
p = P(before=3, after=1)
R(p, "西南石油大学（双一流）", size=9.5, bold=True, color=NAVY)
R(p, "　材料工程 · 硕士　|　2017.07 – 2020.06", size=9, color=GREY, font="Calibri")
p = P(before=2.5, after=1)
R(p, "西南石油大学（双一流）", size=9.5, bold=True, color=NAVY)
R(p, "　高分子材料与工程 · 学士　|　2013.07 – 2017.06", size=9, color=GREY, font="Calibri")
p = P(before=2.5, after=0)
R(p, "实习：强生（中国）Ethicon 缝线部（成都）　|　2018.06 – 2018.09", size=8.5, color=LIGHT, font="Calibri")

# ==================== 底部条 ====================
P(after=0, before=8)
bot = doc.add_table(rows=1, cols=1)
no_borders(bot)
fixed_layout(bot)
cell_margins(bot, top=0.1, bottom=0.1, left=0, right=0)
bc = bot.cell(0, 0)
shade(bc, NAVY_HEX)
bp = bc.paragraphs[0]
bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
bp.paragraph_format.space_after = Pt(0)
R(bp, "懂医院采购决策链　×　能动手把 AI 做成业务工具　——　医疗行业中稀缺的复合型落地者", size=8, color=(0xD5, 0xDE, 0xE7), spacing=1.2)

out = r"C:\Users\Administrator\Desktop\个人简历-刘佳男-医疗AI专家版-单页.docx"
doc.save(out)
print("saved:", out)
