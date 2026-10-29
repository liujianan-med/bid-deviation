# -*- coding: utf-8 -*-
"""刘佳男 · 医疗 AI / FDE 方向 单页简历 + 面试问答稿
沿用 2024 版斜切角标签条排版（#2A4058 / #758495）
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = (0x2A, 0x40, 0x58)
NAVY_HEX = "2A4058"
GREY_BLUE_HEX = "758495"
ACCENT = (0x1F, 0x6F, 0xB2)
DARK = (0x2B, 0x2B, 0x2B)
GREY = (0x59, 0x57, 0x57)
LIGHT = (0x9A, 0xA3, 0xAE)
RED = (0xA8, 0x3A, 0x33)


# ---------------------------------------------------------------- 通用构件
def new_doc():
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
    return doc


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


def P(doc, align=None, before=0, after=0, indent=0, line=None):
    p = doc.add_paragraph()
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


def bar(doc, text, w_left=3.5, w_right=14.5, right=""):
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    no_borders(t)
    fixed_layout(t)
    cell_margins(t, top=0.09, bottom=0.09, left=0.34, right=0)
    a, b = t.cell(0, 0), t.cell(0, 1)
    a.width = Cm(w_left)
    b.width = Cm(w_right)
    shade(a, NAVY_HEX)
    shade(b, GREY_BLUE_HEX)
    cell_valign(a, "center")
    cell_valign(b, "center")
    pa = a.paragraphs[0]
    pa.paragraph_format.space_after = Pt(0)
    R(pa, text, size=10, bold=True, color=(0xFF, 0xFF, 0xFF), spacing=1.5)
    pb = b.paragraphs[0]
    pb.paragraph_format.space_after = Pt(0)
    if right:
        R(pb, "　" + right, size=8.5, color=(0xFF, 0xFF, 0xFF))
    else:
        R(pb, "", size=6)
    return t


def item(doc, parts, after=3, indent=0.1, size=9.5, line=1.16):
    p = P(doc, after=after, indent=indent, line=line)
    R(p, "◆ ", size=7.5, color=ACCENT)
    for text, bold, color in parts:
        R(p, text, size=size, bold=bold, color=color or DARK)
    return p


def header_band(doc, text):
    top = doc.add_table(rows=1, cols=1)
    no_borders(top)
    fixed_layout(top)
    cell_margins(top, top=0.1, bottom=0.1, left=0, right=0)
    tc = top.cell(0, 0)
    shade(tc, NAVY_HEX)
    tp = tc.paragraphs[0]
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_after = Pt(0)
    R(tp, text, size=7.5, color=(0xD5, 0xDE, 0xE7), spacing=2.0)


def bottom_band(doc, text):
    bot = doc.add_table(rows=1, cols=1)
    no_borders(bot)
    fixed_layout(bot)
    cell_margins(bot, top=0.1, bottom=0.1, left=0, right=0)
    bc = bot.cell(0, 0)
    shade(bc, NAVY_HEX)
    bp = bc.paragraphs[0]
    bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    bp.paragraph_format.space_after = Pt(0)
    R(bp, text, size=8, color=(0xD5, 0xDE, 0xE7), spacing=1.2)


PHOTO = r"C:\Users\Administrator\code\bid-deviation\data\photo_hd.png"


# ================================================================ 简历
def build_resume(out):
    doc = new_doc()

    header_band(doc, "医 疗 器 械 行 业 经 验　·　A I　 产 品 落 地 实 践")
    P(doc, after=0, before=6)

    # ---- 姓名 + 照片
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
    R(p, "医疗 AI 产品 / 前沿部署（FDE）方向", size=10.5, color=ACCENT, spacing=1.0)

    info_t = hc_l.add_table(rows=2, cols=4)
    no_borders(info_t)
    fixed_layout(info_t)
    cell_margins(info_t, 0, 0, 0, 0.1)
    cells = [("电　话", "15910885356"), ("年　龄", "29 岁"),
             ("邮　箱", "William_liujianan@163.com"), ("常　驻", "成都（可覆盖西南/西北）")]
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
    rp.add_run().add_picture(PHOTO, width=Cm(2.9))

    P(doc, after=0, before=9)

    # ---- 个人概述
    bar(doc, "个人概述")
    p = P(doc, before=3, after=0, line=1.32)
    R(p, "6 年医疗器械行业经验", size=9.5, bold=True, color=NAVY)
    R(p, "，路径贯穿临床跟台 → 设备销售 → 招投标：强生导管室技术支持起步，后任半岛医疗、伟思医疗区域销售，", size=9.4, color=DARK)
    R(p, "熟悉公立医院采购决策链、皮肤科与整形科工作流、民营医美机构运营", size=9.4, bold=True, color=NAVY)
    R(p, "。2026 年独立完成「大模型 + 规则引擎」投标文档系统并在真实政采项目投产；", size=9.4, color=DARK)
    R(p, "优势是把医生与科室的模糊抱怨拆成可落地的 AI 需求，并自己动手做出可演示的原型。", size=9.4, bold=True, color=ACCENT)

    # ---- 工作经历
    P(doc, after=0, before=7)
    bar(doc, "工作经历")

    p = P(doc, before=3, after=1)
    R(p, "南京伟思医疗科技股份有限公司　|　医美事业部 · 区域销售经理", size=9.5, bold=True, color=NAVY)
    p = P(doc, after=2)
    R(p, "2024 – 至今　|　成都", size=8.5, color=GREY, font="Calibri")
    item(doc, [("负责西南、西北区域皮秒激光、强脉冲光设备销售与经销商渠道管理，长期在临床科室与设备科之间做需求翻译，", False, None),
               ("熟悉从科室提需求到招标落地的完整链路", True, NAVY),
               ("。", False, None)], after=2)
    item(doc, [("基于一线痛点独立开发", False, None),
               ("技术偏离表智能生成系统", True, NAVY),
               ("，把「一份标书 3 天」压到分钟级，已在真实政采项目端到端使用。", False, None)], after=2)

    p = P(doc, before=4, after=1)
    R(p, "深圳半岛医疗科技有限责任公司　|　销售经理", size=9.5, bold=True, color=NAVY)
    p = P(doc, after=2)
    R(p, "2022.08 – 2023.09　|　成都", size=8.5, color=GREY, font="Calibri")
    item(doc, [("负责华西本部及分院、成都市一院/六院及川内地级市三甲医院的招标、销售与售后；年任务 300 万、完成率 ", False, None),
               ("85%", True, ACCENT),
               ("，成都公立板块份额 ", False, None),
               ("15% → 25%", True, ACCENT),
               ("，雅安二院科室采购占比 ", False, None),
               ("0 → 50%", True, ACCENT),
               ("。", False, None)], after=2)

    p = P(doc, before=4, after=1)
    R(p, "强生（中国）医疗器材有限责任公司　|　产品技术专员", size=9.5, bold=True, color=NAVY)
    p = P(doc, after=2)
    R(p, "2020.07 – 2022.06　|　北京 / 福建", size=8.5, color=GREY, font="Calibri")
    item(doc, [("北京协和等医院导管室配合完成复杂心脏介入手术，独立跟台 ", False, None),
               ("101 台", True, ACCENT),
               ("；推动 8F 超声导管完成", False, None),
               ("福建省首例", True, ACCENT),
               ("联合射频消融，全年组织学术会议 ", False, None),
               ("20 场", True, ACCENT),
               ("。", False, None)], after=2)

    # ---- AI 项目
    P(doc, after=0, before=7)
    bar(doc, "AI 项目经验")

    p = P(doc, before=3, after=1)
    R(p, "医疗投标文档智能生成系统（偏离表 / 资料包）", size=9.5, bold=True, color=NAVY)
    R(p, "　独立主导 · 2026", size=8.5, color=GREY, font="Calibri")
    item(doc, [("问题：", True, NAVY),
               ("招标文件 100+ 页，人工逐条比对参数，一份标书 2–3 天；漏项或含糊表述即导致投标无效。", False, None)], after=2)
    item(doc, [("方案：", True, NAVY),
               ("大模型做招标文件语义解析与参数抽取（含表格、图片型参数），规则引擎做数值区间三态判定，"
                "输出带证据页码的可投标 Word；另建资料包与「缺口清单」模块，按厂家/经销商分列责任方。", False, None)], after=2)
    item(doc, [("工程：", True, NAVY),
               ("Python + 大模型 API + python-docx/openpyxl；产品库按品类建模，新增设备只加数据不改代码；"
                "缺数据不兜底、显式留白。", False, None)], after=2)
    item(doc, [("评测：", True, NAVY),
               ("自建 ", False, None),
               ("49 条金标准测试集", True, ACCENT),
               ("（皮秒 31 / 光子 18），对可量化条款做判定回归，", False, None),
               ("可测量项准确率 100%", True, ACCENT),
               ("；不可量化条款单独标注、交人工确认。", False, None)], after=2)
    item(doc, [("结果：", True, NAVY),
               ("单份偏离表 3 天 → ", False, None),
               ("33 秒", True, ACCENT),
               ("；投标资料包 3 天 → 半天；已在真实政府采购项目中端到端跑通。", False, None)], after=2)

    # ---- 能力匹配
    P(doc, after=0, before=7)
    bar(doc, "能力匹配", right="对标医疗 AI 产品 / FDE 岗位要求")
    for k, v in [
        ("场景挖掘", "覆盖华西本部及分院、成都市一/六院、川内地级市三甲与新疆兵团系统医院，理得清临床科室、设备科、信息科三方诉求差异，能把「这个流程太慢」拆成具体待解决问题。"),
        ("原型交付", "用 AI Coding 工具独立完成解析 / 抽取 / 判定 / 成表 / 打包 5 个模块，想法到可演示原型 1–2 天，不等研发排期。"),
        ("效果评测", "习惯先定金标准再改模型：分得清哪些条款能量化、哪些必须人工兜底，用回归集而不是「感觉好多了」证明改进。"),
        ("合规意识", "清楚医疗数据红线（患者信息不出院、不接 HIS）、器械软件注册边界、医美宣传的广告法限制——多数产品经理的盲区。"),
    ]:
        item(doc, [(k + "　", True, NAVY), (v, False, None)], after=2, size=9.2, line=1.2)

    # ---- 教育
    P(doc, after=0, before=7)
    bar(doc, "教育背景")
    p = P(doc, before=3, after=1)
    R(p, "西南石油大学（双一流）", size=9.5, bold=True, color=NAVY)
    R(p, "　材料工程 · 硕士　|　2017.07 – 2020.06", size=9, color=GREY, font="Calibri")
    p = P(doc, before=2, after=0)
    R(p, "西南石油大学（双一流）", size=9.5, bold=True, color=NAVY)
    R(p, "　高分子材料与工程 · 学士　|　2013.07 – 2017.06　｜　强生 Ethicon 实习 2018", size=9, color=GREY, font="Calibri")

    P(doc, after=0, before=6)
    bottom_band(doc, "懂科室场景　×　能动手交付　×　会用评测证明效果　——　医疗 AI 落地最缺的中间层")

    doc.save(out)
    print("saved:", out)


# ================================================================ 面试问答稿
def build_qa(out):
    doc = new_doc()
    doc.sections[0].top_margin = Cm(1.4)
    doc.sections[0].bottom_margin = Cm(1.4)

    header_band(doc, "医 疗 A I　·　产 品 / F D E　面 试 应 答 稿　·　内 部 使 用")
    P(doc, after=0, before=6)

    p = P(doc, after=2, before=2)
    R(p, "刘佳男　|　医疗 AI 产品 / 前沿部署（FDE）方向", size=15, bold=True, color=NAVY, spacing=1.2)
    p = P(doc, after=8, line=1.35)
    R(p, "用法：不要背。每题先记住「一句话结论」，再用自己的话补细节。"
         "所有数字都可查证，被追问时能讲出来源——这是你相对于应届生和纯技术候选人的最大优势。",
      size=9, color=GREY)

    bar(doc, "一、开场")
    p = P(doc, before=3, after=1)
    R(p, "1. 请做一下自我介绍（60 秒版）", size=10, bold=True, color=NAVY)
    for line in [
        "「我是刘佳男，6 年医疗行业经验，走了三段路：强生做心血管介入的术中技术支持，跟台 101 台；"
        "后来到半岛医疗和现在的伟思医疗做医美设备销售，负责川渝和新疆的公立医院与民营机构；"
        "最近一年，我把一线最痛的投标工作做成了一个 AI 系统——招标文件进去，偏离表和投标资料包出来，"
        "一份标书从 3 天压到分钟级，已经在真实的政府采购项目里跑过。」",
        "「所以我给自己的定位不是转行，是把行业理解接到 AI 上：我知道科室里哪句话是真需求，也知道 AI 哪些事干不了。」",
    ]:
        item(doc, [(line, False, None)], after=2, size=9.3, line=1.32)

    bar(doc, "二、核心问题")
    qa = [
        ("2. 你为什么从销售转产品 / FDE？是不是销售干不下去了？",
         "结论：不是逃离，是往上走一层。",
         "「销售的天花板是区域和产品线，我能卖好一台设备，但改变不了整个流程。"
         "做偏离表系统时我第一次发现：把一线痛点变成产品，影响的是所有用这套流程的人。"
         "另外我很清楚自己的短板——不会写生产级代码。所以我选 FDE 而不是纯产品经理："
         "FDE 要的是懂业务、能进现场、能自己搭出原型，这三条正好是我的长板。」"),
        ("3. 你的技术到什么程度？会不会只是「会聊 AI」？",
         "结论：能独立做出可用系统，但不能替代工程师。",
         "「诚实说：我能独立完成需求拆解、数据清洗、Prompt 设计、规则引擎、文档生成和评测，"
         "整个系统是我一个人用 Python + 大模型 API 写完的，2000 多行，跑在真实项目上。"
         "但我不是软件工程师——我不会做高并发、不会做权限体系、不会做前端工程化。"
         "FDE 的定位恰好是『最后一公里』：能用原型证明方案可行，再交给工程团队产品化。」"),
        ("4. 讲一个你从 0 到 1 的项目（准备被追问 5 分钟）",
         "结论：用 STAR，重点讲「你怎么发现问题」和「你怎么证明有效」。",
         "S：每份标书要人工比对上百条参数，3 天一份，漏一条就废标，团队每年在这上面耗掉大量时间。"
         "T：让系统自动完成参数抽取、逐条比对、生成可投标的 Word。"
         "A：先定『什么算做对』——可量化条款做三态判定（响应/正偏离/负偏离），不可量化的显式留白交人工；"
         "再做产品库，把参数按品类建模，新增设备只加数据不改代码；最后建 49 条金标准做回归。"
         "R：一份偏离表 33 秒，资料包 3 天变半天，缺陷是「宁可留白也不编」，用这个换来了投标人的信任。"),
        ("5. 你怎么做需求调研？给你一个陌生科室，你第一周干什么？",
         "结论：先蹲现场，再问三个问题。",
         "「第一周不访谈，先看：看他们一天里重复最多的动作是什么，卡在哪个系统之间。"
         "然后问三个问题：①这件事你现在怎么做、一次多久？②做错了会怎样（谁承担后果）？"
         "③如果有个东西能帮你省一半时间，你愿意换掉现在的方式吗？"
         "第三个问题决定这是不是真需求——很多人会说『挺好』，但不愿改变习惯，那就是伪需求。」"),
        ("6. 你怎么判断一个 AI 功能做得好不好？",
         "结论：先定金标准，再谈指标。",
         "「医疗场景我只看两件事：可量化部分的准确率，和不可量化部分有没有被诚实地标出来。"
         "做偏离表时我把条款分成能测得和不能测的，能测的做回归集，不能测的一律标记人工确认。"
         "产品上也一样：先定 30–50 条真实样本作为验收基线，每次改动跑一遍，不让『感觉变好了』进版本。」"),
        ("7. 医疗 AI 的合规红线，你知道哪些？",
         "结论：三条——数据不出院、诊断不越界、宣传不承诺。",
         "「①患者数据：不能出医院内网，不能碰 HIS 原始数据，脱敏也不行，这是《数据安全法》《个人信息保护法》的硬约束；"
         "②辅助诊断类软件属于医疗器械，要有注册证，没有证就不能说『诊断』，只能说『辅助』『提示』；"
         "③医美宣传受《医疗广告管理办法》管，不能用前后对比、不能承诺效果，AI 生成内容同样受管。"
         "这也是我为什么第一个产品选投标文档——它不碰患者数据，合规风险最低。」"),
        ("8. 你不会写代码，怎么和研发 / 算法团队协作？",
         "结论：我用他们的语言提需求，用原型代替嘴。",
         "「我不跟他们说『做个智能问答』，我说：输入是什么、输出长什么样、错了怎么兜底、验收标准是哪 30 条用例。"
         "先把原型搭出来给他们看，比文字需求省一半沟通。"
         "另外我知道边界：什么时候该用 RAG、什么时候该用规则、什么时候干脆别用模型——"
         "这个判断能力比会写代码更重要，因为它决定项目会不会做废。」"),
        ("9. 你的期望薪资？",
         "结论：给区间，讲清楚为什么。",
         "「我现在是销售岗，收入结构里提成占比不小。转过来我更看重的是能不能真正做产品。"
         "薪资上我给一个区间：18–22K。理由是我有 6 年行业积累和可验证的项目，"
         "不需要 6 个月熟悉医院场景，入职第一个月就能进科室访谈。」"
         "（注：低于 15K 的岗位不建议谈，等于用 6 年经验换实习价。）"),
        ("10. 压力题：你没做过产品经理 / 你 29 岁了 / 你学历不对口",
         "结论：不辩解，把劣势换成差异。",
         "「确实没做过产品经理，所以我没带着一套模板来。但我带了三样别人没有的："
         "一是我在科室里待过 6 年，知道医生说的是什么意思；二是我自己从 0 做成了一个上线系统；"
         "三是我懂这个行业的钱怎么花、谁签字、什么流程会卡住。"
         "产品方法论我可以三个月补上，行业理解补不了。」"),
    ]
    for q, concl, ans in qa:
        P(doc, after=0, before=8)
        p = P(doc, after=1)
        R(p, q, size=10, bold=True, color=NAVY)
        p = P(doc, after=1, indent=0.1)
        R(p, concl, size=9, bold=True, color=RED)
        item(doc, [(ans, False, None)], after=2, size=9.3, line=1.32)

    bar(doc, "三、反问环节")
    for t in [
        "「这个产品现在的真实使用数据是多少？谁在每天用？」——听得出是 demo 还是产品。",
        "「产品团队几个人？上一个版本是谁定的优先级？」——判断你是不是去打杂。",
        "「FDE / 产品经理在这里的边界是什么？出了问题谁负责？」——试探岗位成熟度。",
        "「医院侧的对接人是谁？是信息科还是临床科室？」——判断他们懂不懂落地。",
        "「如果我入职，第一个月你希望我交付什么？」——这个问题会让面试官记住你。",
    ]:
        item(doc, [(t, False, None)], after=2, size=9.3)

    bar(doc, "四、面试前必查")
    for t in [
        "① 他们家的产品有没有医疗器械注册证（三类/二类）——决定合规天花板。",
        "② 最近一轮融资时间和金额——决定还能烧多久。",
        "③ 招股书 / 官网里的客户名单——面试时提一句「你们在 XX 医院的场景」，杀伤力极大。",
    ]:
        item(doc, [(t, False, None)], after=2, size=9.3)

    P(doc, after=0, before=8)
    bottom_band(doc, "所有数字都可查证　·　不确定的事就说不知道　——　医疗 AI 面试最忌编")

    doc.save(out)
    print("saved:", out)


if __name__ == "__main__":
    build_resume(r"C:\Users\Administrator\Desktop\个人简历-刘佳男-医疗AI-FDE方向.docx")
    build_qa(r"C:\Users\Administrator\Desktop\面试问答稿-刘佳男-医疗AI-FDE方向.docx")
