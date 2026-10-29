# -*- coding: utf-8 -*-
"""FDE 转型 16 周执行清单（docx）"""
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
LINE_HEX = "C9D2DC"
BAND_HEX = "EEF2F6"

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin = Cm(1.3)
sec.bottom_margin = Cm(1.3)
sec.left_margin = Cm(1.5)
sec.right_margin = Cm(1.5)

st = doc.styles["Normal"]
st.font.name = "Microsoft YaHei"
st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
st.paragraph_format.line_spacing = 1.12
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


def P(align=None, before=0, after=0, indent=0, line=None):
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


def grid_borders(table, hexcolor=LINE_HEX):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:color"), hexcolor)
        borders.append(e)
    tblPr.append(borders)


def cell_margins(table, top=0.08, bottom=0.08, left=0.14, right=0.14):
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


def bar(text, right=""):
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    no_borders(t)
    fixed_layout(t)
    cell_margins(t, top=0.09, bottom=0.09, left=0.34, right=0)
    a, b = t.cell(0, 0), t.cell(0, 1)
    a.width = Cm(4.0)
    b.width = Cm(14.0)
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


def item(parts, after=2, indent=0.1, size=9.3, line=1.24):
    p = P(after=after, indent=indent, line=line)
    R(p, "◆ ", size=7.5, color=ACCENT)
    for text, bold, color in parts:
        R(p, text, size=size, bold=bold, color=color or DARK)
    return p


def check(text, after=1.5, size=9.3, bold=False):
    p = P(after=after, indent=0.15, line=1.22)
    R(p, "☐  ", size=9.5, color=NAVY)
    R(p, text, size=size, bold=bold, color=DARK)
    return p


def table(headers, rows, widths, font=8.8, header_fill=NAVY_HEX, band=True):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    grid_borders(t)
    fixed_layout(t)
    cell_margins(t)
    hdr = t.rows[0]
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        c.width = Cm(widths[i])
        shade(c, header_fill)
        cell_valign(c, "center")
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        R(p, h, size=8.6, bold=True, color=(0xFF, 0xFF, 0xFF))
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            c = cells[i]
            c.width = Cm(widths[i])
            cell_valign(c, "center")
            if band and ri % 2 == 1:
                shade(c, BAND_HEX)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.12
            bold = v.startswith("**")
            R(p, v.replace("**", ""), size=font, bold=bold, color=NAVY if bold else DARK)
    return t


# ==================== 封面头 ====================
top = doc.add_table(rows=1, cols=1)
no_borders(top)
fixed_layout(top)
cell_margins(top, top=0.1, bottom=0.1, left=0, right=0)
tc = top.cell(0, 0)
shade(tc, NAVY_HEX)
tp = tc.paragraphs[0]
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
tp.paragraph_format.space_after = Pt(0)
R(tp, "F D E　转 型 执 行 清 单　·　从 0 到 第 一 份 O F F E R", size=8, color=(0xD5, 0xDE, 0xE7), spacing=2.2)

P(before=8)
p = P(after=2)
R(p, "16 周落地计划", size=19, bold=True, color=NAVY, spacing=1.5)
p = P(after=8)
R(p, "在职执行 · 每周约 10 小时 · 总预算 1500 元以内　|　刘佳男　2026.10", size=9, color=GREY)

# 总览
bar("一、总览", "先看这一页，决定要不要开始")
P(before=3)
table(
    ["项", "设定", "说明"],
    [
        ["**目标岗位", "FDE / 前沿部署 / 解决方案 / AI 应用实施", "不是产品经理，是进现场能把东西跑起来的人"],
        ["**时间", "16 周到第一份 offer", "每周 10 小时：工作日 1h×5 + 周末 2.5h×2"],
        ["**核心逻辑", "作品集 > 证书 > 学历", "FDE 面试官只认「你能当场演示什么」"],
        ["**成功标准", "≥2 场面试、≥1 个 offer", "拿不到就把岗位降一档，不要硬扛"],
        ["**预算", "API 200-300/月 + 服务器 50/月 + 域名 60/年", "四个月合计 < 1500 元"],
    ],
    [2.6, 6.2, 9.2],
)

# 能力
P(before=10)
bar("二、能力缺口 → 学什么", "只学四样，别贪多")
P(before=3)
table(
    ["能力", "要学到什么程度", "大概用时", "怎么验收"],
    [
        ["**Python", "能读写文件、调 API、处理 Excel；不要求工程能力", "4 周", "不看教程独立写出 100 行脚本"],
        ["**大模型应用", "Prompt 结构化输出、RAG、失败重试、成本估算", "3 周", "能说清「什么时候不该用模型」"],
        ["**低代码平台", "扣子 / Dify 搭工作流，n8n 做自动化", "1 周", "搭出 1 个能跑的工作流"],
        ["**方案表达", "一页纸 PRD、3 分钟 Demo 讲解、需求拆解", "2 周（边做边练）", "能把一个需求讲成「输入-输出-兜底-验收」"],
    ],
    [2.6, 8.0, 2.4, 5.0],
)
P(before=6)
item([("评测能力（金标准、回归、可测量 vs 不可测量）你已经有了，", False, None),
      ("不用学，只要换个说法讲出来", True, NAVY),
      ("——这是 JD 里明写、90% 候选人答不上来的东西。", False, None)])

# 16周
P(before=10)
bar("三、16 周周计划", "每周只做表里这一件事")
P(before=3)
table(
    ["周次", "任务", "当周产出（可给别人看）"],
    [
        ["**W0", "盘点素材：整理偏离表项目的截图/数据/我的角色；注册 GitHub 并上传脱敏代码；确定目标城市与岗位关键词", "一份项目档案 + 一个 GitHub 仓库"],
        ["**W1", "Python 基础：变量/循环/函数/字典/文件读写。目标不是写程序，是「能读懂自己项目里的代码」", "能解释现有脚本每一行在干什么"],
        ["**W2", "调用 API：requests、JSON 处理、把 key 放进环境变量、异常重试", "一个命令行小程序：输入一句话，调用模型返回结构化结果"],
        ["**W3", "数据处理：pandas / openpyxl 读写 Excel、批量处理文件", "批量读取 10 份文档并汇总成一张表"],
        ["**W4", "写一个真实小工具（建议：读招标文件 → 抽参数 → 出 Excel）", "**G1 验收：不看教程，独立完成并跑通"],
        ["**W5", "Prompt 工程：结构化输出（JSON）、few-shot、失败兜底策略", "同一个任务，输出稳定性从「时好时坏」到「可复现」"],
        ["**W6", "RAG：文档切分、向量检索、引用出处、幻觉控制", "一个能回答你产品手册问题的问答原型"],
        ["**W7", "扣子 / Dify 搭工作流 + n8n 自动化（不写代码也能跑）", "1 个工作流：如标讯抓取 → 摘要 → 推送"],
        ["**W8", "把 RAG 原型做成可访问的网页（科室知识问答）", "**G2 验收：手机能打开、答 10 个真实问题、带引用出处"],
        ["**W9", "Demo 2：把偏离表系统做成 Web 演示版（或咨询师陪练）", "第二个可演示案例 + 一页纸 PRD"],
        ["**W10", "作品集网页：3 个案例，每个写 问题 / 方案 / 我的角色 / 结果 / 截图", "一个能发链接的作品集页"],
        ["**W11", "简历 v2（按项目更新）+ 面试问答演练，录音自听 3 遍", "简历终版 + 3 分钟自我介绍不打结"],
        ["**W12", "投递 20 家：5 家冲刺 / 10 家匹配 / 5 家练手；同时找内推", "**G3 验收：20 份投出，至少 3 个内推"],
        ["**W13", "开始面试；每场面完 30 分钟内记录被问到的所有问题", "一份真实问题库（比任何题库都准）"],
        ["**W14", "按问题库补漏：哪个答不上就补哪个，只补高频的", "技术面通过率明显上升"],
        ["**W15", "谈薪与选择：拿到 offer 再谈，先问区间不先报数", "至少 1 个 offer 进入终轮"],
        ["**W16", "决策：看现金流、看团队、看产品有没有真实客户", "**G4 验收：签 offer"],
    ],
    [1.3, 9.6, 7.1],
)

# 作品集
P(before=10)
bar("四、作品集清单", "这是面试的胜负手")
P(before=3)
check("**Demo 1：科室知识问答（必做）** —— 喂皮肤科设备资料/科室 SOP，做成能对话的网页，必须带引用出处。这是最像「医疗 AI」的东西。")
check("**Demo 2：投标偏离表系统（已有，重新包装）** —— 不要讲技术，讲「一份标书 3 天 → 33 秒」和「宁可留白不编」的取舍。")
check("**Demo 3：一个自动化工作流** —— 标讯抓取 → 摘要 → 微信/邮件推送，证明你会用 n8n/扣子这类工具。")
check("**作品集网页** —— 每个案例固定五段：问题、方案、我的角色、结果数字、截图。移动端能看。")
check("**两页纸 PRD** —— 随手写 Demo 1 的需求文档，证明你会产品语言，不只是会敲代码。")

# 求职
P(before=10)
bar("五、投哪里、怎么投")
P(before=3)
p = P(after=2)
R(p, "岗位关键词（在 BOSS / 拉勾 / 脉脉 上搜这些词，不要搜「AI 产品经理」）：", size=9.3, bold=True, color=NAVY)
item([("FDE、前沿部署工程师、解决方案工程师、AI 应用工程师、交付工程师、实施顾问、售前工程师", False, None)])

p = P(before=4, after=2)
R(p, "目标公司分层（先投同行，再投医疗 AI，最后投通用 AI）：", size=9.3, bold=True, color=NAVY)
table(
    ["层级", "公司", "为什么"],
    [
        ["**第一层：同行", "半岛医疗（深圳，FDE 孵化营为 25–27 届校招，你走社招线）、复锐医疗、奇致激光、伟思（内部转岗）", "行业认知 100% 复用，你的 6 年直接算钱"],
        ["**第二层：医疗 AI", "讯飞医疗、东壁科技、深睿、推想、森亿、联影智能、医准", "薪资高，但要能跟信息科对谈技术"],
        ["**第三层：医院信息化", "卫宁健康、创业慧康、东华医为", "门槛最低，薪资一般，当跳板"],
        ["**第四层：通用 AI 交付", "BOSS 上搜 FDE 的中小公司", "练手用，面试完记问题就走"],
    ],
    [2.8, 8.6, 6.6],
)
P(before=6)
item([("内推路线：", True, NAVY),
      ("你认识的医生 → 医院信息科 → 厂商驻场工程师 → 内推；脉脉直接搜目标公司在职员工，话术是「我在医院端做了 6 年，想问问你们 FDE 实际在做什么」。", False, None)])

# 薪资
P(before=10)
bar("六、薪资锚点", "别报低了")
P(before=3)
table(
    ["城市 / 类型", "合理区间", "你的谈判筹码"],
    [
        ["**成都", "12–18K", "6 年行业 + 已上线系统 + 不用 6 个月熟悉医院"],
        ["**深圳 / 上海", "15–25K", "同上，且能覆盖西南/西北区域客户"],
        ["**孵化营 / 校招类", "6–11K", "只当入场券，接受但谈 1 年后转正条件"],
    ],
    [4.0, 4.2, 9.8],
)
P(before=6)
item([("谈薪三条：", True, NAVY),
      ("① 不先报数字，先问「这个岗位的预算区间是多少」；② 有 offer 再谈，别在初面谈钱；③ 用「已上线系统 + 6 年客户资源」撑住下限，不用「我愿意学」换低薪。", False, None)])

# 止损
P(before=10)
bar("七、止损线与备选", "什么情况下换打法")
P(before=3)
table(
    ["信号", "判断", "动作"],
    [
        ["W12 投 20 份，0 回复", "简历或作品集有问题", "找人改简历；把 Demo 放到简历第一屏"],
        ["W14 面试全挂技术面", "动手能力不够", "补 4 周再战，不要连续面"],
        ["W20 仍无 offer", "FDE 门槛对你偏高", "**降一档投「售前/解决方案/实施」**，先入场再转"],
        ["面试官反复问「你会写代码吗」", "岗位其实是开发岗", "放弃这家，不是你的问题"],
    ],
    [5.0, 5.0, 8.0],
)

# 每周检查表
P(before=10)
bar("八、每周自检", "每周日花 5 分钟打勾")
for t in [
    "本周投入 ≥ 10 小时（记录实际数字，不估）",
    "本周有一个可交付物（不是「学完了」，是「做出来了」）",
    "本周写过代码（哪怕 20 行），没有连续三天不碰",
    "把本周做的东西截图存档，进作品集素材库",
    "更新一次「卡住清单」：哪一步没搞懂，写下来",
    "有没有接触一个行业里的人（线上聊也算）",
]:
    check(t, size=9.4)

P(before=8)
item([("最后一句：", True, NAVY),
      ("这份清单最大的风险不是难度，是你在第 3 周就停下来。所以每周只看当周那一行，不要看全表。", False, None)])

P(before=10)
bot = doc.add_table(rows=1, cols=1)
no_borders(bot)
fixed_layout(bot)
cell_margins(bot, top=0.1, bottom=0.1, left=0, right=0)
bc = bot.cell(0, 0)
shade(bc, NAVY_HEX)
bp = bc.paragraphs[0]
bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
bp.paragraph_format.space_after = Pt(0)
R(bp, "16 周　·　3 个 Demo　·　1 份作品集　·　20 份投递　·　1 个 offer", size=8, color=(0xD5, 0xDE, 0xE7), spacing=1.4)

out = r"C:\Users\Administrator\Desktop\FDE转型-16周执行清单.docx"
doc.save(out)
print("saved:", out)
