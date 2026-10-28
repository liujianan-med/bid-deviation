# -*- coding: utf-8 -*-
"""
12_multi_product_compare.py —— 多产品横向参数对比表（IPL 强脉冲光）

场景：一份招标文件（图片/扫描件）的 N 条技术要求，对照 M 家 N 个型号的产品参数，
输出一张可直接看的对比 Excel，不符合招标要求的字符标红。

这是「投标商务外包」里最常被要的一张表——经销商拿到标，第一件事就是问
「我们这几款哪个能投、差在哪几条」。

判定口径（与 06_compare 的三态一致，另加"未载明"）：
    ok      符合（资料里有明确、可引用的数值/功能）
    partial 部分符合（有的条件满足、有的没写；投标时必须补口径或要厂家盖章）
    bad     不符合（数值不达标或功能形态相反，标红）
    na      资料未载明（不等于不符合，但不能直接写进偏离表当"响应"）

单元格文本用 <<...>> 包住的部分会渲染成红色。
"""

import os
from datetime import date

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
OUT = os.path.join(DESKTOP, f"克拉玛依独山子-强脉冲光-参数对比表-{date.today():%Y%m%d}.xlsx")

TEAL = "0F6E56"
HEAD_BG = "1F3864"
RED = "C00000"
AMBER = "BF8F00"
GRAY = "808080"
GREEN = "1D9E75"

PRODUCTS = [
    ("vvl3", "伟思 白金VVL3\n（大师光·白金版）"),
    ("m5", "Master M5\n（标底文件）"),
    ("ky_sn", "科英 KL-L(SN)\n强光一号·30w"),
    ("ky_zn", "科英 KL-L(ZN)\n强光之星·40w"),
    ("ky_z", "科英 KL-L(Z)\n强光之星·50w"),
    ("qz_1", "奇致 NBL-Ⅰ\n55w"),
    ("qz_2", "奇致 NBL-Ⅱ\n新氧·55w"),
]

# VVL3 判定口径（【白金VVL3】强脉冲光治疗仪挂网参数-20260402.docx）：
#   资料载明且达标 -> ok；载明但有缺口 -> partial（缺口字符标红）；
#   挂网参数完全没有的条款 -> blank（单元格留白，浅黄底，由使用者自行补充）。
D_VVL3 = {
    "1": ("智能子脉冲模式（总能量分散为多个子脉冲输出）；<<双/三脉冲独立调节未载明>>", "partial"),
    "2": ("载明改善皮肤外观；<<减少毛发未载明>>（以注册证适用范围为准）", "partial"),
    "3": ("单灯泵浦，发射平顶方波改善皮肤外观——与招标原文一致", "ok"),
    "4": ("触摸屏 ≥15 寸，达标", "ok"),
    "5": ("滤光片自动识别 + 热插拔两条均载明（六方中唯一齐全）", "ok"),
    "6": ("滑动治疗和定点治疗两种治疗模式", "ok"),
    "7": ("至少 8 种：宽谱 6 种（515/560/590/615/640/695-1200nm）+ 窄谱 Acne、Vascular", "ok"),
    "8": ("脉冲能量密度 ≥40 J/cm²（上限达标）；<<下限 5 J/cm² 未载明>>", "partial"),
    "9": ("脉宽 2-20ms；脉冲串宽度 2-800ms；步进 1ms——与招标完全一致", "ok"),
    "10": ("子脉冲间隔 5-150ms——与招标完全一致", "ok"),
    "11": ("最大支持 10 个子脉冲——刚好达线", "ok"),
    "12": ("最大治疗频率 20Hz（厂家挂网参数载明）", "ok"),
    "13": ("至少 4 种：15×35 / 15×8 / 15×15 / φ10mm——刚好达线", "ok"),
    "14": ("蓝宝石接触持续式冷却，冷却范围 -2—32℃", "ok"),
    "15": ("设备自带能量校准功能，操作者自主校准，偏差 ±10%（六方中唯一载明）", "ok"),
    "16": ("内智方案库：自定义方案保存、故障提示、密码设置——三项齐全", "ok"),
    "17": ("能量输出方式：手具和脚踏两种可选——明确载明", "ok"),
    "18": ("滤光片 8 片（<<差 1 片>>）；适配器 4 种（<<差 3 个>>）", "partial"),
    "19": ("", "blank"),
}

# ---- 招标文件 19 条（克拉玛依市独山子人民医院，图片 OCR）----
TENDER = [
    ("1", "▲", "支持多脉冲模式独立调节（双/三脉冲可调）", "多脉冲模式；双/三脉冲独立可调"),
    ("2", "", "适用于改善皮肤外观治疗及减少毛发的治疗，以注册适用范围为准", "适用范围含：改善皮肤外观 + 减少毛发"),
    ("3", "", "强脉冲光治疗仪为单灯泵浦，通过发射平顶方波改善皮肤外观", "单灯泵浦；平顶方波"),
    ("4", "", "屏幕：≥15英寸触摸屏", "≥15 英寸触摸"),
    ("5", "▲", "治疗手具有滤光片自动识别和热插拔的功能", "自动识别 + 热插拔"),
    ("6", "", "具有定点治疗模式", "定点模式"),
    ("7", "▲", "具有≥7种滤光片", "≥7 种"),
    ("8", "", "脉冲能量密度范围：≥5J/cm²~40J/cm²", "5~40 J/cm²"),
    ("9", "", "脉宽范围：≥2-20ms；脉冲串宽度范围：2ms-800ms；步进：1ms", "脉宽 2~20ms；串宽 2~800ms；步进 1ms"),
    ("10", "", "子脉冲间隔的范围是：5ms-150ms", "5~150 ms"),
    ("11", "▲", "≥10个子脉冲", "≥10 个"),
    ("12", "", "频率≥20Hz", "≥20 Hz"),
    ("13", "", "光斑适配器数量≥4种", "≥4 种"),
    ("14", "", "蓝宝石接触持续式冷却", "蓝宝石 + 持续接触冷却"),
    ("15", "", "具有设备自带的能量校准功能，操作者可自主进行校准；调整精度：偏差±10%", "自主校准；偏差 ±10%"),
    ("16", "", "具有设备自有定义方案保存，故障提示，密码设置的功能", "方案保存 + 故障提示 + 密码"),
    ("17", "", "能量输出方式：手具和脚踏两种可选", "手具 + 脚踏"),
    ("18", "", "配备手具1把、滤光片9片、光斑适配器7个", "手具1 / 滤光片9 / 适配器7"),
    ("19", "", "设备设计使用寿命≥10年", "≥10 年"),
]

# ---- 六方参数对照 ----
# 结构：no -> {产品key: (文本, 状态)}
D = {
    "1": {
        "m5": ("“重复脉冲串/单脉冲串”两种工作模式，子脉冲最大≥10个可调", "ok"),
        "ky_sn": ("SOPT 技术，支持多脉冲独立调节，最多 12 个子脉冲", "ok"),
        "ky_zn": ("SOPT 技术，支持多脉冲独立调节，最多 12 个子脉冲", "ok"),
        "ky_z": ("SOPT 技术，支持多脉冲独立调节，最多 12 个子脉冲", "ok"),
        "qz_1": ("2-10 个脉冲可调，最高支持 10 脉冲", "ok"),
        "qz_2": ("2-10 个脉冲可调，最高支持 10 脉冲", "ok"),
    },
    "2": {
        "m5": ("血管性疾病、皮肤表浅色素性疾病、减少毛发；<<未载明“改善皮肤外观”>>", "partial"),
        "ky_sn": ("改善皮肤外观（含寻常痤疮）、血管性疾病、皮肤表浅色素性疾病、减少毛发", "ok"),
        "ky_zn": ("改善皮肤外观（含寻常痤疮）、血管性疾病、皮肤表浅色素性疾病、减少毛发", "ok"),
        "ky_z": ("改善皮肤外观（含寻常痤疮）、血管性疾病、皮肤表浅色素性疾病、减少毛发", "ok"),
        "qz_1": ("浅表良性色素性病变、浅表良性血管性病变、减少毛发；<<未载明“改善皮肤外观”>>", "partial"),
        "qz_2": ("浅表良性色素性病变、浅表良性血管性病变、减少毛发；<<未载明“改善皮肤外观”>>", "partial"),
    },
    "3": {
        "m5": ("<<“多灯输出技术”>>、<<“进口双灯管光源”>>——与单灯泵浦相反", "bad"),
        "ky_sn": ("原装进口<<单氙灯>>；SOPT 技术达平顶输出", "ok"),
        "ky_zn": ("原装进口<<单氙灯>>；SOPT 技术达平顶输出", "ok"),
        "ky_z": ("原装进口<<单氙灯>>；SOPT 技术达平顶输出", "ok"),
        "qz_1": ("<<双氙灯>>；无能量尖峰、无能量衰减（近平顶但非方波）", "bad"),
        "qz_2": ("<<双氙灯>>；无能量尖峰、无能量衰减（近平顶但非方波）", "bad"),
    },
    "4": {
        "m5": ("<<12.1 英寸>>高分辨触摸屏", "bad"),
        "ky_sn": ("<<15 英寸>>触摸屏", "ok"),
        "ky_zn": ("<<15 英寸>>触摸屏", "ok"),
        "ky_z": ("<<15 英寸>>触摸屏", "ok"),
        "qz_1": ("<<13 英寸>>屏", "bad"),
        "qz_2": ("<<13 英寸>>屏", "bad"),
    },
    "5": {
        "m5": ("智能<<热插拔>>技术；<<未载明滤光片自动识别>>", "partial"),
        "ky_sn": ("滤光片<<自动识别>>；<<未载明热插拔>>", "partial"),
        "ky_zn": ("滤光片<<自动识别>>；<<未载明热插拔>>", "partial"),
        "ky_z": ("滤光片<<自动识别>>；<<未载明热插拔>>", "partial"),
        "qz_1": ("<<可插拔>>滤光片技术；<<未载明自动识别>>", "partial"),
        "qz_2": ("<<可插拔>>滤光片技术；<<未载明自动识别>>", "partial"),
    },
    "6": {
        "m5": ("<<资料未载明定点治疗模式>>", "na"),
        "ky_sn": ("定点治疗", "ok"),
        "ky_zn": ("定点治疗", "ok"),
        "ky_z": ("定点治疗", "ok"),
        "qz_1": ("定点", "ok"),
        "qz_2": ("定点", "ok"),
    },
    "7": {
        "m5": ("<<≥6 种>>波长滤光片（差 1 种）", "bad"),
        "ky_sn": ("<<7 个>>（420/515/560/590/615/640/695nm）——刚好达线", "ok"),
        "ky_zn": ("<<8 个>>（7 个基础 + Acne）", "ok"),
        "ky_z": ("<<9 个>>（8 个 + Vascular）", "ok"),
        "qz_1": ("<<6 个>>（差 1 种）", "bad"),
        "qz_2": ("<<6 个>>（差 1 种）", "bad"),
    },
    "8": {
        "m5": ("可调最大能量密度<<≥35 J/cm²>>（上限差 5）", "bad"),
        "ky_sn": ("<<5—40 J/cm²>>，步进 1", "ok"),
        "ky_zn": ("<<5—40 J/cm²>>，步进 1", "ok"),
        "ky_z": ("<<5—40 J/cm²>>，步进 1", "ok"),
        "qz_1": ("<<2—35 J/cm²>>（上限差 5）", "bad"),
        "qz_2": ("<<2—35 J/cm²>>（上限差 5）", "bad"),
    },
    "9": {
        "m5": ("子脉冲宽度最小<<≤230μs>>；脉冲串宽度<<≥800ms>>；<<未载明步进1ms>>", "partial"),
        "ky_sn": ("脉宽<<2—20ms>>；脉冲串宽度<<1000ms>>；<<未载明步进1ms>>", "partial"),
        "ky_zn": ("脉宽<<2—20ms>>；脉冲串宽度<<1000ms>>；<<未载明步进1ms>>", "partial"),
        "ky_z": ("脉宽<<2—20ms>>；脉冲串宽度<<1000ms>>；<<未载明步进1ms>>", "partial"),
        "qz_1": ("最小脉宽<<≤500μs>>；脉冲串宽度<<5—800ms>>；<<未载明步进1ms>>", "partial"),
        "qz_2": ("最小脉宽<<≤500μs>>；脉冲串宽度<<5—800ms>>；<<未载明步进1ms>>", "partial"),
    },
    "10": {
        "m5": ("<<未载明子脉冲间隔>>（仅有脉冲串重复时间 1~5s，非同一参数）", "na"),
        "ky_sn": ("脉冲延迟<<5—150ms>>", "ok"),
        "ky_zn": ("脉冲延迟<<5—150ms>>", "ok"),
        "ky_z": ("脉冲延迟<<5—150ms>>", "ok"),
        "qz_1": ("<<未载明子脉冲间隔>>", "na"),
        "qz_2": ("<<未载明子脉冲间隔>>", "na"),
    },
    "11": {
        "m5": ("子脉冲个数最大<<≥10 个>>可调", "ok"),
        "ky_sn": ("最多<<12 个>>子脉冲", "ok"),
        "ky_zn": ("最多<<12 个>>子脉冲", "ok"),
        "ky_z": ("最多<<12 个>>子脉冲", "ok"),
        "qz_1": ("<<2—10 个>>，最高支持 10 脉冲——刚好达线", "ok"),
        "qz_2": ("<<2—10 个>>，最高支持 10 脉冲——刚好达线", "ok"),
    },
    "12": {
        "m5": ("脉冲串重复时间<<1~5s>>（约 0.2~1Hz）", "partial"),
        "ky_sn": ("重复频率：单次/0.5/1/2/<<3Hz>>", "partial"),
        "ky_zn": ("重复频率：单次/0.5/1/2/<<3Hz>>", "partial"),
        "ky_z": ("重复频率：单次/0.5/1/2/<<3Hz>>", "partial"),
        "qz_1": ("<<未载明整机重复频率>>", "na"),
        "qz_2": ("<<未载明整机重复频率>>", "na"),
    },
    "13": {
        "m5": ("<<≥4 种>>大小光斑适配器", "ok"),
        "ky_sn": ("Φ10mm、35×15mm、15×15mm、15×8mm 共<<4 种>>", "ok"),
        "ky_zn": ("Φ10mm、35×15mm、15×15mm、15×8mm 共<<4 种>>", "ok"),
        "ky_z": ("Φ10mm、35×15mm、15×15mm、15×8mm 共<<4 种>>", "ok"),
        "qz_1": ("φ11mm、φ7mm、15×15mm、45×15mm(ST) 共<<4 种>>", "ok"),
        "qz_2": ("φ11mm、φ7mm、15×15mm、45×15mm(ST) 共<<4 种>>", "ok"),
    },
    "14": {
        "m5": ("0~30℃ 逐级冷却 + 水冷双灯管；<<未载明蓝宝石接触式>>", "partial"),
        "ky_sn": ("<<蓝宝石>>导光晶体；2~6℃ / 0~30℃ 可调", "ok"),
        "ky_zn": ("<<蓝宝石>>导光晶体；2~6℃ / 0~30℃ 可调", "ok"),
        "ky_z": ("<<蓝宝石>>导光晶体；2~6℃ / 0~30℃ 可调", "ok"),
        "qz_1": ("<<蓝宝石>>；0~30℃ 可调，步进 5℃", "ok"),
        "qz_2": ("<<蓝宝石>>；0~30℃ 可调，步进 5℃", "ok"),
    },
    "15": {
        "m5": ("<<资料未载明能量校准功能>>", "na"),
        "ky_sn": ("<<资料未载明能量校准功能>>", "na"),
        "ky_zn": ("<<资料未载明能量校准功能>>", "na"),
        "ky_z": ("<<资料未载明能量校准功能>>", "na"),
        "qz_1": ("<<资料未载明能量校准功能>>", "na"),
        "qz_2": ("<<资料未载明能量校准功能>>", "na"),
    },
    "16": {
        "m5": ("可存储治疗方案<<≥5 个>>；<<未载明故障提示、密码设置>>", "partial"),
        "ky_sn": ("<<资料未载明方案保存/故障提示/密码设置>>", "na"),
        "ky_zn": ("<<资料未载明方案保存/故障提示/密码设置>>", "na"),
        "ky_z": ("<<资料未载明方案保存/故障提示/密码设置>>", "na"),
        "qz_1": ("屏幕自带<<可存储治疗方案>>；<<未载明故障提示、密码设置>>", "partial"),
        "qz_2": ("屏幕自带<<可存储治疗方案>>；<<未载明故障提示、密码设置>>", "partial"),
    },
    "17": {
        "m5": ("<<资料未载明手具/脚踏双输出>>", "na"),
        "ky_sn": ("<<资料未载明手具/脚踏双输出>>", "na"),
        "ky_zn": ("<<资料未载明手具/脚踏双输出>>", "na"),
        "ky_z": ("<<资料未载明手具/脚踏双输出>>", "na"),
        "qz_1": ("<<资料未载明手具/脚踏双输出>>", "na"),
        "qz_2": ("<<资料未载明手具/脚踏双输出>>", "na"),
    },
    "18": {
        "m5": ("<<资料未载明随机配置清单>>", "na"),
        "ky_sn": ("滤光片<<7 片>>（差 2）；适配器 4 种（<<差 3 个>>）", "bad"),
        "ky_zn": ("滤光片<<8 片>>（差 1）；适配器 4 种（<<差 3 个>>）", "partial"),
        "ky_z": ("滤光片<<9 片>>（达标）；适配器 4 种（<<差 3 个>>）", "partial"),
        "qz_1": ("滤光片<<6 片>>（差 3）；适配器 4 种（<<差 3 个>>）", "bad"),
        "qz_2": ("滤光片<<6 片>>（差 3）；适配器 4 种（<<差 3 个>>）", "bad"),
    },
    "19": {
        "m5": ("主机使用期限<<≥10 年>>", "ok"),
        "ky_sn": ("使用年限<<10 年>>", "ok"),
        "ky_zn": ("使用年限<<10 年>>", "ok"),
        "ky_z": ("使用年限<<10 年>>", "ok"),
        "qz_1": ("使用年限<<10 年>>", "ok"),
        "qz_2": ("使用年限<<10 年>>", "ok"),
    },
}

# 合并 VVL3 列
for _no, _v in D_VVL3.items():
    D[_no]["vvl3"] = _v

ICON = {"ok": "✅ ", "partial": "⚠️ ", "bad": "❌ ", "na": "➖ ", "blank": ""}
TXTCOLOR = {"ok": "FF000000", "partial": "FFBF8F00", "bad": "FFC00000", "na": "FF808080", "blank": ""}
BLANK_FILL = PatternFill("solid", fgColor="FFFFF2CC")   # 留白待填（浅黄提示）

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def rich(text, state, size=9):
    """把 <<红字>> 解析成富文本；未被包住的按状态色显示。"""
    base_color = {"bad": "FF000000", "partial": "FF000000", "na": "FF808080", "ok": "FF000000"}[state]
    icon = ICON[state]
    parts = [TextBlock(InlineFont(rFont="等线", sz=size, color=base_color), icon)]

    i = 0
    while i < len(text):
        s = text.find("<<", i)
        if s < 0:
            parts.append(TextBlock(InlineFont(rFont="等线", sz=size, color=base_color), text[i:]))
            break
        if s > i:
            parts.append(TextBlock(InlineFont(rFont="等线", sz=size, color=base_color), text[i:s]))
        e = text.find(">>", s)
        if e < 0:
            parts.append(TextBlock(InlineFont(rFont="等线", sz=size, color=base_color), text[s:]))
            break
        parts.append(TextBlock(InlineFont(rFont="等线", sz=size, color="FFC00000", b=True),
                               text[s + 2:e]))
        i = e + 2
    return CellRichText(*parts)


def sheet_tender(wb):
    ws = wb.active
    ws.title = "① 招标参数提取"
    heads = ["序号", "标记", "技术参数原文（招标文件·克拉玛依独山子人民医院）", "提取的量化指标", "类型"]
    widths = [6, 6, 62, 34, 10]
    for c, (h, w) in enumerate(zip(heads, widths), 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(name="等线", size=10, bold=True, color="FFFFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[1].height = 30

    for r, (no, mark, req, num) in enumerate(TENDER, 2):
        ws.cell(r, 1, no).alignment = Alignment(horizontal="center", vertical="center")
        m = ws.cell(r, 2, mark or "")
        m.alignment = Alignment(horizontal="center", vertical="center")
        m.font = Font(name="等线", size=11, bold=True, color="FFC00000")
        c3 = ws.cell(r, 3, req)
        c3.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c4 = ws.cell(r, 4, num)
        c4.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c4.font = Font(name="等线", size=9, color="FF404040")
        c5 = ws.cell(r, 5, "数值" if any(ch.isdigit() for ch in num) else "功能/形态")
        c5.alignment = Alignment(horizontal="center", vertical="center")
        c5.font = Font(name="等线", size=9, color="FF404040")
        for c in range(1, 6):
            ws.cell(r, c).border = BORDER
        ws.cell(r, 1).font = Font(name="等线", size=10)
        ws.cell(r, 3).font = Font(name="等线", size=10)
        ws.row_dimensions[r].height = 30 if len(req) < 40 else 42

    ws.freeze_panes = "A2"
    ws.cell(21, 3, "共 19 条技术参数，其中 ▲ 关键参数 4 条（第 1、5、7、11 条）。"
                   "▲ 条款不满足通常按实质性负偏离处理，须逐条落实口径后再投标。").font = \
        Font(name="等线", size=9, italic=True, color="FFC00000")
    return ws


def sheet_compare(wb):
    ws = wb.create_sheet("② 七方参数对比")
    heads = ["序号", "标记", "招标技术参数"] + [p[1] for p in PRODUCTS] + ["结论"]
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 5
    ws.column_dimensions["C"].width = 36
    for col in "DEFGHIJ":
        ws.column_dimensions[col].width = 24
    ws.column_dimensions["K"].width = 11

    for c, h in enumerate(heads, 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(name="等线", size=9, bold=True, color="FFFFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[1].height = 40

    for r, (no, mark, req, _num) in enumerate(TENDER, 2):
        ws.cell(r, 1, no).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 1).font = Font(name="等线", size=10)
        m = ws.cell(r, 2, mark or "")
        m.alignment = Alignment(horizontal="center", vertical="center")
        m.font = Font(name="等线", size=11, bold=True, color="FFC00000")

        c3 = ws.cell(r, 3, req)
        c3.font = Font(name="等线", size=9, bold=bool(mark))
        c3.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        stats = []
        for j, (key, _n) in enumerate(PRODUCTS):
            text, state = D[no][key]
            cell = ws.cell(r, 4 + j)
            if state == "blank":
                cell.value = None          # 留白，由使用者自行填写
                cell.fill = BLANK_FILL
                stats.append(state)
                continue
            cell.value = rich(text, state)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if state == "bad":
                cell.fill = PatternFill("solid", fgColor="FFFDE9E9")
            elif state == "partial":
                cell.fill = PatternFill("solid", fgColor="FFFFF8E6")
            stats.append(state)

        vvl3 = stats[0]
        ky = stats[2:5]                     # 科英三款
        others = [stats[1], stats[5], stats[6]]   # M5 / 奇致两款
        if any(s == "blank" for s in stats):
            verdict = "待补资料"
        elif all(s == "ok" for s in stats):
            verdict = "全符合"
        elif vvl3 == "ok" and all(s == "ok" for s in ky):
            verdict = "多方可投"
        elif vvl3 == "ok" and not any(s == "ok" for s in ky):
            verdict = "仅VVL3"
        elif vvl3 != "ok" and all(s == "ok" for s in ky):
            verdict = "仅科英"
        elif vvl3 == "ok" or any(s == "ok" for s in ky):
            verdict = "部分可投"
        elif all(s == "bad" for s in ky):
            verdict = "科英亦不符"
        elif all(s in ("bad", "partial", "na") for s in stats):
            verdict = "均不足"
        else:
            verdict = "待确认"

        cv = ws.cell(r, 11, verdict)
        cv.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cv.font = Font(name="等线", size=9, bold=True,
                       color={"全符合": "FF1D9E75", "多方可投": "FF185FA5",
                              "仅VVL3": "FF0F6E56", "仅科英": "FF0F6E56",
                              "部分可投": "FF185FA5", "待补资料": "FFBF8F00",
                              "均不足": "FFC00000", "科英亦不符": "FFC00000",
                              "待确认": "FFBF8F00"}[verdict])

        for c in range(1, 12):
            ws.cell(r, c).border = BORDER
        ws.row_dimensions[r].height = 52

    # 汇总行
    r = len(TENDER) + 3
    ws.cell(r, 3, "各家符合条数（✅ / 共 19 条）").font = Font(name="等线", size=9, bold=True)
    for j, (key, _n) in enumerate(PRODUCTS):
        ok = sum(1 for no, *_ in TENDER if D[no][key][1] == "ok")
        bad = sum(1 for no, *_ in TENDER if D[no][key][1] == "bad")
        cell = ws.cell(r, 4 + j, f"✅{ok}　❌{bad}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(name="等线", size=10, bold=True,
                         color="FF1D9E75" if ok >= 8 else "FFC00000")
        cell.border = BORDER

    # 图例
    r += 2
    ws.cell(r, 3, "图例：✅ 符合　⚠️ 部分符合（有条款未载明）　❌ 不符合（红字为具体不合规处）　"
                  "➖ 资料未载明　⬜ 浅黄空白＝VVL3 挂网参数未载明该条，留白待补").font = \
        Font(name="等线", size=9, italic=True, color="FF404040")
    ws.freeze_panes = "D2"
    return ws


def sheet_gap(wb):
    ws = wb.create_sheet("③ 不符合项与结论")
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 40
    heads = ["序号", "不满足/存疑条款", "受影响型号", "处理办法"]
    for c, h in enumerate(heads, 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(name="等线", size=10, bold=True, color="FFFFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[1].height = 26

    items = [
        ("1", "▲多脉冲模式独立调节（双/三脉冲可调）", "VVL3（子脉冲可调，独立调节未载明）",
         "VVL3 有智能子脉冲模式，但“双/三脉冲独立调节”字样挂网参数没写。"
         "可从说明书或检测报告找对应描述补口径；找不到则按“具备子脉冲模式，脉冲数可调”响应。"),
        ("2", "适用范围须含改善皮肤外观 + 减少毛发", "VVL3（载明改善皮肤外观，减少毛发未载明）",
         "挂网参数非注册证。以 VVL3 注册证适用范围为准填写，若注册证含减毛发即闭环。"),
        ("3", "单灯泵浦、平顶方波", "Master M5（双灯/多灯）、奇致（双氙灯）",
         "形态性条款，无法用口径弥补。M5 与奇致在本条即被排除，除非改标。VVL3 与科英达标。"),
        ("4", "屏幕≥15 英寸", "Master M5（12.1\"）、奇致（13\"）",
         "硬件参数，不可响应。VVL3（≥15寸）与科英 15\" 达标。"),
        ("7", "▲滤光片≥7 种", "Master M5（6 种）、奇致（6 种）",
         "▲ 条款，负偏离通常按无效处理。VVL3 8 种、科英 ZN 8/Z 9 富余，SN 恰好 7 种达线。"),
        ("8", "能量密度范围 5~40 J/cm²", "Master M5（35）、奇致（35）、VVL3（下限未载明）",
         "M5/奇致上限差 5，属明确负偏离。VVL3 上限≥40 达标，仅 5 J/cm² 下限需从检测报告补。"),
        ("12", "频率≥20Hz 口径", "除 VVL3 外全部不足或未载明",
         "VVL3 挂网参数明确载明“最大治疗频率 20Hz”，是唯一可直接响应的。"
         "其余厂家按整机频率全部不足；建议同时争取答疑澄清口径，双保险。"),
        ("15", "能量校准功能、偏差±10%", "除 VVL3 外全部未载明",
         "VVL3 载明“设备自带能量校准，偏差±10%”，可直接响应并附检测报告。"
         "科英等其余厂家须出具盖章《技术参数确认函》。"),
        ("17", "手具 + 脚踏双输出", "除 VVL3 外全部未载明",
         "VVL3 明确载明“手具和脚踏两种可选”。此条通常随机标配，其余厂家建议索取配置清单原件。"),
        ("18", "滤光片 9 片、适配器 7 个", "VVL3（8 片/4 种）、科英 SN（7 片/4 个）、ZN（8 片）、Z（9 片/4 个）、奇致（6 片）",
         "属随机配置条款，可承诺加配达标（成本项），但要在偏离表写实际交付数量。"
         "VVL3 需加配滤光片 1 片、适配器 3 个。"),
        ("19", "设备设计使用寿命≥10年", "VVL3（挂网参数未载明，表中留白）",
         "待补：从 VVL3 说明书 / 注册证 / 技术要求中找“使用期限”条款填入空白处；"
         "若注册证写 10 年即达标。"),
    ]
    for r, (no, clause, affected, fix) in enumerate(items, 2):
        ws.cell(r, 1, no).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(r, 1).font = Font(name="等线", size=10, bold=True)
        c2 = ws.cell(r, 2, clause)
        c2.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c2.font = Font(name="等线", size=9)
        c3 = ws.cell(r, 3, affected)
        c3.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c3.font = Font(name="等线", size=9, color="FFC00000")
        c4 = ws.cell(r, 4, fix)
        c4.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c4.font = Font(name="等线", size=9)
        for c in range(1, 5):
            ws.cell(r, c).border = BORDER
        ws.row_dimensions[r].height = 46

    r = len(items) + 3
    conclusions = [
        ("结论一", "这份标书的技术参数与伟思白金 VVL3 的挂网参数高度吻合，指向性明显。",
         "第 3 条与 VVL3 挂网原文逐字相同（单灯泵浦+平顶方波），第 9/10 条数值完全一致"
         "（脉宽 2-20ms、串宽 2-800ms、步进 1ms、间隔 5-150ms）；"
         "第 5 条“自动识别+热插拔”、第 15 条“±10% 校准”、第 17 条“手具+脚踏”"
         "——这几条其余五家全部未载明或只中一半，唯 VVL3 齐全。"
         "这份标的十有八九是照着 VVL3 的挂网参数写的。"),
        ("结论二", "VVL3 是本标最强候选：14/19 明确符合，4 条部分符合，1 条留白待补。",
         "部分符合的第 18 条是加配成本项（滤光片 1 片 + 适配器 3 个）；"
         "第 1/2/8 条补口径即可闭环；第 19 条使用寿命从说明书或注册证找依据填入空白。"
         "全表无一条硬性不符合。"),
        ("结论三", "科英三款仍可投（各 12/19），但第 15/17 条须厂家出盖章确认函。",
         "M5 与奇致维持不建议：单灯/双灯、屏幕尺寸、滤光片数量、能量密度上限四条硬伤，"
         "话术和承诺补不了。"),
        ("下一步", "投标前必须落实的三件事",
         "① 第 19 条：从 VVL3 说明书/注册证查“使用期限”填入表中空白；"
         "② 第 12 条：VVL3 载明 20Hz 可直接响应，建议同步争取答疑澄清口径；"
         "③ 第 18 条：加配滤光片 1 片、适配器 3 个，成本计入报价。"),
    ]
    for i, (tag, title, body) in enumerate(conclusions):
        rr = r + i * 4
        t = ws.cell(rr, 1, tag)
        t.font = Font(name="等线", size=10, bold=True, color="FF0F6E56")
        t.alignment = Alignment(horizontal="center", vertical="top")
        h = ws.cell(rr, 2, title)
        h.font = Font(name="等线", size=10, bold=True)
        h.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        b = ws.cell(rr + 1, 2, body)
        b.font = Font(name="等线", size=9, color="FF404040")
        b.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.merge_cells(start_row=rr + 1, start_column=2, end_row=rr + 1, end_column=4)
        ws.row_dimensions[rr + 1].height = 56
    return ws


def main():
    wb = Workbook()
    sheet_tender(wb)
    sheet_compare(wb)
    sheet_gap(wb)
    try:
        wb.save(OUT)
    except PermissionError:
        stem, ext = os.path.splitext(OUT)
        n = 2
        while True:
            alt = f"{stem}-v{n}{ext}"
            try:
                wb.save(alt)
                print(f"目标文件被占用，已另存：{alt}")
                break
            except PermissionError:
                n += 1
    else:
        print(f"已生成：{OUT}")
    ok_total = {}
    for key, name in PRODUCTS:
        ok_total[name.split("\n")[0]] = sum(1 for no, *_ in TENDER if D[no][key][1] == "ok")
    print("符合条数：", ok_total)


if __name__ == "__main__":
    main()
