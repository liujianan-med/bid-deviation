# -*- coding: utf-8 -*-
"""
07_make_docx.py —— 把三态判定结果变成可编辑的 Word 偏离表

输出两个部分：
    1. 偏离表正文：5 列，格式完全对齐真实投标文件
       （序号 / 招标文件技术要求条款 / 投标文件内容对应简述 / 响应情况 / 证明资料及索引）
    2. 存疑清单：程序不敢下结论的条目，列明原因和建议去向

红线对齐（来自招标文件）：
    投标响应内容必须写具体数值，只写"满足/符合"视为投标无效。
    所以响应内容列一律从产品参数库取真实数值，取不到就明确标
    【待人工确认】——宁可让空白醒目，也不制造"看起来填完了"的假象。

运行：
    .venv\\Scripts\\python.exe scripts\\07_make_docx.py
"""

import json
import os
from datetime import date

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn

from param_alias import ALIAS

BASE = os.path.join(os.path.dirname(__file__), "..", "data")
REPORT_FILE = os.path.join(BASE, "deviation_report.json")
PRODUCT_FILE = os.path.join(BASE, "picov_params.json")
EVIDENCE_FILE = os.path.join(BASE, "evidence_map.json")
OUT_FILE = os.path.join(os.path.dirname(__file__), "..",
                        f"皮秒偏离表-自动生成-{date.today():%Y%m%d}.docx")

RED = RGBColor(0xC0, 0x00, 0x00)
GREEN = RGBColor(0x00, 0x70, 0x30)
GRAY = RGBColor(0x60, 0x60, 0x60)

# 证明材料索引库：精确到「文件 + 编号 + 页码 + 序号/条款」
# 数据在 data/evidence_map.json 里维护，脚本只负责按参数名取用。
with open(EVIDENCE_FILE, encoding="utf-8") as _f:
    EVIDENCE_MAP = json.load(_f)
EVIDENCE_ITEMS = EVIDENCE_MAP["items"]


def set_font(run, size=10.5, bold=False, color=None):
    """中文字体必须同时设置 eastAsia，否则中文会是默认字体"""
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def render_product(p: dict, name: str) -> str:
    """按产品参数类型渲染正式响应文案"""
    if p["type"] == "number":
        op = {"<=": "≤", ">=": "≥", "=": ""}.get(p.get("operator"), "")
        return f"{name}：{op}{p['value']}{p.get('unit','')}"
    if p["type"] == "range":
        return f"{name}：{p['min']}~{p['max']}{p.get('unit','')}"
    if p["type"] == "ipx":
        return f"{name}：IPX{p['value']}"
    if p["type"] == "wavelength":
        return f"{name}：{'和'.join(str(w) for w in p['value'])}nm 双波长，允差±{p.get('tolerance',5)}nm"
    if p["type"] == "feature":
        return f"{name}：{p['value']}"
    return None


def response_text(item: dict, product: dict) -> str:
    """生成投标响应内容——红线：必须带具体数值"""
    name = item["param"]
    # 模型起的参数名要过别名映射才能查产品库
    p = product.get(ALIAS.get(name, name))

    if item["verdict"] == "正偏离":
        body = render_product(p, name) if p else None
        return f"{body}（优于招标要求，详见证明资料）" if body else "【待人工确认】响应内容需人工填写具体数值"
    if p is None:
        return f"【待人工确认】{name}：响应内容需人工填写具体数值"
    body = render_product(p, name)
    return body or f"【待人工确认】{p.get('note','产品资料未披露')}——需查检测报告后填写"


DOC_NO = {k: v.get("no", "") for k, v in EVIDENCE_MAP["documents"].items()}


def fmt_cite(e: dict) -> str:
    """把一条出处拼成「文件 编号 第X页 序号Y（条款Z）」"""
    parts = [e["doc"]]
    no = e.get("no") or DOC_NO.get(e["doc"], "")
    if no:
        parts.append(no)
    parts.append(e["loc"])
    return " ".join(parts)


def evidence_entries(item: dict) -> list:
    """取该参数的全部出处；无出处返回空列表"""
    ev = EVIDENCE_ITEMS.get(item["param"])
    if ev is None or (isinstance(ev, dict) and ev.get("unavailable")):
        return []
    return ev


def evidence_text(item: dict) -> str:
    """索引列：精确到页码与序号，最多列两条（其余见附二明细）"""
    ev = EVIDENCE_ITEMS.get(item["param"])
    if ev is None:
        return "【待建索引】该参数尚未在 evidence_map.json 中登记出处"
    if isinstance(ev, dict) and ev.get("unavailable"):
        return "【待补件】现有资料无法指向具体页码，见存疑清单"
    cites = [fmt_cite(e) for e in ev[:2]]
    if len(ev) > 2:
        cites.append(f"等{len(ev)}处，详见附二")
    text = "；\n".join(cites)
    # 存疑条目：出处只能证明部分内容，必须写明缺哪一项，不能让评委误以为全部已证
    if item["verdict"] == "存疑":
        text += "；\n【缺证】上列资料不能完全覆盖本条要求，须补充后再引用"
    return text


def build_document(report: dict, product: dict) -> Document:
    """核心函数：报告 dict + 产品参数库 -> Word 文档对象

    供命令行 main() 和 API（步骤 5）共用——排版逻辑只写这一份。
    调用方拿到 Document 后自行 doc.save(path)。
    """
    doc = Document()

    # ---- 页头（对齐真实偏离表）----
    h = doc.add_paragraph()
    set_font(h.add_run("九、技术响应偏离表"), size=16, bold=True)
    for line in ["项目名称：", "项目编号/包号：", "投标人名称（公章）：",
                 f"日      期：{date.today():%Y年%m月%d日}"]:
        p = doc.add_paragraph()
        set_font(p.add_run(line), size=12)

    note = doc.add_paragraph()
    set_font(note.add_run("注：本表由工具自动生成，【待人工确认】处必须人工核实填写后才能提交。"
                          "响应内容只写“满足/符合”而无具体数值的，视为未实质性满足（投标无效）。"),
             size=9, color=RED)

    # ---- 偏离表正文 ----
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    headers = ["序号", "招标文件技术要求条款", "投标文件内容对应简述", "响应情况", "证明资料及索引"]
    for i, htext in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.paragraphs[0].alignment = 1
        set_font(cell.paragraphs[0].add_run(htext), bold=True)

    for idx, item in enumerate(report["items"], 1):
        row = table.add_row()
        verdict = item["verdict"]
        display = {"正偏离": "正偏离", "响应": "响应", "负偏离": "负偏离",
                   "存疑": "存疑（待人工）"}[verdict]

        cells = row.cells
        set_font(cells[0].paragraphs[0].add_run(str(idx)))
        set_font(cells[1].paragraphs[0].add_run(item["requirement"]))
        set_font(cells[2].paragraphs[0].add_run(response_text(item, product)))
        color = {"正偏离": GREEN, "负偏离": RED, "存疑": RED, "响应": None}[verdict]
        set_font(cells[3].paragraphs[0].add_run(display), bold=(verdict != "响应"), color=color)
        set_font(cells[4].paragraphs[0].add_run(evidence_text(item)))

    # 列宽统一设置（Word 对列宽的生效与否和表格布局模式有关，这里允许自动调整兜底）
    for col, w in zip(table.columns, [1.1, 5.6, 5.2, 1.9, 4.2]):
        col.width = Cm(w)

    # ---- 附二：证明材料索引明细（每条对应到页码+序号+该处载明数值）----
    doc.add_paragraph()
    h2a = doc.add_paragraph()
    set_font(h2a.add_run("附二、证明材料索引明细（评委核查用，可随投标文件一并提交）"),
             size=13, bold=True)
    cap = doc.add_paragraph()
    set_font(cap.add_run("说明：页码为原件印刷页码；检验报告 2024-GL-0320 原件共109页，"
                         "说明书共30页。查证时按「文件→页码→序号/条款」逐级对应。"),
             size=9, color=GRAY)

    t3 = doc.add_table(rows=1, cols=4)
    t3.style = "Table Grid"
    for i, htext in enumerate(["要求条目", "证明材料", "索引位置", "该处载明内容（数值）"]):
        c = t3.rows[0].cells[i]
        c.paragraphs[0].alignment = 1
        set_font(c.paragraphs[0].add_run(htext), bold=True, size=9)

    for idx, item in enumerate(report["items"], 1):
        entries = evidence_entries(item)
        if not entries:
            ev = EVIDENCE_ITEMS.get(item["param"]) or {}
            row = t3.add_row().cells
            set_font(row[0].paragraphs[0].add_run(f"{idx}.{item['param']}"), size=9)
            set_font(row[1].paragraphs[0].add_run("暂无出处"), size=9, color=RED)
            set_font(row[2].paragraphs[0].add_run("—"), size=9)
            set_font(row[3].paragraphs[0].add_run(ev.get("reason", "索引库未登记")), size=9, color=RED)
            continue
        for j, e in enumerate(entries):
            row = t3.add_row().cells
            set_font(row[0].paragraphs[0].add_run(f"{idx}.{item['param']}" if j == 0 else ""), size=9)
            doc_label = e["doc"] + (" " + e["no"] if e.get("no") else
                                    (" " + DOC_NO[e["doc"]] if DOC_NO.get(e["doc"]) else ""))
            set_font(row[1].paragraphs[0].add_run(doc_label), size=9)
            set_font(row[2].paragraphs[0].add_run(e["loc"]), size=9)
            set_font(row[3].paragraphs[0].add_run(e.get("detail", "")), size=9)

    for col, w in zip(t3.columns, [3.0, 3.6, 5.0, 5.8]):
        col.width = Cm(w)

    # ---- 存疑清单 ----
    doc.add_paragraph()
    h2 = doc.add_paragraph()
    set_font(h2.add_run("附：存疑条目清单（人工复核用，本页不随投标文件提交）"), size=13, bold=True, color=RED)

    pending = [it for it in report["items"] if it["verdict"] == "存疑"]
    if pending:
        t2 = doc.add_table(rows=1, cols=4)
        t2.style = "Table Grid"
        for i, htext in enumerate(["要求条目", "存疑原因", "建议去向", "人工结论（填写）"]):
            c = t2.rows[0].cells[i]
            c.paragraphs[0].alignment = 1
            set_font(c.paragraphs[0].add_run(htext), bold=True)
        for it in pending:
            row = t2.add_row().cells
            set_font(row[0].paragraphs[0].add_run(it["requirement"][:50]))
            set_font(row[1].paragraphs[0].add_run(it["reason"]))
            entries = evidence_entries(it)
            if entries:
                goto = fmt_cite(entries[0]) + ("（措辞需与招标原文逐字核对）"
                                               if "措辞" in it["reason"] or "字面" in it["reason"] else "")
            elif isinstance(EVIDENCE_ITEMS.get(it["param"]), dict):
                goto = "产品部出证：见附二该条「暂无出处」说明"
            else:
                goto = "产品部/说明书，索引库待建"
            set_font(row[2].paragraphs[0].add_run(goto))
            set_font(row[3].paragraphs[0].add_run(""), size=12)

    return doc


def main():
    with open(REPORT_FILE, encoding="utf-8") as f:
        report = json.load(f)
    with open(PRODUCT_FILE, encoding="utf-8") as f:
        product = {p["key"]: p for p in json.load(f)["params"]}

    doc = build_document(report, product)
    doc.save(OUT_FILE)
    print(f"已生成: {OUT_FILE}")
    print(f"偏离表 {len(report['items'])} 条；存疑清单 "
          f"{sum(1 for it in report['items'] if it['verdict'] == '存疑')} 条")
    print("用 WPS/Word 打开即可编辑。【待人工确认】红色标记处必须处理后才能投标使用。")


if __name__ == "__main__":
    main()
