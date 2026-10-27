# -*- coding: utf-8 -*-
"""
09_make_sishi_report.py —— 单个项目的定稿偏离表（逐条人工核定过出处与口径）

与 06/07 全自动流水线的关系：
    07 负责「排版」，本脚本负责「内容」。
    06 的自动判定给的是第一稿（结构对不对、哪几条要人看），
    本脚本用 data/sishi_decisions.json 里的核定结果替换判定与响应文案——
    每一条都对应纸质资料的具体页码/条款，不带一个编造的数字。

输入：
    data/params_tender.json     08 从招标文件 Word 切出的条款原文
    data/sishi_decisions.json   逐条核定：判定 / 响应内容 / 出处 / 存疑原因与建议
输出：
    桌面\\第四师医院-皮秒技术偏离表-YYYYMMDD.docx

运行：
    .venv\\Scripts\\python.exe scripts\\09_make_sishi_report.py
"""

import json
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import importlib
m07 = importlib.import_module("07_make_docx")

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")
TENDER_FILE = os.path.join(BASE, "params_tender.json")
DECISION_FILE = os.path.join(BASE, "sishi_decisions.json")
PRODUCT_FILE = os.path.join(BASE, "picov_params.json")
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")

VERDICT_KEYS = ("响应", "正偏离", "负偏离", "存疑")


def main():
    tender = json.load(open(TENDER_FILE, encoding="utf-8"))
    decision = json.load(open(DECISION_FILE, encoding="utf-8"))
    product = {p["key"]: p for p in json.load(open(PRODUCT_FILE, encoding="utf-8"))["params"]}

    items = []
    counts = {k: 0 for k in VERDICT_KEYS}
    for idx, src in enumerate(tender["items"], 1):
        iid = str(src["item_id"])
        dec = decision["items"].get(iid)
        if dec is None:
            print(f"[警告] 条款 {iid} 未在核定文件中登记，跳过")
            continue
        verdict = dec["verdict"]
        counts[verdict] += 1
        items.append({
            "row": iid, "item_id": iid, "mark": src["mark"],
            "requirement": src["requirement"], "param": dec["param"],
            "verdict": verdict, "reason": dec["reason"],
            "response": dec["response"], "evidence": dec.get("evidence", []),
            "advice": dec.get("advice", ""),
        })

    report = {"counts": counts, "items": items,
              "source": tender["source"], "device": tender.get("device", "")}

    doc = m07.build_document(
        report, product,
        project=decision["project"], device=decision["device"],
        note="注：响应内容所列数据均取自第三方检验报告（2024-GL-0320）、产品使用说明书、"
             "配置清单与技术白皮书，索引列页码为原件印刷页码，可逐项核查。"
             "响应内容只写“满足/符合”而无具体数值的，视为未实质性满足（投标无效）。")
    out = os.path.join(DESKTOP, f"第四师医院-皮秒技术偏离表-{date.today():%Y%m%d}.docx")
    doc.save(out)

    print(f"已生成：{out}")
    print(f"共 {len(items)} 条：" + " | ".join(f"{k} {v}" for k, v in counts.items()))
    pending = [it for it in items if it["verdict"] == "存疑"]
    print(f"存疑 {len(pending)} 条需复核：" + "、".join(it["param"] for it in pending))


if __name__ == "__main__":
    main()
