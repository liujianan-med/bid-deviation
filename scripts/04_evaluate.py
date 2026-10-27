# -*- coding: utf-8 -*-
"""
04_evaluate.py —— 给项目算出第一个准确率数字

评测逻辑：
    data/batch_extract.json   模型的输出（31 条）
    data/eval_gold.json       人工标注的标准答案（gold）
    逐条比对 operator 和 is_measurable 两个字段，输出：
      - 两个指标：operator 准确率、is_measurable 准确率
      - 每一条错在哪（错例清单是评测里最值钱的部分）

为什么只评这两个字段？
    param_name 叫「峰值功率」还是「最大峰值功率」，value 写 "1.7" 还是 1.7，
    都不影响投标判定；但 operator 错了（>= 写成 =）和 measurable 判错，
    会直接导致偏离表三态判错。评测要对齐业务目标，不是对齐模型输出格式。

运行：
    .venv\\Scripts\\python.exe scripts\\04_evaluate.py
"""

import json
import os
import sys

# 不带参数：评测皮秒批（默认）
# 带参数  ：python 04_evaluate.py <模型输出json> <gold json> <报告输出json>
if len(sys.argv) >= 4:
    PRED_FILE = os.path.join(os.path.dirname(__file__), "..", sys.argv[1])
    GOLD_FILE = os.path.join(os.path.dirname(__file__), "..", sys.argv[2])
    REPORT_TAG = sys.argv[3]
else:
    BASE = os.path.join(os.path.dirname(__file__), "..", "data")
    PRED_FILE = os.path.join(BASE, "batch_extract.json")
    GOLD_FILE = os.path.join(BASE, "eval_gold.json")
    REPORT_TAG = "eval_report.json"
OUT_FILE = os.path.join(os.path.dirname(PRED_FILE), REPORT_TAG)


def norm_op(op: str) -> str:
    """归一化比较符，避免 = 和 == 这类无意义差异影响评分"""
    mapping = {"==": "=", "＜": "<", "＞": ">", "≧": ">=", "≤": "<=", "≥": ">="}
    return mapping.get((op or "").strip(), (op or "").strip())


def main():
    with open(PRED_FILE, encoding="utf-8") as f:
        preds = {p["row"]: p["parsed"] for p in json.load(f)}
    with open(GOLD_FILE, encoding="utf-8") as f:
        golds = json.load(f)["labels"]

    op_total = op_ok = 0
    ms_total = ms_ok = 0
    errors = []

    for g in golds:
        row = g["row"]
        p = preds.get(row)
        if p is None:
            errors.append({"row": row, "item_id": g["item_id"], "问题": "模型输出缺失"})
            continue

        # --- is_measurable ---
        ms_total += 1
        pred_ms = bool(p.get("is_measurable"))
        gold_ms = bool(g["is_measurable"])
        if pred_ms == gold_ms:
            ms_ok += 1
        else:
            errors.append({"row": row, "item_id": g["item_id"],
                           "问题": f"is_measurable 判错：模型 {pred_ms}，应为 {gold_ms}",
                           "note": g.get("note", "")})

        # --- operator（标注里注明不计分的行跳过）---
        if "operator 不计分" in g.get("note", ""):
            continue
        op_total += 1
        pred_op = norm_op(p.get("operator"))
        gold_op = norm_op(g["operator"])
        if pred_op == gold_op:
            op_ok += 1
        else:
            errors.append({"row": row, "item_id": g["item_id"],
                           "问题": f"operator 判错：模型 '{pred_op}'，应为 '{gold_op}'",
                           "原文": p.get("value", ""), "note": g.get("note", "")})

    print("=" * 60)
    print("评测结果")
    print("=" * 60)
    print(f"is_measurable 准确率 : {ms_ok}/{ms_total} = {ms_ok/ms_total:.1%}")
    print(f"operator 准确率      : {op_ok}/{op_total} = {op_ok/op_total:.1%}")
    print()

    if errors:
        print(f"错例清单（{len(errors)} 条）：")
        for e in errors:
            print(f"  行{e['row']} {e['item_id']}")
            print(f"    {e['问题']}")
            if e.get("note"):
                print(f"    标注依据: {e['note']}")
    else:
        print("没有错例。")

    # 评测报告落盘，README 里引用的就是这个数字
    report = {
        "sample_size": len(golds),
        "is_measurable": {"correct": ms_ok, "total": ms_total,
                          "accuracy": round(ms_ok / ms_total, 3)},
        "operator": {"correct": op_ok, "total": op_total,
                     "accuracy": round(op_ok / op_total, 3)},
        "errors": errors,
    }
    out = OUT_FILE
    with open(out, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"\n报告已保存 -> {os.path.relpath(out)}")


if __name__ == "__main__":
    main()
