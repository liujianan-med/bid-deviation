# -*- coding: utf-8 -*-
"""
06_compare.py —— 三态判定：招标要求 vs 产品参数 -> 响应 / 正偏离 / 负偏离 / 存疑

设计原则（这是本项目最重要的一课）：

    分工是「LLM 负责理解语言，代码负责比较数字」。
    - 语义问题（"最大支持10个"是 >= 还是 <=？）交给模型
    - 数值比较（2.0 > 1.7 吗？）绝对不交给模型——模型算术会错，
      而且同样的输入要得到同样的输出，这叫确定性，只有代码能给。

    判定规则（对齐投标实务）：
      要求 >= X ：产品 > X  -> 正偏离（优于要求）
                  产品 = X  -> 响应
                  产品 < X  -> 负偏离
      要求 <= X ：方向同理
      要求 区间 ：产品区间覆盖要求 -> 正偏离；完全一致 -> 响应；覆盖不足 -> 负偏离
      无法判定 ：一律进「存疑」清单，绝不瞎猜。存疑不是失败，
                 是设计出来的输出——投标里一条判错的代价远大于一条留给人复核。

运行：
    .venv\\Scripts\\python.exe scripts\\06_compare.py
"""

import json
import os
import re

BASE = os.path.join(os.path.dirname(__file__), "..", "data")
BATCH_FILE = os.path.join(BASE, "batch_extract.json")       # 模型抽取的招标要求
PRODUCT_FILE = os.path.join(BASE, "picov_params.json")       # 产品参数库
OUT_FILE = os.path.join(BASE, "deviation_report.json")


def num(x):
    """把 '1.7' / 2 / '10.1' 转成 float，转不动返回 None"""
    try:
        return float(str(x).strip())
    except (ValueError, TypeError):
        return None


# compare() 需要招标原文做功能类比对的兜底，用模块级变量传入（简单但有效）
REQ_TEXT = {"text": ""}


def parse_value(parsed: dict):
    """从模型抽取结果里拿数值和单位。区间返回 (min, max)"""
    v = parsed.get("value")
    if v is None or v == "":
        return None, None
    s = str(v).strip()
    # 区间：2-10 / 0.13-7.75 / 1~40
    m = re.match(r"^(-?[\d.]+)\s*[-~～—]\s*(-?[\d.]+)$", s)
    if m:
        return num(m.group(1)), num(m.group(2))
    return num(s), num(s)


def compare(req: dict, prod: dict) -> tuple:
    """返回 (判定, 理由)。req 是招标要求的结构化结果，prod 是产品参数库条目"""
    op = req.get("operator")
    req_min, req_max = parse_value(req)

    # 特殊类型分支放在最前：它们自己解析原始值，不依赖通用数值解析
    if prod["type"] == "wavelength":
        # 模型对「532nm或1064nm」这类要求会输出 "532或1064"，拆开逐个验证
        raw = str(req.get("value") or "") + str(req.get("unit") or "")
        nums = [n for n in (num(x) for x in re.findall(r"[\d.]+", raw)) if n is not None]
        if not nums:
            return ("存疑", "要求侧波长解析失败")
        missing = [n for n in nums if n not in prod["value"]]
        if missing:
            return ("负偏离", f"产品波长 {prod['value']} 不含要求的 {missing}nm")
        return ("响应", f"产品具备要求的全部波长 {nums}nm")

    # ---- 要求没有数值：功能类/描述类 ----
    if op == "none" or req_min is None:
        if prod["type"] == "feature":
            # 两边都是描述。模型有时不给 value，就用招标原文比对
            rv = str(req.get("value") or "") or REQ_TEXT.get("text", "")
            pv = str(prod.get("value", ""))
            tokens = [w for w in re.findall(r"[\u4e00-\u9fa5A-Za-z]{2,}", rv) if len(w) >= 2]
            # 双向包含：关键词命中产品描述，或整条产品描述是要求的子串
            overlap = any(w in pv for w in tokens) or (pv in rv)
            return ("响应", f"功能类：要求与产品「{pv[:20]}」要点吻合") if overlap \
                else ("存疑", f"功能类需人工比对：要求「{rv[:20]}」 vs 产品「{pv[:20]}」")
        return ("存疑", "要求无数值，产品侧也无法直接比对")

    # ---- 要求有数值，产品没披露 ----
    if prod["type"] == "missing":
        return ("存疑", f"产品资料未披露该参数（{prod.get('note','')}）——需查检测报告或询问产品部")

    # ---- 特殊类型：IPX 防护等级（IPX8 >= IPX4 按数字比）----
    if prod["type"] == "ipx":
        req_num = req_max if req_max is not None else num(req.get("value"))
        if req_num is None:
            return ("存疑", "要求侧数值解析失败")
        return ("正偏离", f"产品 IPX{prod['value']} 优于要求 IPX{int(req_num)}") \
            if prod["value"] > req_num else ("响应", f"产品 IPX{prod['value']} 满足要求 IPX{int(req_num)}")

    # ---- 常规数值/区间比较 ----
    prod_min, prod_max = None, None
    if prod["type"] in ("number", "ipx"):
        prod_min = prod_max = num(prod.get("value"))
    elif prod["type"] == "range":
        prod_min, prod_max = num(prod.get("min")), num(prod.get("max"))
    if prod_min is None:
        return ("存疑", "产品侧数值解析失败")

    if op == ">=":
        if prod_min > req_min: return ("正偏离", f"产品 {prod_min} > 要求 ≥{req_min}")
        if prod_min == req_min: return ("响应", f"产品 {prod_min} = 要求 ≥{req_min}")
        return ("负偏离", f"产品 {prod_min} < 要求 ≥{req_min}")
    if op == ">":
        if prod_min > req_min: return ("响应", f"产品 {prod_min} > 要求 >{req_min}")
        return ("负偏离", f"产品 {prod_min} 不严格大于要求 {req_min}")
    if op == "<=":
        if prod_max < req_max: return ("正偏离", f"产品 {prod_max} < 要求 ≤{req_max}")
        if prod_max == req_max: return ("响应", f"产品 {prod_max} = 要求 ≤{req_max}")
        return ("负偏离", f"产品 {prod_max} > 要求 ≤{req_max}")
    if op == "<":
        if prod_max < req_max: return ("响应", f"产品 {prod_max} < 要求 <{req_max}")
        return ("负偏离", f"产品 {prod_max} 不严格小于要求 {req_max}")
    if op == "=":
        if prod_min == req_min: return ("响应", f"产品 {prod_min} = 要求 {req_min}")
        # 数值型要求写「=」时，产品更好也应算正偏离（如要求=16，产品16响应）
        return ("存疑", f"要求 ={req_min}，产品 {prod_min}，是否构成偏离需人工判断")
    if op == "range":
        covers = prod_min <= req_min and prod_max >= req_max
        same = prod_min == req_min and prod_max == req_max
        if same: return ("响应", f"产品区间 [{prod_min},{prod_max}] 与要求 [{req_min},{req_max}] 一致")
        if covers: return ("正偏离", f"产品区间 [{prod_min},{prod_max}] 覆盖要求 [{req_min},{req_max}]")
        if prod_min >= req_min and prod_max <= req_max:
            return ("存疑", f"产品区间 [{prod_min},{req_max}] 在要求范围内，是否够用需人工判断")
        return ("负偏离", f"产品区间 [{prod_min},{prod_max}] 未覆盖要求 [{req_min},{req_max}]")
    return ("存疑", f"未处理的比较符 {op}")


def main():
    with open(BATCH_FILE, encoding="utf-8") as f:
        reqs = json.load(f)
    with open(PRODUCT_FILE, encoding="utf-8") as f:
        product = {p["key"]: p for p in json.load(f)["params"]}

    # 参数名 -> 产品库 key 的映射已抽到 param_alias.py（06/07 共用单一数据源）
    from param_alias import ALIAS

    report = []
    counts = {"响应": 0, "正偏离": 0, "负偏离": 0, "存疑": 0}

    print("=" * 78)
    print(f"{'行':>3} {'判定':<4} {'招标要求':<34} {'说明'}")
    print("=" * 78)
    for r in reqs:
        p = r["parsed"]
        name = p.get("param_name", "?")
        key = ALIAS.get(name)
        mark = r.get("mark") or " "
        text = r["source_text"][:32]

        if key is None:
            verdict, reason = "存疑", f"参数「{name}」未匹配到产品库（映射表需补充）"
        else:
            REQ_TEXT["text"] = r["source_text"]  # 功能类比对用原文兜底
            verdict, reason = compare(p, product[key])

        counts[verdict] += 1
        flag = {"正偏离": "+", "负偏离": "!", "存疑": "?", "响应": " "}[verdict]
        print(f"{r['row']:>3} {flag} {mark:<2} {text:<34} {reason[:40]}")
        report.append({"row": r["row"], "item_id": r["item_id"], "mark": mark,
                       "requirement": r["source_text"], "param": name,
                       "verdict": verdict, "reason": reason})

    print("=" * 78)
    print(f"响应 {counts['响应']} | 正偏离 {counts['正偏离']} | "
          f"负偏离 {counts['负偏离']} | 存疑 {counts['存疑']}  （共 {len(reqs)} 条）")

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump({"counts": counts, "items": report}, f, ensure_ascii=False, indent=2)
    print(f"报告已保存 -> data/deviation_report.json")


if __name__ == "__main__":
    main()
