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
PRODUCT_FILE = os.path.join(BASE, "picov_params.json")       # 老路径，保留兼容；新加载走 product_db
OUT_FILE = os.path.join(BASE, "deviation_report.json")


def num(x):
    """把 '1.7' / 2 / '±10' / '10.1' 转成 float，转不动返回 None

    招标文件常写「≤±10%」，模型抽出来是 '±10'，不带正负号清洗就会解析失败，
    整条被判成「要求无数值」——这类脏数据在真实标书里是常态，解析层必须容忍。
    """
    if x is None:
        return None
    s = str(x).strip().replace("±", "").replace("＋", "").replace(" ", "")
    s = s.rstrip("%").rstrip("％")
    try:
        return float(s)
    except (ValueError, TypeError):
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        return float(m.group()) if m else None


# 单位 -> (量纲, 换算到基准单位的系数, 基准单位)
# 只处理投标文件里真正常见的同类量；量表之外的组合一律判存疑，绝不瞎猜。
UNITS = {
    "mrad": ("角度", 1.0, "mrad"), "rad": ("角度", 1000.0, "mrad"),
    "弧度": ("角度", 1000.0, "mrad"), "毫弧度": ("角度", 1.0, "mrad"),
    "mW": ("功率", 1.0, "mW"), "W": ("功率", 1000.0, "mW"),
    "kW": ("功率", 1e6, "mW"), "GW": ("功率", 1e12, "mW"),
    "J": ("能量", 1.0, "J"), "mJ": ("能量", 1e-3, "J"),
    "ps": ("脉宽", 1.0, "ps"), "ns": ("脉宽", 1000.0, "ps"),
    "ms": ("脉宽", 1e9, "ps"),
    "mm": ("长度", 1.0, "mm"), "cm": ("长度", 10.0, "mm"),
    "Hz": ("频率", 1.0, "Hz"), "kHz": ("频率", 1e3, "Hz"),
    "KG": ("重量", 1.0, "KG"), "kg": ("重量", 1.0, "KG"),
    "%": ("百分比", 1.0, "%"),
    "英寸": ("屏幕尺寸", 1.0, "英寸"),
}


def unit_info(unit):
    """返回 (量纲, 换算系数, 基准单位)；未知单位返回 (None, 1.0, '')"""
    if not unit:
        return (None, 1.0, "")
    u = str(unit).strip()
    return UNITS.get(u, (None, 1.0, u))


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


def _chunk_hit(w: str, pv: str) -> bool:
    """整词命中之外，词内 ≥4 字的连续片段命中也算。
    「掺钕钇铝石榴石激光器」的片段「掺钕钇铝石榴石」能命中产品描述
    「Nd:YAG掺钕钇铝石榴石」——招标措辞和产品资料措辞几乎从来不会逐字一致。
    """
    if w in pv:
        return True
    L = len(w)
    for size in range(L - 1, 3, -1):
        for i in range(L - size + 1):
            if w[i:i + size] in pv:
                return True
    return False


# 中文数字 -> 阿拉伯数字（「七关节」「十英寸」这类产品描述里的量）
_CN_DIGIT = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
             "六": 6, "七": 7, "八": 8, "九": 9}


def feature_number(text: str):
    """从功能描述里提取首个数值（阿拉伯或简单中文数字），取不到返回 None"""
    s = str(text or "")
    m = re.search(r"\d+(?:\.\d+)?", s)
    if m:
        return float(m.group())
    m = re.search(r"([一二三四五六七八九十]+)", s)
    if m:
        cn = m.group(1)
        if cn == "十":
            return 10.0
        if "十" in cn:
            a, _, b = cn.partition("十")
            return float(_CN_DIGIT.get(a, 1) * 10 + _CN_DIGIT.get(b, 0))
        return float(_CN_DIGIT.get(cn)) if len(cn) == 1 and cn in _CN_DIGIT else None
    return None


def compare(req: dict, prod: dict) -> tuple:
    """返回 (判定, 理由)。req 是招标要求的结构化结果，prod 是产品参数库条目"""
    op = req.get("operator")
    req_min, req_max = parse_value(req)

    # 特殊类型分支放在最前：它们自己解析原始值，不依赖通用数值解析
    if prod["type"] == "wavelength":
        # 模型对「532nm或1064nm」这类要求会输出 "532或1064"，拆开逐个验证。
        # 只认 ≥100 的数字当波长：「1064nm±5nm」里的允差 ±5 若不排除，
        # 会被当成要求波长 5nm，整条误判负偏离（真实招标文件里发生过）。
        raw = str(req.get("value") or "") + str(req.get("unit") or "")
        nums = [n for n in (num(x) for x in re.findall(r"[\d.]+", raw))
                if n is not None and n >= 100]
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
            overlap = any(_chunk_hit(w, pv) for w in tokens) or (pv in rv)
            return ("响应", f"功能类：要求与产品「{pv[:20]}」要点吻合") if overlap \
                else ("存疑", f"功能类需人工比对：要求「{rv[:20]}」 vs 产品「{pv[:20]}」")
        return ("存疑", "要求无数值，产品侧也无法直接比对")

    # ---- 要求有数值，产品没披露 ----
    if prod["type"] == "missing":
        return ("存疑", f"产品资料未披露该参数（{prod.get('note','')}）——需查检测报告或询问产品部")

    # ---- 要求有数值，产品是功能描述：尝试从描述里抠出数值（「七关节」-> 7）----
    if prod["type"] == "feature":
        n = feature_number(prod.get("value"))
        if n is None:
            return ("存疑", f"要求带数值而产品侧是描述「{str(prod.get('value'))[:20]}」，需人工比对")
        prod_min = prod_max = n

    # ---- 特殊类型：IPX 防护等级（IPX8 >= IPX4 按数字比）----
    if prod["type"] == "ipx":
        req_num = req_max if req_max is not None else num(req.get("value"))
        if req_num is None:
            return ("存疑", "要求侧数值解析失败")
        return ("正偏离", f"产品 IPX{prod['value']} 优于要求 IPX{int(req_num)}") \
            if prod["value"] > req_num else ("响应", f"产品 IPX{prod['value']} 满足要求 IPX{int(req_num)}")

    # ---- 常规数值/区间比较 ----
    # feature 分支可能已从描述里抠出数值（prod_min/prod_max 已赋值），不能重置
    if prod["type"] in ("number", "ipx"):
        prod_min = prod_max = num(prod.get("value"))
    elif prod["type"] == "range":
        prod_min, prod_max = num(prod.get("min")), num(prod.get("max"))
    elif prod["type"] != "feature":
        prod_min, prod_max = None, None
    if prod_min is None:
        return ("存疑", "产品侧数值解析失败")

    # ---- 单位归一化 ----
    # 招标写「发散角<0.3弧度」、产品写「150mrad」，直接比大小会判成负偏离。
    # 单位不同的量必须先换算到同一基准，否则结论方向会整条反掉。
    # 基准取「招标文件的单位」：结论里的数字就是评委看到的那个数字，不出现 2e+12mW 这种鬼东西。
    r_dim, r_factor, r_base = unit_info(req.get("unit"))
    p_dim, p_factor, p_base = unit_info(prod.get("unit"))
    U = ""
    if r_dim and r_dim == p_dim and r_factor:
        prod_min = prod_min * p_factor / r_factor
        prod_max = prod_max * p_factor / r_factor
        U = str(req.get("unit") or "").strip()
    elif r_dim and p_dim and r_dim != p_dim:
        return ("存疑", f"单位量纲不一致（要求 {req.get('unit')} vs 产品 {prod.get('unit')}），需人工核对")

    def fmt(v):
        return f"{v:g}{U}" if U else f"{v:g}"

    if op == ">=":
        if prod_min > req_min: return ("正偏离", f"产品 {fmt(prod_min)} > 要求 ≥{fmt(req_min)}")
        if prod_min == req_min: return ("响应", f"产品 {fmt(prod_min)} = 要求 ≥{fmt(req_min)}")
        return ("负偏离", f"产品 {fmt(prod_min)} < 要求 ≥{fmt(req_min)}")
    if op == ">":
        if prod_min > req_min: return ("响应", f"产品 {fmt(prod_min)} > 要求 >{fmt(req_min)}")
        return ("负偏离", f"产品 {fmt(prod_min)} 不严格大于要求 {fmt(req_min)}")
    if op == "<=":
        if prod_max < req_max: return ("正偏离", f"产品 {fmt(prod_max)} < 要求 ≤{fmt(req_max)}")
        if prod_max == req_max: return ("响应", f"产品 {fmt(prod_max)} = 要求 ≤{fmt(req_max)}")
        return ("负偏离", f"产品 {fmt(prod_max)} > 要求 ≤{fmt(req_max)}")
    if op == "<":
        if prod_max < req_max: return ("正偏离", f"产品 {fmt(prod_max)} < 要求 <{fmt(req_max)}")
        return ("负偏离", f"产品 {fmt(prod_max)} 不严格小于要求 {fmt(req_max)}")
    if op == "=":
        if prod_min == req_min: return ("响应", f"产品 {fmt(prod_min)} = 要求 {fmt(req_min)}")
        # 数值型要求写「=」时，产品更好也应算正偏离（如要求=16，产品16响应）
        return ("存疑", f"要求 ={fmt(req_min)}，产品 {fmt(prod_min)}，是否构成偏离需人工判断")
    if op == "range":
        covers = prod_min <= req_min and prod_max >= req_max
        same = prod_min == req_min and prod_max == req_max
        if same: return ("响应", f"产品区间 [{fmt(prod_min)},{fmt(prod_max)}] 与要求 [{fmt(req_min)},{fmt(req_max)}] 一致")
        if covers: return ("正偏离", f"产品区间 [{fmt(prod_min)},{fmt(prod_max)}] 覆盖要求 [{fmt(req_min)},{fmt(req_max)}]")
        # 产品是单点承诺而要求写成区间时，别急着判负偏离：
        # 上限承诺 ≤ 区间下限（脉宽≤248ps vs 要求350~500ps）本身就是优于要求。
        if prod["type"] == "number" and prod.get("operator") == "<=" and prod_max <= req_min:
            return ("正偏离", f"产品 {fmt(prod_max)}（上限承诺）优于要求区间 [{fmt(req_min)},{fmt(req_max)}]")
        if prod["type"] == "number" and prod.get("operator") == ">=" and prod_min >= req_max:
            return ("正偏离", f"产品 {fmt(prod_min)}（下限承诺）优于要求区间 [{fmt(req_min)},{fmt(req_max)}]")
        if prod_min >= req_min and prod_max <= req_max:
            return ("存疑", f"产品区间 [{fmt(prod_min)},{fmt(prod_max)}] 在要求范围内，是否够用需人工判断")
        return ("负偏离", f"产品区间 [{fmt(prod_min)},{fmt(prod_max)}] 未覆盖要求 [{fmt(req_min)},{fmt(req_max)}]")
    return ("存疑", f"未处理的比较符 {op}")


def run_compare(reqs: list, product: dict) -> dict:
    """三态判定核心函数：批量条目 vs 产品参数库 -> 报告 dict

    reqs  ：03 的产出格式 [{row, item_id, mark, source_text, parsed}, ...]
    product：{key: 产品参数条目}
    供命令行 main() 和 API（步骤 5）共用——判定逻辑只写这一份。
    """
    from param_alias import ALIAS

    report = []
    counts = {"响应": 0, "正偏离": 0, "负偏离": 0, "存疑": 0}

    for r in reqs:
        p = r["parsed"]
        name = p.get("param_name", "?")
        key = ALIAS.get(name)
        mark = r.get("mark") or " "

        if key is None:
            verdict, reason = "存疑", f"参数「{name}」未匹配到产品库（映射表需补充）"
        else:
            REQ_TEXT["text"] = r["source_text"]  # 功能类比对用原文兜底
            verdict, reason = compare(p, product[key])

        counts[verdict] += 1
        report.append({"row": r["row"], "item_id": r["item_id"], "mark": mark,
                       "requirement": r["source_text"], "param": name,
                       "verdict": verdict, "reason": reason})

    return {"counts": counts, "items": report}


def main():
    import sys

    from product_db import load as load_product

    spec = None
    if "--product" in sys.argv:
        i = sys.argv.index("--product")
        spec = sys.argv[i + 1] if len(sys.argv) > i + 1 else None
    prod = load_product(spec)
    product = prod["params"]

    with open(BATCH_FILE, encoding="utf-8") as f:
        reqs = json.load(f)

    print("=" * 78)
    print(f"品类：{prod['name']}  产品参数 {len(product)} 项  （{prod['product_path']}）")
    print("=" * 78)
    print(f"{'行':>3} {'判定':<4} {'招标要求':<34} {'说明'}")
    print("=" * 78)
    result = run_compare(reqs, product)
    for r in result["items"]:
        flag = {"正偏离": "+", "负偏离": "!", "存疑": "?", "响应": " "}[r["verdict"]]
        mark = r.get("mark") or " "
        print(f"{r['row']:>3} {flag} {mark:<2} {r['requirement'][:32]:<34} {r['reason'][:40]}")

    counts = result["counts"]
    print("=" * 78)
    print(f"响应 {counts['响应']} | 正偏离 {counts['正偏离']} | "
          f"负偏离 {counts['负偏离']} | 存疑 {counts['存疑']}  （共 {len(result['items'])} 条）")

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"报告已保存 -> data/deviation_report.json")


if __name__ == "__main__":
    main()
