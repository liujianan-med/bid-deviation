# -*- coding: utf-8 -*-
"""
08_parse_tender_docx.py —— 从整本招标文件 Word 里定位并解析目标设备的技术参数

真实招标文件长这样（第四师医院项目）：
    全文 1300+ 段落，皮秒参数埋在「（三）皮秒激光治疗仪」小节里，
    前后还有无影灯、取皮植皮、激光光电平台、二氧化碳激光仪等其他设备。

解析思路（代码找结构，LLM 理解语义——和全项目同一个分工原则）：
    1. 找设备小节标题：「（一）无影灯」「（三）皮秒激光治疗仪」这种格式
    2. 选出目标设备小节（默认关键词「皮秒」，可选）
    3. 小节内按条款编号切分：新条款 = 「数字.」开头且不是 3.1 这种子条款
    4. 多行条款合并（15.控制系统 一条占 4 段），▲/★ 从编号前后两个位置都能认

输出 params_tender.json，schema 与 params.json 一致，后面 03/06/07 全部直接复用。

运行：
    .venv\\Scripts\\python.exe scripts\\08_parse_tender_docx.py <招标文件.docx> [设备关键词]
"""

import json
import re
import sys
from pathlib import Path

from docx import Document

BASE_DIR = Path(__file__).resolve().parent.parent
OUT_FILE = BASE_DIR / "data" / "params_tender.json"

# 设备小节标题：全角括号 + 中文数字，如「（三）皮秒激光治疗仪」
DEVICE_HEADER_RE = re.compile(r"^（([一二三四五六七八九十]+)）\s*(\S.{0,40})$")
# 新条款：数字 + 顿号/点，但排除 3.1 / 18.1 这类子条款（后面跟数字的不算）
CLAUSE_START_RE = re.compile(r"^(\d{1,2})\s*[\.、．](?!\d)")
# 子条款/列举续行：2) 或 3)开头
SUBITEM_RE = re.compile(r"^\d+\)")


def clean(text: str) -> str:
    return str(text or "").replace("\t", " ").replace("\u3000", " ").strip()


def split_mark(text: str):
    """▲/★ 可能写在编号前（▲1、激光波长）也可能写在编号后（5.▲治疗手具）"""
    mark = ""
    for sym in ("★", "▲"):
        if text.startswith(sym):
            mark = sym
            text = text[len(sym):].strip()
            break
    return mark, text


def find_device_sections(paras: list) -> list:
    """返回 [{name, start, end}]，end 为下一节标题的下标（不含）"""
    heads = []
    for i, t in enumerate(paras):
        m = DEVICE_HEADER_RE.match(t)
        if m and len(t) <= 45:  # 标题行都很短，长的是正文段落
            heads.append({"name": clean(m.group(2)), "start": i})
    for j, h in enumerate(heads):
        h["end"] = heads[j + 1]["start"] if j + 1 < len(heads) else len(paras)
    return heads


def parse_section(paras: list, start: int, end: int) -> list:
    """小节内按条款编号切分并合并多行，返回 [{item_id, mark, requirement}]"""
    items = []
    cur = None
    for i in range(start + 1, end):
        t = clean(paras[i])
        if not t:
            continue
        m = CLAUSE_START_RE.match(t)
        if m and not SUBITEM_RE.match(t):
            # 新条款开始：编号后的文本里可能藏着 ▲/★
            body = t[m.end():].strip()
            mark, body = split_mark(body)
            cur = {"item_id": m.group(1), "mark": mark, "requirement": body}
            items.append(cur)
        elif cur is not None:
            # 续行：并入当前条款（子条款、列举、被排版拆开的半句）
            extra_mark, t2 = split_mark(t)  # 极少数情况 ▲ 单独出现在续行行首
            if extra_mark and not cur["mark"]:
                cur["mark"] = extra_mark
            cur["requirement"] = (cur["requirement"] + " " + t2).strip()
        # cur is None：标题后、第一条编号前的散句（如「1. 设备用途」前的引言），跳过
    return items


def parse_tender(path: str, keyword: str = "皮秒") -> dict:
    """主入口：整本招标 -> 目标设备条款列表。找不到设备小节时抛 ValueError。"""
    doc = Document(path)
    paras = [x.text.strip() for x in doc.paragraphs if x.text.strip()]

    sections = find_device_sections(paras)
    if not sections:
        raise ValueError("没找到「（一）设备名」格式的设备小节标题——"
                         "这份文件的排版可能不在预期内，需要人工看一眼")

    hits = [s for s in sections if keyword in s["name"]]
    if not hits:
        names = "、".join(s["name"] for s in sections)
        raise ValueError(f"设备小节里没有名字含「{keyword}」的。全部小节：{names}")
    target = hits[0]

    items = parse_section(paras, target["start"], target["end"])
    if len(items) < 3:
        raise ValueError(f"小节「{target['name']}」只切出 {len(items)} 条，切分逻辑可能失配")

    for it in items:
        it["row"] = it["item_id"]          # 沿用上游字段名，03/06/07 无需改动
        it["quantifiable"] = bool(re.search(r"\d", it["requirement"]))

    return {
        "source": Path(path).name,
        "device": target["name"],
        "count": len(items),
        "sections_found": [s["name"] for s in sections],
        "items": items,
    }


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    path = sys.argv[1]
    keyword = sys.argv[2] if len(sys.argv) >= 3 else "皮秒"

    result = parse_tender(path, keyword)
    print(f"文件：{result['source']}")
    print(f"设备小节：{result['device']}（共发现 {len(result['sections_found'])} 个小节）")
    print(f"切出条款：{result['count']} 条")
    print("-" * 68)
    for it in result["items"]:
        flag = it["mark"] or " "
        print(f"  [{flag}] {it['item_id']:<4} {it['requirement'][:56]}")

    OUT_FILE.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n已保存 -> {OUT_FILE}")


if __name__ == "__main__":
    main()
