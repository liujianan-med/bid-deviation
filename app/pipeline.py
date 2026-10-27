# -*- coding: utf-8 -*-
"""
app/pipeline.py —— 完整流水线：Excel 上传 -> 抽取 -> 三态判定 -> Word

把 scripts/ 里已验证的逻辑串起来，供 API 层调用：
    1. parse_excel()          读招标参数表（复用 01 的解析规则）
    2. extract_with_retry()   带重试的 LLM 抽取（复用 02 的 extract_one）
    3. run_compare()          三态判定（06 的核心函数，重构后暴露）
    4. build_document()       生成 Word（07 的核心函数，重构后暴露）

进度通过 on_progress 回调抛给上层（API 层用它更新任务状态）。
"""

import importlib
import json
import os
import re
import sys
import time

import openpyxl

# scripts/ 里的模块文件名以数字开头，必须用 importlib 按名字导入
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(PROJECT_ROOT, "scripts")
for p in (PROJECT_ROOT, SCRIPTS_DIR):
    if p not in sys.path:
        sys.path.insert(0, p)

from dotenv import load_dotenv
load_dotenv(os.path.join(PROJECT_ROOT, ".env"))

m01 = importlib.import_module("01_read_params")
m02 = importlib.import_module("02_call_llm")
m06 = importlib.import_module("06_compare")
m07 = importlib.import_module("07_make_docx")
m08 = importlib.import_module("08_parse_tender_docx")

# 与 01 相同的模板结构约定；换新模板结构时这里要跟着改
HEADER_ROW = m01.HEADER_ROW
FIRST_DATA_ROW = m01.FIRST_DATA_ROW

SLEEP_BETWEEN = 0.5
MAX_RETRY = 3


def parse_excel(path: str) -> dict:
    """读招标参数 Excel -> {items: [{row, item_id, mark, requirement}...]}

    解析规则与 01_read_params.py 完全一致（★/▲ 拆分、空行跳过、章节标题跳过），
    这里不复用它的 main() 是因为 main 还带统计打印和写盘，职责不同。
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Sheet1" not in wb.sheetnames:
        raise ValueError(f"Excel 里没有 Sheet1（现有：{wb.sheetnames}），请按模板格式提供")
    ws = wb["Sheet1"]

    items = []
    for row in range(FIRST_DATA_ROW, ws.max_row + 1):
        raw_id = m01.clean(ws.cell(row, 1).value)
        raw_req = m01.clean(ws.cell(row, 2).value)
        if not raw_id and not raw_req:
            continue
        mark, item_id = m01.split_mark(raw_id)
        if raw_req == "":
            continue  # 章节标题行
        if not mark:
            mark, raw_req = m01.split_mark(raw_req)
        items.append({"row": row, "item_id": item_id, "mark": mark,
                      "requirement": raw_req,
                      "quantifiable": m01.has_number(raw_req)})

    if len(items) < 3:
        raise ValueError(f"只解析出 {len(items)} 条参数，远少于正常标书的几十条——"
                         f"请确认这是「第3行表头、第5行起数据」的标准参数表模板")
    return {"count": len(items), "items": items}


def extract_with_retry(text: str, on_progress=None) -> dict:
    """带重试的抽取。只对值得重试的错误重试：429 限速、超时、5xx。"""
    last_err = None
    for attempt in range(1, MAX_RETRY + 1):
        try:
            return m02.extract_one(text)
        except RuntimeError as e:
            msg = str(e)
            if "HTTP 429" in msg or "HTTP 5" in msg:
                last_err = e
                wait = 2 * attempt
                if on_progress:
                    on_progress(f"重试 {attempt}/{MAX_RETRY}：{msg[:50]}，等 {wait}s")
                time.sleep(wait)
            else:
                raise  # 401/402/400 重试无意义，直接抛
        except Exception as e:
            last_err = e
            wait = 2 * attempt
            if on_progress:
                on_progress(f"重试 {attempt}/{MAX_RETRY}：{type(e).__name__}，等 {wait}s")
            time.sleep(wait)
    raise last_err


def run_pipeline(excel_path: str, out_docx_path: str, on_progress=None,
                 device_keyword: str = "皮秒") -> dict:
    """完整流水线。返回 {report, docx_path}；异常向上抛由 API 层捕获。

    输入两种都认：
      .xlsx —— 标准参数表模板（01 的规则）
      .docx —— 整本招标文件 Word，自动定位目标设备小节（08 的规则）
    """
    log = on_progress or (lambda msg: None)

    # ---- 1. 读参数 ----
    lower = excel_path.lower()
    if lower.endswith(".docx"):
        log(f"识别为整本招标文件 Word，定位「{device_keyword}」设备小节…")
        params = m08.parse_tender(excel_path, device_keyword)
        log(f"设备小节「{params['device']}」切出 {params['count']} 条条款")
    else:
        log("读取招标参数表…")
        params = parse_excel(excel_path)
        log(f"解析出 {params['count']} 条参数")

    # ---- 2. 批量抽取（LLM，占全程 95% 时间）----
    results = []
    failed = []
    t0 = time.time()
    for i, item in enumerate(params["items"], 1):
        log(f"[{i}/{params['count']}] 行{item['row']} {item['mark']} {item['requirement'][:30]}")
        try:
            parsed = extract_with_retry(item["requirement"], on_progress=log)
            results.append({"row": item["row"], "item_id": item["item_id"],
                            "mark": item["mark"], "source_text": item["requirement"],
                            "parsed": parsed})
        except Exception as e:
            failed.append({"row": item["row"], "item_id": item["item_id"],
                           "error": str(e)[:150]})
            log(f"   [失败] {str(e)[:60]}")
        time.sleep(SLEEP_BETWEEN)
    log(f"抽取完成：成功 {len(results)}，失败 {len(failed)}，耗时 {time.time()-t0:.0f}s")

    if not results:
        raise RuntimeError("全部条目抽取失败（检查 API Key / 余额 / 网络）")

    # ---- 3. 三态判定 ----
    # 品类由 device_keyword 解析（如「皮秒」「光子」），产品库与证据库都按品类加载
    from product_db import load as load_product
    prod = load_product(device_keyword)
    product = prod["params"]
    m07.load_evidence(prod["evidence_path"])
    log(f"品类：{prod['name']}（产品参数 {len(product)} 项）")
    log("三态判定中…")
    report = m06.run_compare(results, product)
    c = report["counts"]
    log(f"判定完成：响应 {c['响应']} / 正偏离 {c['正偏离']} / "
        f"负偏离 {c['负偏离']} / 存疑 {c['存疑']}")

    # ---- 4. 生成 Word ----
    log("生成 Word 偏离表…")
    doc = m07.build_document(report, product)
    doc.save(out_docx_path)
    log(f"已生成 {os.path.basename(out_docx_path)}")

    # 失败条目也记进报告，前端能看到哪几行没跑出来
    report["failed"] = failed
    report["source"] = os.path.basename(excel_path)
    return {"report": report, "docx_path": out_docx_path}
