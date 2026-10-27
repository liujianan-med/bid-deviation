# -*- coding: utf-8 -*-
"""
03_batch_extract.py —— 批量处理全部 31 条招标参数

和 02 的区别：
    02 是「跑通一次」，这里是「跑稳一批」。批量场景多出三个工程问题：
    1. 限速：连续请求太快会被 429 拒绝，每条之间 sleep
    2. 重试：网络抖动是常态，单条失败不能让整批崩掉
    3. 记账：记录每次调用消耗的 token，跑批前先知道这批花多少钱

运行：
    .venv\\Scripts\\python.exe scripts\\03_batch_extract.py
"""

import json
import os
import sys
import time

from dotenv import load_dotenv

# 直接复用 02 里已经调通的函数——代码复用不是复制粘贴，是 import
from importlib import import_module
m02 = import_module("02_call_llm")  # 文件名以数字开头，只能这样导入
extract_one = m02.extract_one

load_dotenv()

# 不带参数：跑皮秒 31 条（默认行为，和之前完全一样）
# 带参数  ：python 03_batch_extract.py <输入json> <输出json>
#          例：python 03_batch_extract.py data/params_photon.json data/batch_extract_photon.json
if len(sys.argv) >= 3:
    PARAMS_FILE = os.path.join(os.path.dirname(__file__), "..", sys.argv[1])
    OUT_FILE = os.path.join(os.path.dirname(__file__), "..", sys.argv[2])
    TAG = "[光子]"
else:
    PARAMS_FILE = os.path.join(os.path.dirname(__file__), "..", "data", "params.json")
    OUT_FILE = os.path.join(os.path.dirname(__file__), "..", "data", "batch_extract.json")
    TAG = ""

SLEEP_BETWEEN = 0.5   # 每条之间停 0.5 秒，避免触发限速
MAX_RETRY = 3         # 单条最多重试 3 次


def extract_with_retry(text: str) -> dict:
    """带重试的抽取。只对「值得重试」的错误重试：429 限速、超时、5xx。"""
    last_err = None
    for attempt in range(1, MAX_RETRY + 1):
        try:
            return extract_one(text)
        except RuntimeError as e:
            msg = str(e)
            # 4xx 里只有 429（太快）值得重试，401/402 重试也没用
            if "HTTP 429" in msg or "HTTP 5" in msg or "Timeout" in type(e).__name__:
                last_err = e
                wait = 2 * attempt  # 第 1 次等 2 秒，第 2 次等 4 秒……这叫指数退避
                print(f"      [重试 {attempt}/{MAX_RETRY}] {msg[:60]}  等 {wait}s")
                time.sleep(wait)
            else:
                raise  # 401/402/400 这类重试无意义，直接抛
        except Exception as e:  # 超时、连接错误等
            last_err = e
            wait = 2 * attempt
            print(f"      [重试 {attempt}/{MAX_RETRY}] {type(e).__name__}: {str(e)[:60]}  等 {wait}s")
            time.sleep(wait)
    raise last_err


def main():
    with open(PARAMS_FILE, encoding="utf-8") as f:
        params = json.load(f)

    items = params["items"]
    print(f"{TAG}共 {len(items)} 条参数，接口 {os.getenv('LLM_API_BASE')}  模型 {os.getenv('LLM_MODEL')}")
    print("=" * 72)

    results = []
    failed = []
    t0 = time.time()

    for i, item in enumerate(items, 1):
        text = item["requirement"]
        row = item["row"]
        mark = item.get("mark") or " "
        print(f"[{i:>2}/{len(items)}] 行{row} {mark} {text[:46]}")

        try:
            parsed = extract_with_retry(text)
            results.append({"row": row, "item_id": item["item_id"], "mark": mark,
                            "source_text": text, "parsed": parsed})
            print(f"      OK  {parsed.get('param_name','?')} | {parsed.get('operator','?')} "
                  f"{parsed.get('value','')} {parsed.get('unit','')}")
        except Exception as e:
            failed.append({"row": row, "item_id": item["item_id"], "error": str(e)[:120]})
            print(f"      [失败] {str(e)[:80]}")

        time.sleep(SLEEP_BETWEEN)

    elapsed = time.time() - t0
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print("\n" + "=" * 72)
    print(f"完成：成功 {len(results)} / {len(items)}，失败 {len(failed)}，耗时 {elapsed:.0f} 秒")
    if failed:
        print("失败清单：")
        for f_ in failed:
            print(f"  行{f_['row']} {f_['item_id']}: {f_['error']}")
    print(f"结果已保存 -> data/batch_extract.json")


if __name__ == "__main__":
    main()
