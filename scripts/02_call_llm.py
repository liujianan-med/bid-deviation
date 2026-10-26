# -*- coding: utf-8 -*-
"""
02_call_llm.py —— 第一次调用大模型 API，把一条模糊的技术要求变成结构化数据

为什么第一步就做这个？
因为 FDE 日常 80% 的活就是这个形状：
    非结构化输入 -> 调模型 -> 结构化输出 -> 校验

运行前准备：
    1. 复制 .env.example 为 .env
    2. 在 .env 里填上你的 API Key

运行：
    .venv\\Scripts\\python.exe scripts\\02_call_llm.py
    .venv\\Scripts\\python.exe scripts\\02_call_llm.py "▲设备重量≥130KG"
"""

import json
import os
import sys

import requests
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("LLM_API_KEY", "").strip()
API_BASE = os.getenv("LLM_API_BASE", "https://api.deepseek.com/v1").rstrip("/")
MODEL = os.getenv("LLM_MODEL", "deepseek-chat")

# 系统提示词：告诉模型它的角色和输出格式
SYSTEM_PROMPT = """你是一个医疗设备招标文件解析助手。
用户会给你一条技术参数要求（可能带 ★ 或 ▲ 标记），你要把它拆成结构化字段。

字段定义：
- param_name   ：参数名称，如「峰值功率」
- operator     ：比较关系，只能是 >, >=, =, <=, <, range, none 之一
                  "≥" -> ">="  "≧" -> ">="  "≤" -> "<="  ">" -> ">"  "＜" -> "<"
                  "2-10mm" 这种区间用 "range"
- value         ：数值部分，保留原文数字，如 "1.7"；区间写成 "2-10"
- unit          ：单位，如 "GW", "ps", "mm", "J/cm2", "KG"；没有就留空
- is_measurable ：布尔值，这条要求能不能用数值比对（"进口七关节导光臂"这种结构描述为 false）

只输出 JSON，不要任何解释文字，不要 markdown 代码块。"""


def extract_one(requirement: str) -> dict:
    """把一条技术要求发给模型，拿回结构化结果。"""

    url = f"{API_BASE}/chat/completions"
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }
    body = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": requirement},
        ],
        # temperature 越低，输出越稳定。抽取任务应该接近 0
        "temperature": 0,
        # 强制模型返回 JSON 对象（DeepSeek / 通义 / 智谱 都支持这个参数）
        "response_format": {"type": "json_object"},
    }

    # timeout 必须设。没有超时的网络请求是生产事故的常见来源
    resp = requests.post(url, headers=headers, json=body, timeout=30)

    # 不同厂商对 response_format 的支持不一致。智谱部分模型、硅基流动的
    # 开源模型会直接返回 400。这里做一次降级重试：去掉这个参数再发一次，
    # 然后在提示词里已经有"只输出 JSON"的约束，多数情况仍然能拿到合法 JSON。
    # 这就是 FDE 常做的事——同一份代码要跑在多家客户的模型上。
    if resp.status_code == 400 and "response_format" in resp.text.lower():
        print("    [降级] 当前模型不支持 response_format，去掉后重试")
        body.pop("response_format", None)
        resp = requests.post(url, headers=headers, json=body, timeout=30)

    # 这一行很关键：状态码不是 200 时，直接抛出带响应内容的错误
    # 否则你会拿到一个 KeyError，完全看不出真实原因
    if resp.status_code != 200:
        raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:300]}")

    data = resp.json()
    content = data["choices"][0]["message"]["content"]

    # 模型有时会在 JSON 外面套一层 ```json ```，这里兜一下
    content = content.strip()
    if content.startswith("```"):
        content = content.split("\n", 1)[1] if "\n" in content else content
        content = content.rsplit("```", 1)[0]

    return json.loads(content.strip())


def main():
    if not API_KEY:
        print("[未配置 API Key]")
        print()
        print("请先做两步：")
        print("  1. 复制 .env.example 为 .env")
        print("        copy .env.example .env")
        print("  2. 用记事本打开 .env，把 LLM_API_KEY 换成你的真实 Key")
        print()
        print("去哪儿拿 Key（三家都有免费额度，选一个即可）：")
        print("  DeepSeek  https://platform.deepseek.com/api_keys")
        print("  通义千问  https://bailian.console.aliyun.com/")
        print("  智谱 GLM  https://open.bigmodel.cn/usercenter/apikeys")
        return

    # 命令行传了参数就用参数，没传就用默认样例
    if len(sys.argv) > 1:
        samples = [" ".join(sys.argv[1:])]
    else:
        samples = [
            "★峰值功率≧1.7GW（第三方检测报告）",
            "▲光斑直径可调：2-10mm，误差≤0.2mm",
            "传输方式：进口七关节重锤式导光臂",
        ]

    print(f"接口 {API_BASE}   模型 {MODEL}")
    print("=" * 72)

    results = []
    for i, text in enumerate(samples, 1):
        print(f"\n[{i}] 原文：{text}")
        try:
            out = extract_one(text)
            print("    结构化：", json.dumps(out, ensure_ascii=False))
            results.append({"source_text": text, "parsed": out})
        except json.JSONDecodeError as e:
            print(f"    [解析失败] 模型返回的不是合法 JSON：{e}")
        except requests.Timeout:
            print("    [超时] 30 秒没响应，检查网络或换接口地址")
        except Exception as e:
            print(f"    [失败] {type(e).__name__}: {e}")

    print("\n" + "=" * 72)
    print(f"成功 {len(results)} / {len(samples)} 条")

    out_file = os.path.join(os.path.dirname(__file__), "..", "data", "llm_extract.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"已保存 -> data/llm_extract.json")


if __name__ == "__main__":
    main()
