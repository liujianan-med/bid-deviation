# -*- coding: utf-8 -*-
"""
check_key.py —— 检查 .env 里的 API Key 是否可用

它只做三件事：
    1. 看 .env 有没有被正确加载
    2. 检查 Key 格式对不对（不会打印完整内容）
    3. 发一个最小请求，验证 Key 真的能通

运行：
    .venv\\Scripts\\python.exe scripts\\check_key.py

为什么要单独做这个检查？
因为「Key 填错」和「网络不通」和「余额不足」在正式脚本里
都会表现为同一句"调用失败"，新手根本分不清。先单独验一次，省很多时间。
"""

import os
import sys

import requests
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENV_FILE = os.path.join(BASE_DIR, ".env")

print("=" * 60)
print("API Key 自检")
print("=" * 60)

# ---------- 第 1 步：文件是否存在 ----------
if not os.path.exists(ENV_FILE):
    print("[X] 找不到 .env 文件")
    print("    执行：copy .env.example .env")
    sys.exit(1)
print(f"[OK] 找到配置文件 {ENV_FILE}")

# ---------- 第 2 步：加载 ----------
# override=True 让它覆盖系统环境变量里可能存在的同名值
load_dotenv(ENV_FILE, override=True)

key = (os.getenv("LLM_API_KEY") or "").strip()
base = (os.getenv("LLM_API_BASE") or "").strip().rstrip("/")
model = (os.getenv("LLM_MODEL") or "").strip()

# ---------- 第 3 步：格式检查 ----------
if not key:
    print("[X] LLM_API_KEY 是空的")
    print("    用记事本打开 .env，在 LLM_API_KEY= 后面粘贴你的 Key")
    sys.exit(1)

# 只显示头尾，中间打码。养成习惯：日志里永远不要出现完整密钥
masked = key[:6] + "*" * 8 + key[-4:] if len(key) > 12 else "*" * len(key)
print(f"[OK] Key 已加载：{masked}（共 {len(key)} 字符）")
print(f"     接口：{base}")
print(f"     模型：{model}")

problems = []
if not key.startswith("sk-"):
    problems.append("Key 不是 sk- 开头")
if " " in key or "\t" in key:
    problems.append("Key 里含有空格（粘贴时多带了字符）")
if key.startswith('"') or key.startswith("'"):
    problems.append("Key 被引号包住了，去掉引号")
if "..." in key or "*" in key:
    problems.append("这是打码版本，不是完整 Key")
if key.count("-") < 1 or len(key) < 20:
    problems.append(f"长度只有 {len(key)}，正常 Key 一般 30 字符以上")

if problems:
    print()
    print("[X] 格式有问题：")
    for p in problems:
        print("    - " + p)
    sys.exit(1)

print("[OK] 格式检查通过")

# ---------- 第 4 步：真实请求 ----------
print()
print("正在向接口发起测试请求……")

url = f"{base}/chat/completions"
headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
body = {
    "model": model,
    "messages": [{"role": "user", "content": "回复两个字：正常"}],
    "max_tokens": 8,
    "temperature": 0,
}

try:
    resp = requests.post(url, headers=headers, json=body, timeout=30)
except requests.Timeout:
    print("[X] 请求超时（30 秒）")
    print("    网络连不上接口。试试换用通义或智谱的接口地址")
    sys.exit(1)
except requests.RequestException as e:
    print(f"[X] 网络错误：{type(e).__name__}")
    print(f"    {e}")
    sys.exit(1)

if resp.status_code == 200:
    data = resp.json()
    reply = data["choices"][0]["message"]["content"].strip()
    usage = data.get("usage", {})
    print("[OK] 调用成功！")
    print(f"     模型回复：{reply}")
    print(f"     Token 消耗：{usage.get('total_tokens', '?')}")
    print()
    print("=" * 60)
    print("Key 完全可用。接下来运行：")
    print("    .venv\\Scripts\\python.exe scripts\\02_call_llm.py")
    print("=" * 60)
    sys.exit(0)

# ---------- 错误对号入座 ----------
hints = {
    401: "Key 无效或已失效 —— 回平台重新创建一个",
    402: "账户余额不足 —— DeepSeek 新号一般有赠送额度，去平台看余额",
    403: "没有权限 —— 检查 Key 是否属于这个接口地址",
    404: "接口地址不对 —— 检查 .env 里的 LLM_API_BASE",
    422: "参数格式错误 —— 通常是模型名写错了，检查 LLM_MODEL",
    429: "请求太频繁 —— 等一分钟再试",
    500: "服务端故障 —— 稍后重试",
    503: "服务繁忙 —— 稍后重试",
}

print(f"[X] 请求失败，HTTP {resp.status_code}")
print(f"    {hints.get(resp.status_code, '未知错误')}")
print()
print("响应内容（前 300 字符）：")
print(resp.text[:300])
sys.exit(1)
