# -*- coding: utf-8 -*-
"""
app/main.py —— FastAPI 服务：偏离表生成器的对外接口 + 网页界面

设计要点：
    1. 一份 Excel 的完整流水线要跑 1-2 分钟（31 条 × LLM 调用），
       HTTP 请求不能同步等——用「提交任务 + 轮询进度」模式：
           POST /api/jobs        上传 Excel，立刻返回任务号
           GET  /api/jobs/{id}   轮询进度和结果
           GET  /api/jobs/{id}/docx   下载生成的 Word
    2. 任务状态存内存（单机单人用足够；多用户/重启持久化是以后的事）
    3. 密钥只从 .env / 环境变量读，绝不写进代码或日志

启动：
    .venv\\Scripts\\python.exe -m uvicorn app.main:app --host 127.0.0.1 --port 8000
然后浏览器打开 http://127.0.0.1:8000
"""

import os
import re
import secrets
import threading
import time
from datetime import date
from uuid import uuid4

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))

from app import pipeline

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UPLOAD_DIR = os.path.join(BASE_DIR, "data", "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = FastAPI(title="招标参数偏离表生成器", version="0.5")

# ---------- 内存任务表 ----------
_jobs = {}
_lock = threading.Lock()
_JOB_PROGRESS_RE = re.compile(r"^\[(\d+)/(\d+)\]")  # 匹配 "[3/31] ..." 进度消息


def _set_job(job_id: str, **kw):
    with _lock:
        _jobs[job_id].update(kw)


def _run_job(job_id: str, excel_path: str, out_path: str, device_keyword: str = "皮秒"):
    """后台线程：跑流水线，把进度写回任务表"""
    def on_progress(msg: str):
        with _lock:
            j = _jobs[job_id]
            j["log"].append({"t": time.strftime("%H:%M:%S"), "msg": msg})
            m = _JOB_PROGRESS_RE.match(msg)
            if m:  # "[3/31] ..." -> 更新进度条
                j["done"], j["total"] = int(m.group(1)), int(m.group(2))
    try:
        result = pipeline.run_pipeline(excel_path, out_path, on_progress=on_progress,
                                       device_keyword=device_keyword)
        with _lock:
            j = _jobs[job_id]
            j["status"] = "done"
            j["result"] = result["report"]
            j["docx"] = result["docx_path"]
            j["log"].append({"t": time.strftime("%H:%M:%S"), "msg": "全部完成"})
    except Exception as e:
        with _lock:
            j = _jobs[job_id]
            j["status"] = "error"
            j["error"] = f"{type(e).__name__}: {e}"
            j["log"].append({"t": time.strftime("%H:%M:%S"), "msg": f"出错：{e}"})
    finally:
        try:
            os.remove(excel_path)  # 上传的临时文件用完即删
        except OSError:
            pass


# ---------- 接口 ----------

@app.get("/api/health")
def health():
    return {"ok": True, "model": os.getenv("LLM_MODEL", ""),
            "api_base": os.getenv("LLM_API_BASE", "")}


@app.post("/api/jobs")
async def create_job(file: UploadFile = File(...), device_keyword: str = "皮秒"):
    """上传招标参数 Excel 或整本招标文件 Word，创建后台任务"""
    name = os.path.basename(file.filename or "")
    if not name.lower().endswith((".xlsx", ".docx")):
        raise HTTPException(400, "只接受 .xlsx（参数表模板）或 .docx（整本招标文件）")

    job_id = uuid4().hex[:12]
    excel_path = os.path.join(UPLOAD_DIR, f"{job_id}_{secrets.token_hex(4)}_{name}")
    content = await file.read()
    if len(content) > 20 * 1024 * 1024:
        raise HTTPException(400, "文件超过 20MB，疑似传错文件")
    with open(excel_path, "wb") as f:
        f.write(content)

    out_path = os.path.join(OUTPUT_DIR, f"偏离表-{os.path.splitext(name)[0]}-{date.today():%Y%m%d}-{job_id[:6]}.docx")

    with _lock:
        _jobs[job_id] = {"status": "running", "log": [], "done": 0, "total": 0,
                         "created": time.strftime("%H:%M:%S"),
                         "filename": name, "error": None, "result": None, "docx": None}

    threading.Thread(target=_run_job, args=(job_id, excel_path, out_path, device_keyword),
                     daemon=True).start()
    return {"job_id": job_id}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str):
    with _lock:
        j = _jobs.get(job_id)
        if not j:
            raise HTTPException(404, "任务不存在")
        return dict(j)


@app.get("/api/jobs/{job_id}/docx")
def job_docx(job_id: str):
    with _lock:
        j = _jobs.get(job_id)
        if not j:
            raise HTTPException(404, "任务不存在")
        if j["status"] != "done":
            raise HTTPException(409, "任务尚未完成")
        path = j["docx"]
    if not path or not os.path.exists(path):
        raise HTTPException(410, "生成的文件已不存在")
    return FileResponse(path, filename=os.path.basename(path),
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


class ExtractReq(BaseModel):
    text: str


@app.post("/api/extract")
def extract_one_text(req: ExtractReq):
    """单条参数快速抽取（调试用）：给一句招标要求，返回结构化结果"""
    try:
        return {"ok": True, "parsed": pipeline.m02.extract_one(req.text)}
    except Exception as e:
        return JSONResponse(status_code=502, content={"ok": False, "error": str(e)[:200]})


@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))
