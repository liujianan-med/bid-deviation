# 招标参数偏离表生成器 —— 部署镜像
# 本机（Win10/8G 内存）跑不了 Docker Desktop，这个文件供日后部署到
# 公司服务器 / 云主机使用（那些机器上 docker build -t bid-deviation . 即可）。

FROM python:3.11-slim

WORKDIR /srv/app

# 先装依赖再拷代码：依赖不变时重新构建能走缓存，快得多
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# 代码与运行时必需的数据文件（产品参数库、证明材料索引库）
COPY scripts/ scripts/
COPY app/ app/
COPY data/picov_params.json data/evidence_map.json data/

# 真实标书、密钥、输出产物一律不进镜像（安全红线）
# .dockerignore 里做了硬排除，这里是双保险
RUN rm -rf data/uploads

EXPOSE 8000

# .env 不进镜像，密钥通过运行时环境变量传入：
#   docker run -d -p 8000:8000 \
#     -e LLM_API_BASE=https://api.deepseek.com/v1 \
#     -e LLM_API_KEY=sk-xxxx \
#     -e LLM_MODEL=deepseek-chat \
#     bid-deviation
CMD ["python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
