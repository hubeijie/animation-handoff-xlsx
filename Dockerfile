# 美术统筹表工坊 — 供 Render / Fly.io / Railway 等从 GitHub 构建
FROM python:3.12-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY static ./static
COPY scripts/generate_art_coordination_xlsx.py ./scripts/

ENV PYTHONUNBUFFERED=1
# 显式指向镜像内脚本（可覆盖）
ENV GENERATOR_SCRIPT=/app/scripts/generate_art_coordination_xlsx.py

EXPOSE 8080
CMD uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8080}
