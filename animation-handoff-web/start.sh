#!/usr/bin/env bash
# 美术统筹表工坊 — 在本目录启动 FastAPI，勿对技能包目录使用 python -m http.server。
cd "$(dirname "$0")"
exec uvicorn app.main:app --host 127.0.0.1 --port 8765
