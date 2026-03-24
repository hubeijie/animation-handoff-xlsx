# 美术统筹表工坊 — 公开访问（GitHub + Render）

本仓库是 **FastAPI**，需要能跑 Docker 的平台托管，**不能**只靠 GitHub Pages。

## 1. 代码推到 GitHub

```bash
cd /path/to/animation-handoff-web
git remote add origin https://github.com/<你的用户名>/<仓库名>.git   # 若已添加则跳过
git push -u origin main
```

## 2. 在 Render 上部署

1. 打开 [render.com](https://render.com)，用 GitHub 登录并授权仓库。
2. **New → Blueprint**（或 **Web Service** 也可）：
   - Blueprint：选中本仓库，识别根目录 `render.yaml` 后一路确认。
   - 若用 **Web Service**：**Runtime** 选 **Docker**，Root Directory 留空，保存。
3. 等待首次构建与启动（免费档冷启动约 1 分钟）。
4. 控制台会给出公网地址，形如 `https://art-handoff-workshop.onrender.com`，即为「大家都能打开的工坊」。

## 3. 环境变量（按需）

| 变量 | 何时需要 |
|------|----------|
| `OPENAI_API_KEY` | 用户在网页上勾选「改由**服务端**调用模型」时必填 |
| `OPENAI_BASE_URL` | 非默认 OpenAI 端点时设置 |
| `OPENAI_MODEL` | 可选，默认见 `app/main.py` |
| `DRY_RUN=1` | 仅调试：不调模型，用固定 JSON 测生成 xlsx（**勿长期开**） |

浏览器内填 Key 直连模型时，**不必**在 Render 配 `OPENAI_API_KEY`。

## 4. 说明免费档

- 免费 Web Service 一段时间无访问会休眠，**第一次打开可能较慢**。
- 私有 GitHub 仓库在 Render 上通常**不能**用 `free` 计划，需改 `render.yaml` 里的 `plan` 或在控制台选付费档，或将仓库设为 Public。

---

## 5. GitHub Pages + 后端「拆站」（他人可打开页面，导出走后端）

适合：**不想给 Render 绑卡**或希望访问入口在 `github.io`。

1. 按上文把 **FastAPI** 部署到任意公网地址，记为 `https://你的后端`（须为启用 CORS 的当前版本，见 `app/main.py`）。
2. 本仓库 **Settings → Pages**：Branch **`main`**，目录 **`/docs`**。
3. 访客打开：`https://<用户>.github.io/<仓库名>/`（会自动进 **`workshop.html`**）。
4. 在工坊左侧 **「工坊 API 根地址」** 填入 `https://你的后端`（勿带末尾 `/`），或分享带参链接：  
   `https://<用户>.github.io/<仓库名>/workshop.html?api=https://你的后端`

说明：静态页与 API **域名不同**时依赖浏览器跨域；当前后端默认 `Access-Control-Allow-Origin: *`。
