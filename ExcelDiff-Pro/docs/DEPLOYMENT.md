# Deploying ExcelDiff Pro / 部署 ExcelDiff Pro

> This project is a deployable portfolio app, not a claim of independent penetration testing or enterprise security certification.
>
> 这是一个可部署的作品集 Web App，并不表示已经过第三方渗透测试或企业安全认证。

## English

### Option 1 — Local development

Requirements: Python 3.11–3.12, pip, 500 MB+ free memory for small workbooks (large files may need substantially more).

```bash
python -m venv .venv
# macOS / Linux:
source .venv/bin/activate
# Windows PowerShell:
# .venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m streamlit run app.py
```

Visit `http://localhost:8501`. The app offers a **Try sample** button so that you can try it without supplying files.

### Option 2 — Streamlit Community Cloud

1. Push the folder contents to a **GitHub** repository; include the hidden `.streamlit/config.toml` folder.
2. Open [share.streamlit.io](https://share.streamlit.io) and select **Create app** / **Deploy**. Choose your repository and branch.
3. Set **Main file path** to `app.py`. The platform reads `requirements.txt` automatically.
4. Deploy, then test sample files, column selection and the downloaded XLSX report.
5. Check application logs, available memory, uptime limits and Community Cloud terms before sharing it with clients.

**Important:** Treat free public hosting as a demo. Do not encourage strangers to upload payroll, banking statements, customer personal information or other sensitive datasets to a public instance. Use private infrastructure with access control where required.

Official documentation: https://docs.streamlit.io/deploy/streamlit-community-cloud and https://docs.streamlit.io/deploy/tutorials/docker

### Option 3 — Docker

```bash
docker build -t exceldiff-pro .
docker run --rm -p 8501:8501 exceldiff-pro
```

Open `http://localhost:8501`. Docker image runs as an unprivileged user and exposes the default Streamlit port. For an internet-facing server, place it behind an HTTPS reverse proxy with authentication, sensible request rate limits and monitoring. Review your organization's policies for handling uploaded files.

### Environment and resource controls

- Default maximum upload: **10 MB per file**, configured in `.streamlit/config.toml`.
- Maximum per imported worksheet: **30,000 physical rows**, **100 columns**.
- Max XLSX uncompressed ZIP contents: **120 MB** across at most **500 members**.
- Files are processed in **server session memory**; the system does not intentionally write input files to disk or a database.
- Report generation, memory and CPU can scale badly for large worksheets. Measure under load before promising high concurrency.
- No external AI API or embedded analytics are used; the UI theme does not fetch Google Fonts.

### Language, branding and enquiry link

- **English is the default.** Visitors can select 简体中文 using the sidebar language control.
- A sidebar panel and a section after the main workflow promote **custom Python / Excel / SQL automation**. They link to `greatfoodprefer@gmail.com` with a normal `mailto:` URL and do not record leads.
- To personalize your installation, edit `CONTACT_EMAIL` and the bilingual strings inside `render_contact_cta()` in `app.py`.
- The site CSS comes from `assets/theme.css`, which **must contain raw CSS only**. It is loaded through `st.html(Path(...))`; do not insert `<style>` or `</style>` tags in this file, or raw CSS may appear as page text.
- Confirm the contact button opens an email draft on your operating system. `mailto:` depends on visitors having a registered mail client; always show a copyable email address as a fallback.

### Pre-release checklist

- [ ] Install dependencies on the actual host and start the Streamlit app.
- [ ] Confirm the top of the page contains **no visible raw CSS text**; English should be selected by default.
- [ ] Verify the bilingual Work With Me section and email action (mailto).
- [ ] Click Try sample and inspect all result tabs.
- [ ] Select A: B/E, B: C/F; confirm two matches, one mismatch, duplicate/empty/invalid rows.
- [ ] Download report and open it in Excel/LibreOffice.
- [ ] Test CSV, XLSX and an invalid upload.
- [ ] Run `python -m pytest -q`.
- [ ] Configure HTTPS, authentication (if needed), rate limits and backups for deployment configuration.
- [ ] Document data retention with the chosen hosting provider.

## 中文（简体）

### 方式一：电脑本地运行

安装 Python 3.11–3.12，进入项目目录：

```bash
python -m venv .venv
# Windows PowerShell:
# .venv\Scripts\Activate.ps1
# macOS / Linux:
source .venv/bin/activate
pip install -r requirements.txt
python -m streamlit run app.py
```

浏览器打开 `http://localhost:8501`，点击 **试用示例** 就能使用两份内建 Excel 完整测试流程。

### 方式二：Streamlit Community Cloud

1. 把整个项目 Push 到自己的 GitHub Repository，保留隐藏的 `.streamlit/config.toml`。
2. 打开 [share.streamlit.io](https://share.streamlit.io)，选择部署现有 GitHub 项目。
3. Main file path 填写 `app.py`，平台会根据 `requirements.txt` 安装依赖。
4. 部署后打开网页，运行 Sample、进行 Column Mapping，并下载报告确认正常。
5. 注意免费平台的内存限制、公开访问属性和托管服务条款。

**注意：** 公开的免费部署建议只用作 GitHub Portfolio Demo。若处理薪资、客户资料、财务报表等敏感文件，建议自行部署到内部服务器，并额外加入认证与 HTTPS。

### 方式三：Docker

```bash
docker build -t exceldiff-pro .
docker run --rm -p 8501:8501 exceldiff-pro
```

访问 `http://localhost:8501`。如果要正式对公网开放，建议放在 HTTPS Reverse Proxy 后面，并部署登录认证、访问速率限制、资源监控和错误日志策略。

### 语言、推广和邮箱设置

- 默认显示 **English**，用户可以在 Sidebar 切换简体中文。
- Sidebar 与工作流程底部会显示定制开发介绍，包含 Excel 专属规则、Python / SQL 自动化及 API 整合服务。
- 邮箱是 `greatfoodprefer@gmail.com`，由 `app.py` 的 `CONTACT_EMAIL` 设置。按钮使用普通 `mailto:`，不会收集或储存访客资料；如果对方没有预设邮件软件，可直接复制屏幕上的邮箱地址。
- `assets/theme.css` **只能放纯 CSS**，不要再添加 `<style>` 或 `</style>`；Streamlit 会自动载入样式文件，避免顶部出现原始 CSS 文字。

### 上线前必查

- [ ] 在实际部署环境安装依赖并启动网页。
- [ ] 检查英文为默认语言、顶部没有多余 CSS 文字、定制联系按钮可用。
- [ ] 验证上传、Sheet 切换、Column Mapping、金额比较、文字比较和仅检查存在模式。
- [ ] 测试空 Key、重复 Key、无效金额以及多笔付款汇总功能。
- [ ] 核对生成的 `OVERVIEW`、`ALL RESULTS`、`EXCEPTIONS`、`SETTINGS`。
- [ ] 运行 `python -m pytest -q`。
- [ ] 确认服务器的内存限制及文件处理流程。
- [ ] 按实际用户群决定是否需要账号、HTTPS 与访问控制。
- [ ] 向用户明确解释所选主机对暂存文件与日志的政策。
