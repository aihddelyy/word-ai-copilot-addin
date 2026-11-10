# Word AI 智能创作助手 (Word AI Copilot Add-in) - 架构与历次开发记录

> **文档用途**：本文档旨在供后续开发者或接入其他 AI 开发工具（如 Cursor、Claude Code、Windsurf、GitHub Copilot 等）时，能够快速、全面、精准地理解本项目的架构设计、技术演进脉络、各模块实现机制以及已知避坑指南，以便无缝进行二次开发与功能扩展。

---

## 1. 项目概览与技术选型

### 1.1 项目定位
本项目是深度嵌入 **Microsoft Word**（桌面端 Windows/Mac 以及 Office 网页版）的现代化智能创作辅助插件。旨在打通大语言模型（LLM）与 Word 文档之间的双向交互，实现选中文本的即时润色、扩写、精简、纠错、中英翻译，支持多版本 A/B 候选、整篇长文档审阅分析，并能将生成的 Markdown、表格与分级标题以 **Word 原生富文本格式** 直接写回文档。

### 1.2 关键技术栈
- **宿主扩展规范**：Microsoft Office Web Add-in 规范（基于 XML 清单 + Taskpane Web 视图）。
- **Office 交互 SDK**：`Office.js`（Word.js API 1.3+ 规范），基于异步上下文代理机制（`Word.run` / `context.sync()`）。
- **前端架构**：原生现代 JavaScript（ES6+ 无编译打包依赖），HTML5，CSS3（严格遵循微软 Fluent Design 2.0 视觉规范）。
- **Markdown 渲染**：`marked.js`（CDN 引入，负责将大模型流式输出转换为可视化 HTML 及 Word 原生兼容 DOM）。
- **本地服务与调试工具**：Python 3 内置标准库（实现自签名 SSL 证书自动生成、HTTPS 静态文件托管以及支持 SSE 的跨协议反向代理中继）。

---

## 2. 历次开发过程与功能演进记录 (Changelog)

### 【Phase 1】基础骨架与核心文档交互 (v1.0)
* **目标**：搭建符合微软官方规范的插件骨架，打通从 Word 读取选区到调用 LLM 并写回的基本闭环。
* **主要工作**：
  1. **清单文件设计 (`manifest.xml`)**：
     - 定义唯一插件 UUID（`d3b07384-d113-4bf8-9ae7-5b6c31e9c8a1`）。
     - 在 Word 顶部主页功能区（`TabHome`）注册专用的 Ribbon 按钮组【AI 智能创作】，设置 16px、32px、80px 图标。
     - 申请 `ReadWriteDocument` 权限，指定默认侧边栏入口为 `https://localhost:3000/src/taskpane.html`。
  2. **选区感知与监听**：
     - 利用 `Office.onReady()` 检测宿主环境（同时容错普通浏览器预览模式）。
     - 监听 `context.document.onSelectionChanged` 事件，当用户在 Word 中划选文字时，前端实时同步显示选中文本与实时字数统计。
  3. **OpenAI 兼容 SSE 流式输出**：
     - 编写通用的 SSE 解析器，读取 `ReadableStream` 并处理 `data: {...}`，实现文字实时打字机上屏。
     - 预设 8 大快捷动作：润色、扩写、精简、纠错、商务公文、学术规范、中英互译。
  4. **本地极简开发服务 (`server.py` + `start.bat`)**：
     - 针对 Office 插件强制要求 HTTPS 调试的限制，编写免依赖 Python 服务，自动调用 `openssl` 生成本地自签名证书。

---

### 【Phase 2】本地接口连接异常攻坚与反向代理网关 (v1.1)
* **用户痛点反馈**：
  - 用户在插件设置中填入本地大模型接口 `http://127.0.0.1:8045/v1` 时，测试连接抛出 `Failed to fetch` 异常。
* **深层根因排查**：
  1. **混合内容安全限制（Mixed Content）**：Word 插件面板运行在 HTTPS（`https://localhost:3000`）环境下，现代浏览器及 Word WebView2 内核出于安全机制，**强制静默拦截**在 HTTPS 页面中发往 HTTP 明文端口（如 `http://127.0.0.1:8045`）的异步网络请求。
  2. **跨域资源共享（CORS）拦截**：本地 8045 接口未配置 `Access-Control-Allow-Origin: *` 响应头，导致浏览器跨源预检直接失败。
* **架构解决方案**：
  - 在 `server.py` 中实现反向代理端点 `/api-proxy`。
  - 在 `taskpane.js` 中封装透明请求路由 `resolveFetchEndpointAndHeaders`：若检测到当前页面是 HTTPS 且请求目标是 `http://`（或本地地址），自动重写请求地址为同源的 `/api-proxy`，并通过 `X-Target-Url` Header 携带真实上游地址。
  - `server.py` 在服务端使用标准库 `urllib.request` 进行跨协议中继，流式转发 chunked 数据，彻底绕过浏览器的 Mixed Content 与 CORS 限制。

---

### 【Phase 3】生产力全面跃升：五大高级能力 (v2.0)
根据深度办公场景的实际痛点，全面重构 UI 与核心交互逻辑：

1. **Word 原生富文本与表格排版写回 (`insertHtml`)**：
   - *问题*：早期纯文本替换（`insertText`）会导致模型输出的 Markdown 加粗、标题、列表、表格退化为无意义的纯字符。
   - *解决*：实现 `markdownToWordHtml` 适配器，将 Markdown 解析为带特定 Word 内联样式的 HTML，调用 `selection.insertHtml(wordHtml, Word.InsertLocation.replace)`。使得 Word 自动将其转变为原生物理加粗、分级标题、编号列表以及**原生可编辑表格**。同时保留纯文本替换选项供用户切换。
2. **多版本候选生成 (Multi-draft A/B 测试)**：
   - 增加【多版本候选 (A/B)】开关。开启后提示模型并行输出 3 种独立风格版本（标准平衡版、精简干练版、文雅生动版）。
   - 前端动态解析版本分隔符，在结果卡片顶部渲染选项卡，用户可无缝左右切换对比并挑选写入。
3. **DeepSeek-R1 深度思考链折叠解析**：
   - 针对当前热门的推理大模型（如 DeepSeek-R1、Ollama r1 等），流式引擎实时监测 `<think>` 标签。
   - 思考过程自动渲染到专用的浅灰色折叠面板中（支持收起/展开），正文区域只展示 `</think>` 之后的实际回答。
   - **关键机制**：向 Word 文档写回时，纯净剔除思考部分，确保不污染文档。
4. **整篇长文档全局审阅模式**：
   - 顶部增加【选中文字处理】与【整篇文档审阅】Tab。
   - 在全文模式下，通过 `context.document.body.text` 抓取全文内容，快捷提供：提炼结构大纲（分级标题）、全文要点摘要、全面校对审查报告（错别字/逻辑矛盾/术语不统一）、提取关键数据表格。
5. **历史记录与 Prompt 收藏库 (LocalStorage)**：
   - **历史记录**：保存最近 30 条处理记录，支持在侧拉抽屉中查看、一键载入重新写回与复制。
   - **Prompt 收藏库**：内置多种高频实用模板，支持用户在界面中自定义新增、编辑、删除高频 Prompt。

---

### 【Phase 4】UI 美化：精致 Office 原生风 (v2.1)
* **目标**：让侧边栏与 Word 界面融为一体，专业、克制、不突兀。
* **主要改动**：
  1. **`taskpane.css` 全量重写**：采用 Fluent 2 中性灰层次（`#f5f5f4` 底 / `#fff` 卡片 / `#e6e6e6` 描边）+ Word 品牌蓝 `#185abd`；字体栈优先 `Segoe UI Variable` / `Microsoft YaHei UI`；统一 4/6/8/10px 圆角体系；所有颜色、阴影、动效集中在 `:root` CSS 变量中，后续换肤只需改变量。
  2. **图标统一为内联 SVG**：`taskpane.html` 顶部用 `<symbol>` 定义整套线性图标（`#i-sparkle`、`#i-translate` 等），通过 `<svg class="ic"><use href="#i-xxx"/></svg>` 引用，替换原先的 emoji（emoji 在 Windows/Mac/网页版 Word 中显示不一致）。快捷动作图标放在带低饱和色调的方块中（`.tile-icon.tone-blue/teal/violet/amber`）。
  3. **组件升级**：分段控件式模式切换（`.mode-tab-bar`/`.segmented`）；Fluent 风开关（`.switch`）替代多版本复选框；输入框改为带底栏的 composer（含 `Ctrl+Enter` 提示）；全文审阅动作改为带箭头的列表；Fluent 底线式表单输入；结果卡片加品牌色光圈；写回栏主按钮独占一行。
  4. **状态反馈**：`taskpane.js` 中的内联样式占位改为样式类——加载用跳动圆点（`.loading-line .dots`），错误用 `.result-error`（并对错误信息做了 `escapeHtml`），流式输出时 `#aiResultContent` 带 `.is-streaming` 显示闪烁光标，历史记录空状态 `.empty-state`。
  5. **无障碍**：所有可交互元素有 `:focus-visible` 焦点环；图标按钮加 `aria-label`；支持 `prefers-reduced-motion`；宽度 <300px 时快捷动作自动变单列。
* **约束**：所有 `id` 保持不变，JS 业务逻辑未改动；`data-action` / `data-doc-action` 与 JS 中的动作键一一对应。
* **注意**：`#promptChipsList` 在 HTML 中为空，由 `renderPromptChips()` 根据收藏库动态填充。

---

### 【Phase 5】桌面端一键全自动安装工具 (v2.2)
* **目标**：彻底免去用户手动创建 Windows 网络共享、手动进入 Word 信任中心配置 UNC 路径的繁琐操作。
* **主要工作**：
  1. **注册表注入 (`install-to-word.bat`)**：
     - 利用 Office Web 扩展框架底层的开发者注册机制，将插件绝对路径写入 `HKCU\Software\Microsoft\Office\16.0\WEF\Developer\<AddIn-ID>`。无需管理员权限即可让桌面端 Word 自动识别该插件。
     - 自动检测并写入受信任目录项 `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs`，自动清理过期的 WEF 缓存。
  2. **配套卸载脚本 (`uninstall-from-word.bat`)**：
     - 一键清理上述注册表项与共享配置，方便重置与调试。

---

## 3. 项目目录结构与文件职责

```text
word-ai-copilot-addin/
├── manifest.xml           # [核心] Office Web Add-in 清单描述文件 (Ribbon/权限/URL)
├── install-to-word.bat    # [核心] Windows 桌面版 Word 一键全自动安装脚本
├── uninstall-from-word.bat# Windows 桌面版 Word 一键卸载脚本
├── server.py              # [核心] 本地开发服务器 (HTTPS + 自签名证书 + API 代理中继)
├── start.bat              # Windows 双击一键启动批处理脚本
├── README.md              # 面向终端用户的快速上手与侧载安装手册
├── DEVELOPMENT.md         # [本文档] 面向开发者与 AI 工具的架构解析与开发记录
├── assets/                # 图标资源目录
│   ├── icon-16.png        # Ribbon 功能区 16x16 按钮图标
│   ├── icon-32.png        # Ribbon 功能区 32x32 按钮图标
│   ├── icon-64.png        # 高分屏 64x64 图标
│   └── icon-80.png        # 插件管理器 80x80 图标
├── certs/                 # 本地调试 SSL 证书存储目录 (由 server.py 自动生成)
│   ├── localhost.crt
│   └── localhost.key
└── src/                   # 插件前端应用源码
    ├── taskpane.html      # 侧边栏主界面 (Fluent 2.0 DOM 结构、抽屉组件)
    ├── taskpane.css       # 样式文件 (微软 Fluent 主题变量、卡片、响应式排版)
    ├── taskpane.js        # [核心] 业务逻辑 (Word.js 交互/流式 SSE/思考过滤/历史管理)
    └── artifact_preview.html # 独立自包含的沙箱预览测试单文件
```

---

## 4. 核心逻辑与代码实现深度剖析

### 4.1 Word.js 富文本与表格原生写回
```javascript
// src/taskpane.js
function markdownToWordHtml(markdownText) {
  let rawHtml = window.marked ? marked.parse(markdownText) : `<p>${markdownText}</p>`;
  // 为 table 注入满足 Word 布局引擎解析的内联样式
  rawHtml = rawHtml.replace(/<table>/g, '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #94a3b8;">');
  rawHtml = rawHtml.replace(/<th>/g, '<th style="background-color: #f1f5f9; font-weight: bold; border: 1px solid #94a3b8; padding: 6px 10px;">');
  rawHtml = rawHtml.replace(/<td>/g, '<td style="border: 1px solid #94a3b8; padding: 6px 10px;">');
  return rawHtml;
}

async function replaceSelectionWithRichText(markdownText) {
  const wordHtml = markdownToWordHtml(markdownText);
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    // 关键：使用 insertHtml 将 HTML 映射为 Word 原生排版对象
    selection.insertHtml(wordHtml, Word.InsertLocation.replace);
    await context.sync();
  });
}
```

### 4.2 跨协议反向代理中继网关
```python
# server.py (核心代理片段)
def do_POST(self):
    if self.path.startswith('/api-proxy'):
        target_url = self.headers.get('X-Target-Url')
        content_length = int(self.headers.get('Content-Length', 0))
        post_body = self.rfile.read(content_length) if content_length > 0 else None

        upstream_req = urllib.request.Request(
            target_url,
            data=post_body,
            headers={'Content-Type': 'application/json', 'Authorization': self.headers.get('Authorization', '')},
            method='POST'
        )
        with urllib.request.urlopen(upstream_req, timeout=60) as upstream_resp:
            self.send_response(upstream_resp.status)
            for k, v in upstream_resp.headers.items():
                if k.lower() in ['content-type', 'cache-control']:
                    self.send_header(k, v)
            self.end_headers()
            while True:
                chunk = upstream_resp.read(1024)
                if not chunk: break
                self.wfile.write(chunk)
                self.wfile.flush()
```

### 4.3 DeepSeek-R1 思考标签动态分离
```javascript
// src/taskpane.js
function handleStreamChunk(fullText) {
  if (fullText.includes("<think>")) {
    const parts = fullText.split("</think>");
    if (parts.length > 1) {
      reasoningStatusText.textContent = "思考完成";
      reasoningBody.textContent = parts[0].replace("<think>", "").trim();
      renderMarkdownResult(parts[1].trim()); // 正文只渲染思考之后的部分
    } else {
      reasoningStatusText.textContent = "深度思考中...";
      reasoningBody.textContent = fullText.replace("<think>", "").trim();
    }
  } else {
    renderMarkdownResult(fullText);
  }
}
```

---

## 5. 关键技术踩坑与调试经验 (Gotchas)

1. **Word 内部右键无反应 / 无法调出 F12 控制台**：
   - Word 桌面端的 WebView2 默认可能禁用了右键菜单。
   - **解决方式**：在 Windows 应用商店安装 **Microsoft Edge DevTools Preview**，打开后会自动捕获并列出当前正在运行的 Word Add-in 进程，点击即可直接进入完整的 Chromium 开发者工具。
2. **首次启动提示证书不信任 / ERR_CERT_AUTHORITY_INVALID**：
   - 现象：Word 侧边栏白屏或报错无法连接。
   - **解决方式**：启动 `server.py` 后，务必先在 Chrome/Edge 浏览器中打开 `https://localhost:3000/src/taskpane.html`，点击【高级】→【继续前往 localhost】信任自签名证书，之后 Word 内部的 WebView2 才能信任该本地连接。
3. **Word.js 中的异步时序问题**：
   - 所有对文档的操作必须放在 `Word.run(async (context) => { ... await context.sync(); })` 闭包内。如果读取属性（如 `selection.text`）未先调用 `selection.load("text")` 并 `await context.sync()`，会抛出 `PropertyNotLoaded` 异常。

---

## 6. 后续迭代路线与扩展建议 (Roadmap for Next Developers)

若其他 AI 开发工具继续迭代本项目，建议按以下优先级进行扩展：

1. **批注与修订模式支持 (Track Changes & Comments)**：
   - 利用 `Word.Comment` API，将 AI 的语法修改建议作为 Word 原生“批注”插入，或者直接开启“修订模式”展示被删除与新增的字词，方便用户逐条点击【接受修订】。
2. **团队风格知识库 (RAG & Few-Shot)**：
   - 在设置中支持配置企业级行文规范文档或术语表，AI 在处理公文或学术论文时，自动将其作为上下文注入。
3. **生产环境云端托管**：
   - 本项目纯前端设计，零后端依赖。若要团队共享，可将 `src/` 与 `assets/` 文件夹直接部署到 **Vercel / Cloudflare Pages / GitHub Pages**。只需将 `manifest.xml` 中的 `https://localhost:3000` 批量替换为正式 HTTPS 域名，即可通过 Microsoft 365 管理中心向全公司一键分发。
