/**
 * Word AI Copilot Add-in 2.0 - Core Logic
 * 支持 Word.js API 原生富文本与表格排版、DeepSeek-R1 思考折叠、多版本 A/B 候选、全文审阅与历史收藏
 */

// 预设服务商配置
const PROVIDER_PRESETS = {
  local_custom: {
    name: "本地通用接口 (如 127.0.0.1:8045)",
    baseUrl: "http://127.0.0.1:8045/v1",
    model: "",
    help: "本地大模型接口或转发器 (支持自动中继代理，不受 HTTPS 混合内容限制)"
  },
  deepseek: {
    name: "DeepSeek",
    baseUrl: "https://api.deepseek.com/v1",
    model: "deepseek-chat",
    help: "中国大陆高性价比顶尖模型，推荐使用"
  },
  openai: {
    name: "OpenAI",
    baseUrl: "https://api.openai.com/v1",
    model: "gpt-4o-mini",
    help: "OpenAI 官方或反代接口"
  },
  siliconflow: {
    name: "SiliconFlow (硅基流动)",
    baseUrl: "https://api.siliconflow.cn/v1",
    model: "deepseek-ai/DeepSeek-V3",
    help: "国内高速推理平台"
  },
  moonshot: {
    name: "Moonshot (Kimi)",
    baseUrl: "https://api.moonshot.cn/v1",
    model: "moonshot-v1-8k",
    help: "Moonshot 开放平台"
  },
  ollama: {
    name: "本地 Ollama",
    baseUrl: "http://localhost:11434/v1",
    model: "qwen2.5:7b",
    help: "本地运行的开源大模型，API Key 可随意填写"
  },
  custom: {
    name: "自定义",
    baseUrl: "",
    model: "",
    help: "任意兼容 OpenAI 规范的 API 接口"
  }
};

// 选中文本快捷动作
const ACTION_PROMPTS = {
  polish: {
    name: "润色优化",
    prompt: "请对以下文档内容进行精细润色与优化。保持原意不变，提升文字的流畅性、表现力与优美度，修正生硬表达：\n\n"
  },
  expand: {
    name: "扩写充实",
    prompt: "请在保留原有核心论点和逻辑的基础上，对以下内容进行逻辑扩充与细节丰富，增加合理的论述、说明或背景铺垫：\n\n"
  },
  summarize: {
    name: "提炼精简",
    prompt: "请对以下内容进行高度概括与精炼，提炼出最核心的观点与结论，去除冗余修饰，语言简明扼要：\n\n"
  },
  grammar: {
    name: "语法纠错",
    prompt: "请严格检查以下文本中的错别字、标点符号误用、语病以及语法错误。请直接输出修正后的完整文本：\n\n"
  },
  formal: {
    name: "商务公文",
    prompt: "请将以下内容改写为严谨、专业、规范的商务公文/职场公函风格，注意措辞得体分寸，符合正式公文行文习惯：\n\n"
  },
  academic: {
    name: "学术规范",
    prompt: "请将以下内容改写为客观、严密、符合学术论文规范的学术论述风格，注意逻辑连贯性与专业用语规范：\n\n"
  },
  to_en: {
    name: "翻译为英文",
    prompt: "请将以下内容翻译为地道流畅的英文（采用母语级表达，避免直译与中式英语）：\n\n"
  },
  to_zh: {
    name: "翻译为中文",
    prompt: "请将以下内容准确翻译为流畅易读的中文，做到信达雅：\n\n"
  }
};

// 全文审阅动作
const DOC_ACTION_PROMPTS = {
  doc_outline: {
    name: "提炼结构大纲",
    prompt: "请分析以下整篇文档，提取并生成清晰的结构化大纲目录。要求使用标准的 Markdown 分级标题（#、##、###）和要点列表，标明每个章节的核心主旨：\n\n"
  },
  doc_summary: {
    name: "全文要点摘要",
    prompt: "请对以下整篇文档进行全局分析，输出一份结构化摘要。包括：1. 文档核心背景与主题；2. 核心论点或主要事实；3. 关键结论与行动建议：\n\n"
  },
  doc_audit: {
    name: "全面校对审查报告",
    prompt: "请作为资深文档主编，对整篇文档进行全方位审阅，输出一份专业的【文档审查与修改意见报告】。重点排查：1. 错别字与语病；2. 专业术语前后不统一；3. 论据或逻辑前后矛盾；4. 排版与标点规范性，并逐条给出修改建议：\n\n"
  },
  doc_table: {
    name: "整理关键数据表格",
    prompt: "请提炼以下文档中的核心数据、指标、对比项或关键信息，整理为一个或多个结构清晰的标准 Markdown 表格（包含表头、对齐线和具体数据）：\n\n"
  }
};

// 默认内置的 Prompt 收藏项
const DEFAULT_PROMPT_PRESETS = [
  { id: "p1", title: "整理对比三线表格", content: "请将以下内容提炼整理为标准 Markdown 对比表格，包含清晰的表头与核心指标对比：" },
  { id: "p2", title: "要点加粗突出", content: "请对以下内容进行提炼总结，并将关键专业术语、核心数据与结论用 **加粗** 重点标注：" },
  { id: "p3", title: "小红书生动文风", content: "请将以下内容改编为生动活泼的小红书爆款文案风格，语言风趣幽默，适当搭配 emoji 表情与标签：" },
  { id: "p4", title: "合规与免责声明", content: "请为以下内容补充严谨周密的合规提示与法律免责声明条款：" }
];

// 全局应用状态
const AppState = {
  isOfficeInitialized: false,
  activeMode: "selection", // "selection" | "document"
  currentSelection: "",
  fullDocText: "",
  lastGeneratedText: "",
  // 划选局部写回状态
  selectedResultText: "",
  // 多版本草稿状态
  drafts: [], // Array of string
  activeDraftIndex: 0,
  // 思考过程状态
  reasoningText: "",
  isReasoning: false,
  isReasoningCollapsed: false,
  // 生成控制
  abortController: null,
  isGenerating: false,
  // 配置持久化
  config: {
    provider: "local_custom",
    baseUrl: "http://127.0.0.1:8045/v1",
    apiKey: "sk-test",
    model: "",
    temperature: 0.7,
    systemPrompt: "你是一位专业的 Microsoft Word 文档创作助手与资深文字编辑。请根据用户的指令直接输出高质量处理结果。支持 Markdown 排版（加粗、列表、表格、标题），无需多余寒暄。"
  },
  // 历史记录与收藏
  history: [],
  promptPresets: []
};

// ==========================================================================
// 初始化与生命周期
// ==========================================================================
document.addEventListener("DOMContentLoaded", () => {
  loadConfigFromStorage();
  loadHistoryFromStorage();
  loadPromptsFromStorage();
  initUIEventListeners();

  if (window.Office && window.Office.onReady) {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Word) {
        AppState.isOfficeInitialized = true;
        console.log("Office.js Word 环境已成功就绪");

        registerSelectionChangeListener();
        refreshDocumentSelection();
      } else {
        console.warn("当前不在 Word 宿主环境中，进入浏览器演示模式");
        initMockData();
      }
    });
  } else {
    console.warn("未检测到 Office 环境，进入浏览器演示模式");
    initMockData();
  }
});

function initMockData() {
  AppState.currentSelection = "人工智能正在深刻重塑数字化办公体验。通过无缝结合大语言模型，现代办公软件能够高效协助用户完成文稿润色、内容扩充、多语言翻译及精简总结。";
  AppState.fullDocText = "数字化转型报告\n\n一、项目背景\n人工智能技术与办公自动化的深度融合，成为企业降本增效的核心抓手。\n\n二、核心效益\n提升文稿起草速度 60%，降低错别字与语法错误率 90%。";
  updateSelectionUI(AppState.currentSelection);
}

// ==========================================================================
// Word.js 文档交互（含原生富文本排版写回）
// ==========================================================================

async function registerSelectionChangeListener() {
  if (!AppState.isOfficeInitialized) return;

  try {
    await Word.run(async (context) => {
      context.document.onSelectionChanged.add(() => {
        if (AppState.activeMode === "selection") {
          refreshDocumentSelection();
        }
      });
      await context.sync();
    });
  } catch (error) {
    console.error("注册选区监听失败:", error);
  }
}

async function refreshDocumentSelection() {
  if (!AppState.isOfficeInitialized) {
    return;
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const text = selection.text ? selection.text.trim() : "";
      AppState.currentSelection = text;
      updateSelectionUI(text);
    });
  } catch (error) {
    console.error("读取选区失败:", error);
  }
}

async function readFullDocumentContent() {
  if (!AppState.isOfficeInitialized) {
    showToast("测试模式：已读取模拟全文");
    const countEl = document.getElementById("docTotalWords");
    countEl.textContent = `共 ${AppState.fullDocText.length} 字`;
    return;
  }

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();

      const text = body.text ? body.text.trim() : "";
      AppState.fullDocText = text;
      const countEl = document.getElementById("docTotalWords");
      countEl.textContent = `共 ${text.length} 字`;
      showToast(`已读取全文（共 ${text.length} 字）`);
    });
  } catch (error) {
    console.error("读取全文失败:", error);
    showToast("读取全文失败: " + error.message);
  }
}

/**
 * 将 Markdown 转化为符合 Word 样式的标准 HTML
 */
function markdownToWordHtml(markdownText) {
  let rawHtml = "";
  if (window.marked) {
    rawHtml = marked.parse(markdownText);
  } else {
    rawHtml = `<p>${markdownText.replace(/\n/g, "<br>")}</p>`;
  }

  // 为表格注入优雅的 Word 原生支持样式
  rawHtml = rawHtml.replace(/<table>/g, '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #94a3b8; font-family: Segoe UI, Arial, sans-serif;">');
  rawHtml = rawHtml.replace(/<th>/g, '<th style="background-color: #f1f5f9; font-weight: bold; border: 1px solid #94a3b8; padding: 6px 10px;">');
  rawHtml = rawHtml.replace(/<td>/g, '<td style="border: 1px solid #94a3b8; padding: 6px 10px;">');

  return rawHtml;
}

/**
 * 原生富文本替换选中文本
 */
async function replaceSelectionWithRichText(markdownText) {
  if (!AppState.isOfficeInitialized) {
    showToast("测试模式：已模拟富文本替换");
    AppState.currentSelection = markdownText;
    updateSelectionUI(markdownText);
    return;
  }

  try {
    const wordHtml = markdownToWordHtml(markdownText);
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertHtml(wordHtml, Word.InsertLocation.replace);
      await context.sync();
      showToast("已按 Word 原生富文本格式替换成功！");
      refreshDocumentSelection();
    });
  } catch (error) {
    console.error("富文本替换失败:", error);
    showToast("富文本替换失败，正在尝试纯文本替换...");
    replaceSelectionWithPlainText(markdownText);
  }
}

/**
 * 纯文本替换选中文本
 */
async function replaceSelectionWithPlainText(text) {
  if (!AppState.isOfficeInitialized) {
    showToast("测试模式：已替换纯文本");
    AppState.currentSelection = text;
    updateSelectionUI(text);
    return;
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await context.sync();
      showToast("已成功替换为纯文本！");
      refreshDocumentSelection();
    });
  } catch (error) {
    console.error("纯文本替换失败:", error);
    showToast("写入失败: " + error.message);
  }
}

/**
 * 在选中文本后方追加新段落 (富文本)
 */
async function appendAfterSelectionInDoc(markdownText) {
  if (!AppState.isOfficeInitialized) {
    showToast("测试模式：已在后方追加内容");
    return;
  }

  try {
    const wordHtml = markdownToWordHtml(markdownText);
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertHtml(wordHtml, Word.InsertLocation.after);
      await context.sync();
      showToast("已成功在后方插入新段落！");
    });
  } catch (error) {
    console.error("追加内容失败:", error);
    showToast("追加失败: " + error.message);
  }
}

/**
 * 在光标当前位置插入内容
 */
async function insertAtCursorInDoc(markdownText) {
  if (!AppState.isOfficeInitialized) {
    showToast("测试模式：已在光标处插入内容");
    return;
  }

  try {
    const wordHtml = markdownToWordHtml(markdownText);
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertHtml(wordHtml, Word.InsertLocation.end);
      await context.sync();
      showToast("已成功插入到光标位置！");
    });
  } catch (error) {
    console.error("光标插入失败:", error);
    showToast("插入失败: " + error.message);
  }
}

// ==========================================================================
// UI 状态与事件绑定
// ==========================================================================

function updateSelectionUI(text) {
  const previewEl = document.getElementById("selectionPreview");
  const countEl = document.getElementById("selectionWordCount");

  if (!text || text.length === 0) {
    previewEl.textContent = "在 Word 文档中选中文字后，此处将自动同步预览...";
    previewEl.classList.add("empty");
    countEl.textContent = "0 字";
  } else {
    previewEl.textContent = text;
    previewEl.classList.remove("empty");
    countEl.textContent = `${text.length} 字`;
  }
}

function initUIEventListeners() {
  // 模式切换
  const tabSelection = document.getElementById("tabModeSelection");
  const tabDoc = document.getElementById("tabModeDocument");
  const secSelection = document.getElementById("selectionSection");
  const secDoc = document.getElementById("documentSection");
  const secQuick = document.getElementById("quickActionsSection");

  tabSelection.onclick = () => {
    AppState.activeMode = "selection";
    tabSelection.classList.add("active");
    tabDoc.classList.remove("active");
    secSelection.style.display = "block";
    secQuick.style.display = "block";
    secDoc.style.display = "none";
  };

  tabDoc.onclick = () => {
    AppState.activeMode = "document";
    tabDoc.classList.add("active");
    tabSelection.classList.remove("active");
    secSelection.style.display = "none";
    secQuick.style.display = "none";
    secDoc.style.display = "block";
    if (!AppState.fullDocText) {
      readFullDocumentContent();
    }
  };

  document.getElementById("btnReadFullDoc").onclick = readFullDocumentContent;
  document.getElementById("btnRefreshSelection").onclick = refreshDocumentSelection;

  // 选中文本快捷动作
  document.querySelectorAll(".action-btn").forEach((btn) => {
    btn.onclick = () => {
      const actionKey = btn.getAttribute("data-action");
      handleQuickAction(actionKey);
    };
  });

  // 全文快捷动作
  document.querySelectorAll(".doc-action-btn").forEach((btn) => {
    btn.onclick = () => {
      const docActionKey = btn.getAttribute("data-doc-action");
      handleDocAction(docActionKey);
    };
  });

  // 自定义指令执行
  document.getElementById("btnSubmitCustom").onclick = handleCustomPromptSubmit;
  document.getElementById("customPromptInput").onkeydown = (e) => {
    if (e.ctrlKey && e.key === "Enter") {
      e.preventDefault();
      handleCustomPromptSubmit();
    }
  };

  // 思考折叠面板点击收放
  document.getElementById("reasoningToggle").onclick = () => {
    const body = document.getElementById("reasoningBody");
    const arrow = document.getElementById("reasoningArrow");
    AppState.isReasoningCollapsed = !AppState.isReasoningCollapsed;
    if (AppState.isReasoningCollapsed) {
      body.classList.add("collapsed");
      arrow.classList.add("collapsed");
    } else {
      body.classList.remove("collapsed");
      arrow.classList.remove("collapsed");
    }
  };

  // 停止生成
  document.getElementById("btnStopGeneration").onclick = stopAIGeneration;

  // 写回按钮组
  document.getElementById("btnReplaceRichText").onclick = () => {
    const text = getCurrentActiveOutputText();
    if (text) replaceSelectionWithRichText(text);
  };

  document.getElementById("btnReplacePlainText").onclick = () => {
    const text = getCurrentActiveOutputText();
    if (text) replaceSelectionWithPlainText(text);
  };

  document.getElementById("btnAppendAfter").onclick = () => {
    const text = getCurrentActiveOutputText();
    if (text) appendAfterSelectionInDoc(text);
  };

  document.getElementById("btnInsertAtCursor").onclick = () => {
    const text = getCurrentActiveOutputText();
    if (text) insertAtCursorInDoc(text);
  };

  document.getElementById("btnCopyResult").onclick = () => {
    const text = getCurrentActiveOutputText();
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
      showToast("已成功复制到剪贴板！");
    });
  };

  // 差异对比
  document.getElementById("btnToggleDiff").onclick = () => {
    const diffContainer = document.getElementById("diffContainer");
    const isHidden = diffContainer.style.display === "none";
    diffContainer.style.display = isHidden ? "flex" : "none";
    document.getElementById("btnToggleDiff").textContent = isHidden ? "收起对比 (Diff)" : "查看与原文对比 (Diff)";

    if (isHidden) {
      document.getElementById("diffOriginalText").textContent =
        (AppState.activeMode === "selection" ? AppState.currentSelection : AppState.fullDocText) || "（无原文字）";
      document.getElementById("diffNewText").textContent = getCurrentActiveOutputText() || "（无生成结果）";
    }
  };

  // 历史记录抽屉
  document.getElementById("btnOpenHistory").onclick = openHistoryDrawer;
  document.getElementById("btnCloseHistory").onclick = closeHistoryDrawer;
  document.getElementById("btnClearHistory").onclick = clearHistory;

  // Prompt 收藏抽屉
  document.getElementById("btnOpenPrompts").onclick = openPromptsDrawer;
  document.getElementById("btnClosePrompts").onclick = closePromptsDrawer;
  document.getElementById("btnAddPromptPreset").onclick = addNewPromptPreset;

  // 设置抽屉
  document.getElementById("btnSettings").onclick = openSettingsDrawer;
  document.getElementById("btnCloseSettings").onclick = closeSettingsDrawer;
  document.getElementById("btnSaveSettings").onclick = saveSettingsFromDrawer;
  document.getElementById("providerPreset").onchange = (e) => applyProviderPreset(e.target.value);
  document.getElementById("btnTestConnection").onclick = testAPIConnection;
  document.getElementById("btnToggleKeyVisible").onclick = () => {
    const keyInput = document.getElementById("apiKeyInput");
    const btn = document.getElementById("btnToggleKeyVisible");
    if (keyInput.type === "password") {
      keyInput.type = "text";
      btn.textContent = "隐藏";
    } else {
      keyInput.type = "password";
      btn.textContent = "显示";
    }
  };
  document.getElementById("temperatureSlider").oninput = (e) => {
    document.getElementById("temperatureVal").textContent = e.target.value;
  };

  // 划词局部写回监听
  const resultContentEl = document.getElementById("aiResultContent");
  const scopeNoticeEl = document.getElementById("selectionScopeNotice");
  const clearSelectionBtn = document.getElementById("btnClearResultSelection");

  function updateResultSelectionState() {
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed && resultContentEl.contains(sel.anchorNode)) {
      const selected = sel.toString().trim();
      if (selected.length > 0) {
        AppState.selectedResultText = selected;
        scopeNoticeEl.style.display = "flex";
        document.getElementById("selectionScopeText").textContent = `已划选 ${selected.length} 字（将仅写入此部分）`;
        return;
      }
    }
  }

  resultContentEl.addEventListener("mouseup", updateResultSelectionState);
  resultContentEl.addEventListener("keyup", updateResultSelectionState);

  if (clearSelectionBtn) {
    clearSelectionBtn.onclick = () => {
      AppState.selectedResultText = "";
      scopeNoticeEl.style.display = "none";
      const sel = window.getSelection();
      if (sel) sel.removeAllRanges();
    };
  }

  // 多版本 Tab 点击委托
  document.querySelectorAll(".draft-tab").forEach(tab => {
    tab.onclick = () => {
      const idx = parseInt(tab.getAttribute("data-index"), 10);
      switchDraft(idx);
    };
  });
}

function getCurrentActiveOutputText() {
  if (AppState.drafts.length > 0 && AppState.drafts[AppState.activeDraftIndex]) {
    return AppState.drafts[AppState.activeDraftIndex];
  }
  return AppState.lastGeneratedText;
}

function switchDraft(index) {
  AppState.activeDraftIndex = index;
  document.querySelectorAll(".draft-tab").forEach(tab => {
    const i = parseInt(tab.getAttribute("data-index"), 10);
    tab.classList.toggle("active", i === index);
  });
  const currentText = AppState.drafts[index] || "";
  renderMarkdownResult(currentText);
}

// ==========================================================================
// AI 业务核心与流式中继调用
// ==========================================================================

function resolveFetchEndpointAndHeaders(targetUrl, headers = {}) {
  const isPageHttps = window.location.protocol === "https:";
  const isTargetHttp = targetUrl.startsWith("http://");
  const isLocalHost = window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";

  // 仅在本地运行 server.py 时通过 /api-proxy 中继 HTTP 请求解决 Mixed Content
  if (isPageHttps && isTargetHttp && isLocalHost) {
    return {
      fetchUrl: "/api-proxy",
      fetchHeaders: {
        ...headers,
        "X-Target-Url": targetUrl
      }
    };
  }
  return {
    fetchUrl: targetUrl,
    fetchHeaders: headers
  };
}

async function handleQuickAction(actionKey) {
  const action = ACTION_PROMPTS[actionKey];
  if (!action) return;

  if (!AppState.currentSelection) {
    showToast("请先在 Word 文档中选中需要处理的文字！");
    return;
  }

  const isMultiDraft = document.getElementById("multiDraftCheck").checked;
  if (isMultiDraft) {
    const prompt = `请针对以下选中文档内容执行【${action.name}】。\n为了方便用户对比挑选，请同时生成 3 种不同风格的独立候选版本，并使用严格的分隔符分隔：\n\n===版本1:标准平衡===\n(版本1内容)\n===版本2:精简干练===\n(版本2内容)\n===版本3:文雅生动===\n(版本3内容)\n\n待处理文本：\n${AppState.currentSelection}`;
    runAIGeneration(prompt, `正在生成【${action.name}】3 个候选版本...`, true);
  } else {
    const prompt = `${action.prompt}\n\n待处理文本：\n${AppState.currentSelection}\n\n【严格输出规范】：\n1. 仅输出最终修改后的单一一版正文结果，严禁输出方案一/方案二等多个候选方案；\n2. 严禁包含任何前缀引言、解释说明、提示语（如“修改如下：”、“好的”、“润色结果：”等）或寒暄客套；\n3. 直接输出可直接写回 Word 的纯净最终正文内容。`;
    runAIGeneration(prompt, `正在执行【${action.name}】...`, false);
  }
}

async function handleDocAction(docActionKey) {
  const action = DOC_ACTION_PROMPTS[docActionKey];
  if (!action) return;

  if (!AppState.fullDocText) {
    await readFullDocumentContent();
  }

  if (!AppState.fullDocText) {
    showToast("文档内容为空，无法分析！");
    return;
  }

  const prompt = `${action.prompt}${AppState.fullDocText}`;
  runAIGeneration(prompt, `正在进行全局分析【${action.name}】...`, false);
}

async function handleCustomPromptSubmit() {
  const promptInput = document.getElementById("customPromptInput");
  const userInstruction = promptInput.value.trim();

  if (!userInstruction) {
    showToast("请输入您的指令或要求！");
    return;
  }

  const includeSelection = document.getElementById("includeSelectionCheck").checked;
  let contextText = AppState.activeMode === "selection" ? AppState.currentSelection : AppState.fullDocText;
  let finalPrompt = userInstruction;

  if (includeSelection && contextText) {
    finalPrompt = `参考以下文档内容：\n"""\n${contextText}\n"""\n\n指令要求：\n${userInstruction}`;
  }

  runAIGeneration(finalPrompt, "正在根据指令生成...", false);
}

/**
 * 发起请求（处理流式 SSE 与 DeepSeek-R1 <think> 折叠）
 */
async function runAIGeneration(userPrompt, statusText, isMultiDraft = false) {
  if (AppState.isGenerating) {
    stopAIGeneration();
  }

  const isLocalHttp = AppState.config.baseUrl && (AppState.config.baseUrl.includes("127.0.0.1") || AppState.config.baseUrl.includes("localhost"));
  if (!AppState.config.apiKey && AppState.config.provider !== "ollama" && !isLocalHttp) {
    showToast("请先在右上角【设置】中配置您的 API Key！");
    openSettingsDrawer();
    return;
  }

  // 重置 UI 显示
  const resultSection = document.getElementById("resultSection");
  const resultContent = document.getElementById("aiResultContent");
  const statusIndicator = document.getElementById("generationStatus");
  const stopBtn = document.getElementById("btnStopGeneration");
  const reasoningPanel = document.getElementById("reasoningPanel");
  const reasoningBody = document.getElementById("reasoningBody");
  const reasoningStatusText = document.getElementById("reasoningStatusText");
  const draftTabsContainer = document.getElementById("draftTabsContainer");

  resultSection.style.display = "block";
  resultContent.innerHTML = `<div class="loading-line"><span class="dots"><i></i><i></i><i></i></span>正在连接模型…</div>`;
  resultContent.classList.add("is-streaming");
  statusIndicator.textContent = statusText || "思考中...";
  stopBtn.style.display = "inline-flex";

  reasoningPanel.style.display = "none";
  reasoningBody.textContent = "";
  draftTabsContainer.style.display = "none";

  AppState.isGenerating = true;
  AppState.lastGeneratedText = "";
  AppState.drafts = [];
  AppState.activeDraftIndex = 0;
  AppState.reasoningText = "";
  AppState.isReasoning = false;
  AppState.abortController = new AbortController();

  let accumulatedRaw = "";

  try {
    const rawEndpoint = `${AppState.config.baseUrl.replace(/\/+$/, "")}/chat/completions`;
    const defaultHeaders = {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${AppState.config.apiKey || "dummy-key"}`
    };
    const { fetchUrl, fetchHeaders } = resolveFetchEndpointAndHeaders(rawEndpoint, defaultHeaders);

    const response = await fetch(fetchUrl, {
      method: "POST",
      headers: fetchHeaders,
      signal: AppState.abortController.signal,
      body: JSON.stringify({
        model: AppState.config.model,
        messages: [
          { role: "system", content: AppState.config.systemPrompt },
          { role: "user", content: userPrompt }
        ],
        temperature: parseFloat(AppState.config.temperature) || 0.7,
        stream: true
      })
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`API 响应错误 (${response.status}): ${errText}`);
    }

    statusIndicator.textContent = "生成中...";

    const reader = response.body.getReader();
    const decoder = new TextDecoder("utf-8");
    let buffer = "";

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop();

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || !trimmed.startsWith("data:")) continue;

        const dataStr = trimmed.slice(5).trim();
        if (dataStr === "[DONE]") break;

        try {
          const parsed = JSON.parse(dataStr);
          const delta = parsed.choices?.[0]?.delta?.content || "";
          if (delta) {
            accumulatedRaw += delta;
            handleStreamChunk(accumulatedRaw);
          }
        } catch (e) {
          // ignore chunk parse
        }
      }
    }

    // 生成完成处理
    finishGeneration(accumulatedRaw, isMultiDraft, userPrompt);

  } catch (error) {
    if (error.name === "AbortError") {
      statusIndicator.textContent = "已停止生成";
      showToast("已停止生成");
    } else {
      console.error("AI 生成异常:", error);
      statusIndicator.textContent = "生成失败";
      resultContent.innerHTML = `<div class="result-error">生成失败：${escapeHtml(error.message)}</div>`;
      showToast("生成失败，请检查设置与网络");
    }
  } finally {
    AppState.isGenerating = false;
    stopBtn.style.display = "none";
    resultContent.classList.remove("is-streaming");
  }
}

/**
 * 实时流式解析：提取 DeepSeek-R1 <think> 与 正文
 */
function handleStreamChunk(fullText) {
  const reasoningPanel = document.getElementById("reasoningPanel");
  const reasoningBody = document.getElementById("reasoningBody");
  const reasoningStatusText = document.getElementById("reasoningStatusText");

  // 检查是否包含 <think>
  if (fullText.includes("<think>")) {
    reasoningPanel.style.display = "block";
    const parts = fullText.split("</think>");

    if (parts.length > 1) {
      // 思考已闭合
      reasoningStatusText.textContent = "思考完成";
      reasoningBody.textContent = parts[0].replace("<think>", "").trim();
      const actualContent = parts[1].trim();
      renderMarkdownResult(actualContent);
    } else {
      // 正在思考中
      reasoningStatusText.textContent = "深度思考中...";
      reasoningBody.textContent = fullText.replace("<think>", "").trim();
      document.getElementById("aiResultContent").innerHTML = `<div class="loading-line"><span class="dots"><i></i><i></i><i></i></span>模型正在思考，正文即将输出…</div>`;
    }
  } else {
    // 普通模型无 <think>
    renderMarkdownResult(fullText);
  }
}

/**
 * 完成生成：多版本解析与历史记录入库
 */
function finishGeneration(rawOutput, isMultiDraft, userPrompt) {
  const statusIndicator = document.getElementById("generationStatus");
  statusIndicator.textContent = "生成完成";

  // 剔除 <think> 部分只留最终正文
  let cleanText = rawOutput;
  if (rawOutput.includes("</think>")) {
    cleanText = rawOutput.split("</think>")[1].trim();
  } else {
    cleanText = rawOutput.replace(/<think>[\s\S]*?<\/think>/g, "").trim();
  }

  // 多版本解析
  if (isMultiDraft) {
    const drafts = [];
    const v1Match = cleanText.match(/===版本1[^=]*===([\s\S]*?)(?:===版本2|$)/);
    const v2Match = cleanText.match(/===版本2[^=]*===([\s\S]*?)(?:===版本3|$)/);
    const v3Match = cleanText.match(/===版本3[^=]*===([\s\S]*?)$/);

    if (v1Match) drafts.push(v1Match[1].trim());
    if (v2Match) drafts.push(v2Match[1].trim());
    if (v3Match) drafts.push(v3Match[1].trim());

    if (drafts.length >= 2) {
      AppState.drafts = drafts;
      AppState.activeDraftIndex = 0;
      document.getElementById("draftTabsContainer").style.display = "block";
      switchDraft(0);
    } else {
      AppState.lastGeneratedText = cleanText;
      renderMarkdownResult(cleanText);
    }
  } else {
    AppState.lastGeneratedText = cleanText;
    renderMarkdownResult(cleanText);
  }

  // 记录到历史记录
  addHistoryRecord({
    prompt: userPrompt.substring(0, 80),
    result: getCurrentActiveOutputText(),
    mode: AppState.activeMode
  });
}

function renderMarkdownResult(text) {
  const resultContent = document.getElementById("aiResultContent");
  if (window.marked) {
    resultContent.innerHTML = marked.parse(text);
  } else {
    resultContent.textContent = text;
  }
}

function stopAIGeneration() {
  if (AppState.abortController) {
    AppState.abortController.abort();
    AppState.abortController = null;
  }
  AppState.isGenerating = false;
}

// ==========================================================================
// 历史记录管理器
// ==========================================================================

function loadHistoryFromStorage() {
  try {
    const saved = localStorage.getItem("word_ai_copilot_history");
    if (saved) {
      AppState.history = JSON.parse(saved);
    }
  } catch (e) {
    AppState.history = [];
  }
}

function addHistoryRecord(record) {
  const item = {
    id: Date.now().toString(),
    time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
    ...record
  };
  AppState.history.unshift(item);
  if (AppState.history.length > 30) {
    AppState.history.pop();
  }
  try {
    localStorage.setItem("word_ai_copilot_history", JSON.stringify(AppState.history));
  } catch (e) {}
}

function openHistoryDrawer() {
  renderHistoryList();
  document.getElementById("historyDrawer").style.display = "flex";
}

function closeHistoryDrawer() {
  document.getElementById("historyDrawer").style.display = "none";
}

function clearHistory() {
  if (confirm("确定要清空所有生成历史记录吗？")) {
    AppState.history = [];
    localStorage.removeItem("word_ai_copilot_history");
    renderHistoryList();
    showToast("历史记录已清空");
  }
}

function renderHistoryList() {
  const container = document.getElementById("historyListContainer");
  if (AppState.history.length === 0) {
    container.innerHTML = `<div class="empty-state"><svg class="ic" width="32" height="32"><use href="#i-history"/></svg>还没有历史记录<br>处理过的内容会自动保存在这里</div>`;
    return;
  }

  container.innerHTML = AppState.history.map(item => `
    <div class="history-item">
      <div class="history-item-header">
        <span>${item.mode === "document" ? "📑 全文审阅" : "📝 选区处理"}</span>
        <span>${item.time}</span>
      </div>
      <div class="history-item-prompt">${escapeHtml(item.prompt)}</div>
      <div class="history-item-body">${escapeHtml(item.result)}</div>
      <div class="history-item-actions">
        <button class="btn-text" onclick="applyHistoryItem('${item.id}')">载入并写回</button>
        <button class="btn-text" onclick="copyHistoryItem('${item.id}')">复制</button>
      </div>
    </div>
  `).join("");
}

window.applyHistoryItem = function(id) {
  const item = AppState.history.find(h => h.id === id);
  if (!item) return;
  AppState.lastGeneratedText = item.result;
  document.getElementById("resultSection").style.display = "block";
  renderMarkdownResult(item.result);
  closeHistoryDrawer();
  showToast("已载入该条历史记录，可直接点击写回");
};

window.copyHistoryItem = function(id) {
  const item = AppState.history.find(h => h.id === id);
  if (!item) return;
  navigator.clipboard.writeText(item.result).then(() => {
    showToast("已复制到剪贴板！");
  });
};

// ==========================================================================
// Prompt 预设收藏库管理器
// ==========================================================================

function loadPromptsFromStorage() {
  try {
    const saved = localStorage.getItem("word_ai_copilot_prompts");
    if (saved) {
      AppState.promptPresets = JSON.parse(saved);
    } else {
      AppState.promptPresets = DEFAULT_PROMPT_PRESETS;
    }
  } catch (e) {
    AppState.promptPresets = DEFAULT_PROMPT_PRESETS;
  }
  renderPromptChips();
}

function savePromptsToStorage() {
  try {
    localStorage.setItem("word_ai_copilot_prompts", JSON.stringify(AppState.promptPresets));
  } catch (e) {}
  renderPromptChips();
}

function renderPromptChips() {
  const container = document.getElementById("promptChipsList");
  container.innerHTML = AppState.promptPresets.slice(0, 5).map(p => `
    <button class="chip" data-text="${escapeHtml(p.content)}">${escapeHtml(p.title)}</button>
  `).join("");

  // 绑定点击
  container.querySelectorAll(".chip").forEach(chip => {
    chip.onclick = () => {
      const input = document.getElementById("customPromptInput");
      input.value = chip.getAttribute("data-text");
      input.focus();
    };
  });
}

function openPromptsDrawer() {
  renderPromptsDrawerList();
  document.getElementById("promptsDrawer").style.display = "flex";
}

function closePromptsDrawer() {
  document.getElementById("promptsDrawer").style.display = "none";
}

function renderPromptsDrawerList() {
  const container = document.getElementById("promptsPresetList");
  container.innerHTML = AppState.promptPresets.map(p => `
    <div class="preset-item" onclick="selectPresetPrompt('${p.id}')">
      <div>
        <div class="preset-title">${escapeHtml(p.title)}</div>
        <div class="preset-desc">${escapeHtml(p.content)}</div>
      </div>
      <button class="btn-text danger" onclick="event.stopPropagation(); deletePresetPrompt('${p.id}')">删除</button>
    </div>
  `).join("");
}

window.selectPresetPrompt = function(id) {
  const p = AppState.promptPresets.find(x => x.id === id);
  if (!p) return;
  document.getElementById("customPromptInput").value = p.content;
  closePromptsDrawer();
  document.getElementById("customPromptInput").focus();
};

window.deletePresetPrompt = function(id) {
  AppState.promptPresets = AppState.promptPresets.filter(x => x.id !== id);
  savePromptsToStorage();
  renderPromptsDrawerList();
  showToast("已从收藏库中删除");
};

function addNewPromptPreset() {
  const titleInput = document.getElementById("newPromptTitleInput");
  const contentInput = document.getElementById("newPromptContentInput");
  const title = titleInput.value.trim();
  const content = contentInput.value.trim();

  if (!title || !content) {
    showToast("请填写标题和详细内容！");
    return;
  }

  AppState.promptPresets.push({
    id: Date.now().toString(),
    title,
    content
  });
  savePromptsToStorage();
  renderPromptsDrawerList();
  titleInput.value = "";
  contentInput.value = "";
  showToast("已成功添加到 Prompt 收藏库！");
}

// ==========================================================================
// 设置与连通性测试
// ==========================================================================

function loadConfigFromStorage() {
  try {
    const saved = localStorage.getItem("word_ai_copilot_config");
    if (saved) {
      AppState.config = { ...AppState.config, ...JSON.parse(saved) };
    }
  } catch (e) {}
  updateHeaderModelBadge();
}

function saveConfigToStorage() {
  try {
    localStorage.setItem("word_ai_copilot_config", JSON.stringify(AppState.config));
    updateHeaderModelBadge();
  } catch (e) {}
}

function updateHeaderModelBadge() {
  const badge = document.getElementById("currentModelBadge");
  if (badge) {
    badge.textContent = AppState.config.model || "未配置模型";
    badge.title = `${AppState.config.baseUrl} (${AppState.config.model})`;
  }
}

function openSettingsDrawer() {
  const drawer = document.getElementById("settingsDrawer");
  document.getElementById("providerPreset").value = AppState.config.provider || "local_custom";
  document.getElementById("apiUrlInput").value = AppState.config.baseUrl || "";
  document.getElementById("apiKeyInput").value = AppState.config.apiKey || "";
  document.getElementById("modelNameInput").value = AppState.config.model || "";
  document.getElementById("temperatureSlider").value = AppState.config.temperature || 0.7;
  document.getElementById("temperatureVal").textContent = AppState.config.temperature || 0.7;
  document.getElementById("systemPromptInput").value = AppState.config.systemPrompt || "";
  document.getElementById("testConnectionStatus").textContent = "";

  drawer.style.display = "flex";
}

function closeSettingsDrawer() {
  document.getElementById("settingsDrawer").style.display = "none";
}

function applyProviderPreset(providerKey) {
  const preset = PROVIDER_PRESETS[providerKey];
  if (!preset) return;
  if (preset.baseUrl) document.getElementById("apiUrlInput").value = preset.baseUrl;
  if (preset.model) document.getElementById("modelNameInput").value = preset.model;
}

function saveSettingsFromDrawer() {
  AppState.config.provider = document.getElementById("providerPreset").value;
  AppState.config.baseUrl = document.getElementById("apiUrlInput").value.trim();
  AppState.config.apiKey = document.getElementById("apiKeyInput").value.trim();
  AppState.config.model = document.getElementById("modelNameInput").value.trim();
  AppState.config.temperature = parseFloat(document.getElementById("temperatureSlider").value);
  AppState.config.systemPrompt = document.getElementById("systemPromptInput").value.trim();

  saveConfigToStorage();
  closeSettingsDrawer();
  showToast("设置已保存并生效！");
}

async function testAPIConnection() {
  const statusEl = document.getElementById("testConnectionStatus");
  statusEl.className = "test-status";
  statusEl.textContent = "正在连接中...";

  const baseUrl = document.getElementById("apiUrlInput").value.trim();
  const apiKey = document.getElementById("apiKeyInput").value.trim();
  const model = document.getElementById("modelNameInput").value.trim();

  if (!baseUrl) {
    statusEl.className = "test-status error";
    statusEl.textContent = "请先填写 Base URL";
    return;
  }

  if (!model) {
    statusEl.className = "test-status error";
    statusEl.textContent = "请填写您加载的模型名 (Model)";
    return;
  }

  try {
    const rawEndpoint = `${baseUrl.replace(/\/+$/, "")}/chat/completions`;
    const defaultHeaders = {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey || "dummy-key"}`
    };
    const { fetchUrl, fetchHeaders } = resolveFetchEndpointAndHeaders(rawEndpoint, defaultHeaders);

    const res = await fetch(fetchUrl, {
      method: "POST",
      headers: fetchHeaders,
      body: JSON.stringify({
        model: model,
        messages: [{ role: "user", content: "hi" }],
        max_tokens: 5
      })
    });

    if (res.ok) {
      statusEl.className = "test-status success";
      statusEl.textContent = "✅ 连接成功！";
    } else {
      const err = await res.text();
      statusEl.className = "test-status error";
      statusEl.textContent = `❌ 失败 (${res.status}): ${err.substring(0, 50)}`;
    }
  } catch (e) {
    statusEl.className = "test-status error";
    statusEl.textContent = `❌ 异常: ${e.message}`;
  }
}

// 辅助方法
let toastTimer = null;
function showToast(message) {
  const toast = document.getElementById("toastMessage");
  toast.textContent = message;
  toast.classList.add("show");

  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.remove("show");
  }, 2400);
}

function escapeHtml(str) {
  if (!str) return "";
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
