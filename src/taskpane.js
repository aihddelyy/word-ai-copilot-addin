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
  },
  continue_write: {
    name: "智能往下续写",
    prompt: "请根据以下上文语境，顺着原有的文风、逻辑脉络与核心主题，自然流畅地往下续写下一段或后续内容。注意：请直接输出新增的续写内容本身，严禁重复上文已有的句子：\n\n"
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
    prompt: "请作为资深文档主编，对整篇文档进行全方位审阅，输出一份专业的【文档审查与修改意见报告】。重点排查：1. 错别字与语病；2. 专业术语前后不统一；3. 论据或逻辑前后矛盾；4. 排版与标点规范性。\n\n【重要格式要求】：在列出具体文字修改建议时，请务必使用直角引号「」标出文档中的原句片段与修改后的句子，严格采用如下格式：\n- 原文：「此处为文档中的原始句子片段」 -> 建议改为：「此处为修改后的句子」（原因说明）\n以便系统自动为用户提供一键定位与替换功能。\n\n待审阅全文如下：\n\n"
  },
  doc_table: {
    name: "整理关键数据表格",
    prompt: "请提炼以下文档中的核心数据、指标、对比项或关键信息，整理为一个或多个结构清晰的标准 Markdown 表格（包含表头、对齐线和具体数据）：\n\n"
  }
};

// 默认内置的 Prompt 收藏项（覆盖公文写作、论文规范、总结提炼、降重润色等常用场景）
const DEFAULT_PROMPT_PRESETS = [
  { id: "p1", title: "公文结构改写", content: "请将以下内容改写为标准体制内公文汇报材料风格：讲政治、站位高、逻辑严密、层次清晰，标题对仗工整，采用一、（一）1. 的公文分层规范：", showInChips: true },
  { id: "p2", title: "论文降重与改写", content: "请在严格保留以下学术段落原意和关键技术术语的前提下，进行深度重构降重：调整句式语态（主动与被动互换）、同义替换非关键修饰词、拆分长难句，彻底消除查重痕迹：", showInChips: true },
  { id: "p3", title: "整理对比三线表", content: "请提炼以下内容的核心指标与对比维度，整理为规范的标准 Markdown 三线对比表格（包含表头、对齐方式及精炼数据）：", showInChips: true },
  { id: "p4", title: "核心要点加粗", content: "请对以下内容进行逻辑梳理，提炼为清晰的无序要点列表，并将每个要点中的关键专业术语、核心数据、时间节点与结论用 **加粗** 重点标出：", showInChips: true },
  { id: "p5", title: "消除套话与口语", content: "请对以下内容进行深度精炼，删除文中空洞的套话、废话、口头禅以及冗长的修饰词，使文风紧凑干练、字字珠玑：", showInChips: true },
  { id: "p6", title: "会议纪要待办化", content: "请将以下零散的会议记录或研讨发言内容整理为规范的【会议纪要与行动清单】，包含：1. 会议核心共识；2. 关键决议事项；3. 具体待办任务（分派角色、事项、截止要求）：", showInChips: false },
  { id: "p7", title: "领导致辞起草", content: "请根据以下核心主题与要点，起草一篇富有感染力与思想深度的领导讲话致辞，结构包含热情洋溢的开篇、高屋建瓴的形势回顾、求真务实的部署要求与催人奋进的号召结语：", showInChips: false },
  { id: "p8", title: "SCI 论文学术润色", content: "Please polish the following academic text to meet the high standards of top-tier international journals (e.g., Nature, IEEE). Improve lexical variety, use precise academic terminology, ensure cohesive transitions, and eliminate colloquial expressions:", showInChips: false },
  { id: "p9", title: "金字塔逻辑汇报", content: "请按照麦肯锡金字塔原理（结论先行、以上统下、归类分组、逻辑递进），将以下散乱内容重构成条理清晰的高管汇报陈述：", showInChips: false },
  { id: "p10", title: "法律免责与合规条款", content: "请为以下商业或业务说明内容补充严密周全的法律风险提示、合规边界及免责声明条款：", showInChips: false }
];

// 全局应用状态
const AppState = {
  isOfficeInitialized: false,
  activeMode: "selection", // "selection" | "document"
  currentSelection: "",
  surroundingContext: "", // 选区所在的前后段落上下文
  fullDocText: "",
  lastGeneratedText: "",
  lastUserPrompt: "",
  lastActionType: "", // 记录最近执行的动作类型（如 doc_audit / continue_write 等）
  // 划选局部写回状态
  selectedResultText: "",
  // 多版本草稿状态
  drafts: [], // Array of string
  activeDraftIndex: 0,
  // 思考过程状态
  reasoningText: "",
  isReasoning: false,
  isReasoningCollapsed: false,
  // 缓存的可用模型列表
  cachedModels: [],
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
    systemPrompt: "你是一位专业的 Microsoft Word 文档创作助手与资深文字编辑。请根据用户的指令直接输出高质量处理结果。支持 Markdown 排版（加粗、列表、表格、标题）。默认只输出一份最终成稿正文，严禁输出寒暄客套、解释说明或未要求的多个备选方案，确保内容可直接写回文档。"
  },
  // 历史记录与收藏
  history: [],
  promptPresets: [],
  editingPresetId: null
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

let selectionDebounceTimer = null;

async function registerSelectionChangeListener() {
  if (!AppState.isOfficeInitialized) return;

  const onSelectionEvent = () => {
    // 若用户当前不在“选中文字”模式或 AI 正在流式生成中，则不打断
    if (AppState.activeMode !== "selection" || AppState.isGenerating) return;

    if (selectionDebounceTimer) clearTimeout(selectionDebounceTimer);
    selectionDebounceTimer = setTimeout(() => {
      refreshDocumentSelection(true);
    }, 280);
  };

  try {
    // 1. 优先使用 Office.EventType.DocumentSelectionChanged (最稳定、支持各版本 Word)
    if (window.Office && Office.context && Office.context.document && Office.context.document.addHandlerAsync) {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSelectionEvent,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("已成功挂载 Office.EventType.DocumentSelectionChanged 监听");
          }
        }
      );
    }

    // 2. 结合 Word.run context.document.onSelectionChanged 双保险
    await Word.run(async (context) => {
      if (context.document.onSelectionChanged && typeof context.document.onSelectionChanged.add === "function") {
        context.document.onSelectionChanged.add(onSelectionEvent);
        await context.sync();
      }
    });
  } catch (error) {
    console.warn("注册选区监听异常:", error);
  }
}

async function refreshDocumentSelection(isAutoSync = false) {
  if (!AppState.isOfficeInitialized) {
    return;
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      const paras = selection.paragraphs;
      paras.load("items/text");
      await context.sync();

      const text = selection.text ? selection.text.trim() : "";

      // 若文字完全相同则跳过重新计算上下文，避免多余性能开销
      if (isAutoSync && text === AppState.currentSelection) {
        return;
      }

      AppState.currentSelection = text;

      // 提取选区所在段落及前后相邻段落作为上下文语境
      let ctxParts = [];
      if (paras.items && paras.items.length > 0) {
        const firstPara = paras.items[0];
        const lastPara = paras.items[paras.items.length - 1];
        const prevPara = firstPara.getPreviousOrNullObject();
        const nextPara = lastPara.getNextOrNullObject();
        prevPara.load("text");
        nextPara.load("text");
        await context.sync();

        if (!prevPara.isNullObject && prevPara.text && prevPara.text.trim()) {
          ctxParts.push(prevPara.text.trim());
        }
        paras.items.forEach(p => {
          if (p.text && p.text.trim()) ctxParts.push(p.text.trim());
        });
        if (!nextPara.isNullObject && nextPara.text && nextPara.text.trim()) {
          ctxParts.push(nextPara.text.trim());
        }
      }
      AppState.surroundingContext = ctxParts.join("\n").substring(0, 1200);
      updateSelectionUI(text);
    });
  } catch (error) {
    console.error("读取选区失败:", error);
  }
}

/**
 * 获取光标或选区前方的上文段落（供“智能往下续写”使用）
 */
async function getCursorPrecedingText() {
  if (!AppState.isOfficeInitialized) {
    return AppState.currentSelection || AppState.fullDocText || "数字化转型正在深入推进，企业亟需构建智能化办公体系。";
  }

  try {
    let precedingText = "";
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      const paras = selection.paragraphs;
      paras.load("items/text");
      await context.sync();

      if (selection.text && selection.text.trim().length > 20) {
        precedingText = selection.text.trim();
        return;
      }

      const chunks = [];
      if (paras.items && paras.items.length > 0) {
        const curPara = paras.items[paras.items.length - 1];
        const p1 = curPara.getPreviousOrNullObject();
        p1.load("text");
        await context.sync();

        if (!p1.isNullObject) {
          const p2 = p1.getPreviousOrNullObject();
          p2.load("text");
          await context.sync();
          if (!p2.isNullObject && p2.text && p2.text.trim()) {
            chunks.push(p2.text.trim());
          }
          if (p1.text && p1.text.trim()) {
            chunks.push(p1.text.trim());
          }
        }
        if (curPara.text && curPara.text.trim()) {
          chunks.push(curPara.text.trim());
        }
      }
      precedingText = chunks.join("\n\n").trim();
    });
    return precedingText;
  } catch (error) {
    console.warn("获取光标上文失败:", error);
    return AppState.currentSelection || "";
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
 * 将 Markdown 转化为符合 Word 样式的标准 HTML（支持注入原文档字体与字号保真）
 */
function markdownToWordHtml(markdownText, inheritedStyle = null) {
  let rawHtml = "";
  if (window.marked) {
    rawHtml = marked.parse(markdownText);
  } else {
    rawHtml = `<p>${markdownText.replace(/\n/g, "<br>")}</p>`;
  }

  // 为表格注入优雅的 Word 原生支持样式
  const tableFont = (inheritedStyle && inheritedStyle.fontName) ? inheritedStyle.fontName : "Segoe UI, Microsoft YaHei, sans-serif";
  rawHtml = rawHtml.replace(/<table>/g, `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #94a3b8; font-family: ${tableFont};">`);
  rawHtml = rawHtml.replace(/<th>/g, '<th style="background-color: #f1f5f9; font-weight: bold; border: 1px solid #94a3b8; padding: 6px 10px;">');
  rawHtml = rawHtml.replace(/<td>/g, '<td style="border: 1px solid #94a3b8; padding: 6px 10px;">');

  // 若启用了原文档字体与字号保真，包裹外层内联样式
  if (inheritedStyle && (inheritedStyle.fontName || inheritedStyle.fontSize)) {
    const styleParts = [];
    if (inheritedStyle.fontName) styleParts.push(`font-family: '${inheritedStyle.fontName}', sans-serif`);
    if (inheritedStyle.fontSize) styleParts.push(`font-size: ${inheritedStyle.fontSize}pt`);
    if (inheritedStyle.fontColor) styleParts.push(`color: ${inheritedStyle.fontColor}`);
    rawHtml = `<div style="${styleParts.join("; ")}">${rawHtml}</div>`;
  }

  return rawHtml;
}

/**
 * 原生富文本替换选中文本（支持公文/论文原字体与缩进保真）
 */
async function replaceSelectionWithRichText(markdownText) {
  if (!AppState.isOfficeInitialized) {
    showToast("测试模式：已模拟富文本替换");
    AppState.currentSelection = markdownText;
    updateSelectionUI(markdownText);
    return;
  }

  const preserveStyle = document.getElementById("preserveFontStyleCheck")?.checked;

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      let inherited = null;
      let paraProps = null;

      if (preserveStyle) {
        try {
          selection.font.load("name,size,color");
          const firstPara = selection.paragraphs.getFirstOrNullObject();
          firstPara.load("firstLineIndent,lineSpacing,alignment");
          await context.sync();

          inherited = {
            fontName: selection.font.name || "",
            fontSize: selection.font.size || 0,
            fontColor: selection.font.color || ""
          };
          if (!firstPara.isNullObject) {
            paraProps = {
              firstLineIndent: firstPara.firstLineIndent,
              lineSpacing: firstPara.lineSpacing,
              alignment: firstPara.alignment
            };
          }
        } catch (styleErr) {
          console.warn("读取原选区字体样式跳过:", styleErr);
        }
      }

      const wordHtml = markdownToWordHtml(markdownText, inherited);
      const insertedRange = selection.insertHtml(wordHtml, Word.InsertLocation.replace);

      // 写回后二次加固原字体名与字号（防止 Word HTML 解析覆盖公文字体）
      if (preserveStyle && inherited) {
        if (inherited.fontName) insertedRange.font.name = inherited.fontName;
        if (inherited.fontSize && inherited.fontSize > 0) insertedRange.font.size = inherited.fontSize;
      }
      if (preserveStyle && paraProps && !markdownText.includes("|")) {
        const insertedParas = insertedRange.paragraphs;
        insertedParas.load("items");
        await context.sync();
        insertedParas.items.forEach(p => {
          if (typeof paraProps.firstLineIndent === "number") p.firstLineIndent = paraProps.firstLineIndent;
          if (typeof paraProps.lineSpacing === "number" && paraProps.lineSpacing > 0) p.lineSpacing = paraProps.lineSpacing;
        });
      }

      await context.sync();
      showToast(preserveStyle ? "已替换并保留原文档字体排版！" : "已按 Word 原生富文本格式替换成功！");
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

/**
 * 将 AI 结果作为 Word 原生批注气泡插入 (WordApi 1.4+)
 */
async function insertAsWordComment(commentText) {
  if (!AppState.isOfficeInitialized) {
    showToast("测试模式：已作为 Word 批注气泡插入！");
    return;
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      if (typeof selection.insertComment === "function") {
        selection.insertComment(commentText);
        await context.sync();
        showToast("已成功在右侧插入 Word 原生批注！");
      } else {
        // 兼容旧版 Word：在选区后方插入醒目批注标记
        const fallbackRange = selection.insertText(` 【AI批注：${commentText}】`, Word.InsertLocation.after);
        fallbackRange.font.color = "#d97706";
        fallbackRange.font.bold = true;
        await context.sync();
        showToast("当前 Word 版本不支持原生气泡，已插入内联批注标记");
      }
    });
  } catch (error) {
    console.error("插入批注失败:", error);
    showToast("插入批注失败（请确保已在文档中选中要批注的词句）");
  }
}

/**
 * 以审阅红线模式插入 Word 正文（字词级划线：红色删除线表示原句，绿色下划线表示 AI 润色新句，纯划线无底色）
 */
async function insertAsTrackedChanges(newText) {
  if (!AppState.isOfficeInitialized) {
    showToast("测试模式：已模拟以审阅红线插入正文！");
    return;
  }

  const oldText = AppState.currentSelection;
  if (!oldText || oldText.trim().length === 0) {
    showToast("请先在 Word 文档中选中需要审阅修改的原句");
    return;
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("font/name, font/size");
      await context.sync();

      const fontName = selection.font.name || "Calibri";
      const fontSize = selection.font.size || 12;

      // 使用精密 LCS 算法计算字词级差异
      const diffResult = computeInlineDiffHtml(oldText, newText);

      // 将 LCS 生成的 del 和 ins 转换为 Word 专用的内联富文本样式（只保留纯划线/下划线，绝不添加背景底色）
      let reviewHtml = diffResult.html
        .replace(/class="diff-del"/g, 'style="color:#c00000; text-decoration:line-through; font-family:\'' + fontName + '\'; font-size:' + fontSize + 'pt;"')
        .replace(/class="diff-ins"/g, 'style="color:#008000; text-decoration:underline; font-weight:500; font-family:\'' + fontName + '\'; font-size:' + fontSize + 'pt;"');

      // 若没有直接的字词合并（例如全文重写），则输出标准红线块
      if (!diffResult.html.includes("<del") && !diffResult.html.includes("<ins")) {
        reviewHtml = `
          <del style="color:#c00000; text-decoration:line-through; font-family:'${fontName}'; font-size:${fontSize}pt;">${escapeHtml(oldText)}</del>
          <ins style="color:#008000; text-decoration:underline; font-weight:500; font-family:'${fontName}'; font-size:${fontSize}pt;">${escapeHtml(newText)}</ins>
        `;
      }

      const wrapperHtml = `<span style="font-family:'${fontName}'; font-size:${fontSize}pt;">${reviewHtml}</span>`;
      selection.insertHtml(wrapperHtml, Word.InsertLocation.replace);
      await context.sync();

      showToast("已以审阅红线(逐字)模式插入正文！可在 Word 中直观核对修改");
    });
  } catch (error) {
    console.error("审阅红线插入失败:", error);
    showToast("审阅插入失败: " + error.message);
  }
}

/**
 * 以 Word 审阅方式逐段对照插入 Word 正文（一段对一段：原段落使用 <del> 红色删除线，新段落使用 <ins> 绿色下划线，纯划线无底色）
 */
async function insertAsParagraphDiff(newText) {
  if (!AppState.isOfficeInitialized) {
    showToast("测试模式：已模拟以 Word 审阅逐段对照插入正文！");
    return;
  }

  const oldText = AppState.currentSelection;
  if (!oldText || oldText.trim().length === 0) {
    showToast("请先在 Word 文档中选中需要对照的原段落");
    return;
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("font/name, font/size");
      await context.sync();

      const fontName = selection.font.name || "Calibri";
      const fontSize = selection.font.size || 12;

      // 分割为段落列表
      const oldParas = oldText.split(/\r?\n+/).map(p => p.trim()).filter(Boolean);
      const newParas = newText.split(/\r?\n+/).map(p => p.trim()).filter(Boolean);

      let htmlBlocks = [];
      const maxCount = Math.max(oldParas.length, newParas.length);

      for (let i = 0; i < maxCount; i++) {
        if (i < oldParas.length) {
          htmlBlocks.push(`
            <p style="line-height:1.6; margin:0 0 4pt 0; font-family:'${fontName}'; font-size:${fontSize}pt;">
              <del style="color:#c00000; text-decoration:line-through; font-family:'${fontName}'; font-size:${fontSize}pt;">${escapeHtml(oldParas[i])}</del>
            </p>
          `);
        }
        if (i < newParas.length) {
          const formattedPara = escapeHtml(newParas[i]).replace(/\*\*(.*?)\*\*/g, '<b>$1</b>');
          htmlBlocks.push(`
            <p style="line-height:1.6; margin:0 0 10pt 0; font-family:'${fontName}'; font-size:${fontSize}pt;">
              <ins style="color:#008000; text-decoration:underline; font-weight:500; font-family:'${fontName}'; font-size:${fontSize}pt;">${formattedPara}</ins>
            </p>
          `);
        }
      }

      const finalHtml = htmlBlocks.join("");
      selection.insertHtml(finalHtml, Word.InsertLocation.replace);
      await context.sync();

      showToast("已以 Word 审阅方式(逐段)插入正文！可在 Word 中逐段核对修改");
    });
  } catch (error) {
    console.error("逐段对照插入失败:", error);
    showToast("段落对照插入失败: " + error.message);
  }
}

/**
 * 在 Word 正文中搜索指定片段并滚动高亮选中（供校对审查报告联动使用）
 */
window.locateTextInWord = async function (targetSnippet) {
  if (!targetSnippet) return;
  if (!AppState.isOfficeInitialized) {
    showToast(`测试模式：已定位并高亮「${targetSnippet.substring(0, 15)}…」`);
    return;
  }

  try {
    await Word.run(async (context) => {
      // Word search API 单次搜索串长度限制在 255 字符以内
      const query = targetSnippet.trim().substring(0, 120);
      const results = context.document.body.search(query, { matchCase: false, ignoreSpace: true });
      results.load("items");
      await context.sync();

      if (results.items.length > 0) {
        results.items[0].select();
        await context.sync();
        showToast("📍 已在文档中定位并高亮该句！");
      } else {
        showToast("未能在文档中精确匹配到该原句（可能已被修改）");
      }
    });
  } catch (error) {
    console.error("定位原文失败:", error);
    showToast("定位失败: " + error.message);
  }
};

/**
 * 在 Word 正文中搜索指定原句并一键替换为修改建议
 */
window.replaceSpecificTextInWord = async function (targetSnippet, replacementText, btnEl) {
  if (!targetSnippet || !replacementText) return;
  if (!AppState.isOfficeInitialized) {
    if (btnEl) {
      btnEl.textContent = "✅ 已修正";
      btnEl.classList.add("done");
    }
    showToast(`测试模式：已将「${targetSnippet.substring(0, 10)}…」替换为「${replacementText.substring(0, 10)}…」`);
    return;
  }

  try {
    await Word.run(async (context) => {
      const query = targetSnippet.trim().substring(0, 120);
      const results = context.document.body.search(query, { matchCase: false, ignoreSpace: true });
      results.load("items");
      await context.sync();

      if (results.items.length > 0) {
        const targetRange = results.items[0];
        targetRange.insertText(replacementText, Word.InsertLocation.replace);
        targetRange.select();
        await context.sync();
        if (btnEl) {
          btnEl.textContent = "✅ 已修正";
          btnEl.classList.add("done");
        }
        showToast("⚡ 已成功在文档中定位并修正该处！");
      } else {
        showToast("未能在文档中找到该原句（可能已被修改）");
      }
    });
  } catch (error) {
    console.error("一键修正失败:", error);
    showToast("修正失败: " + error.message);
  }
};

// ==========================================================================
// UI 状态与事件绑定
// ==========================================================================

function updateSelectionUI(text) {
  const previewEl = document.getElementById("selectionPreview");
  const countEl = document.getElementById("selectionWordCount");

  if (!text || text.length === 0) {
    if (previewEl) {
      previewEl.textContent = "在 Word 文档中选中文字后，此处将自动同步预览...";
      previewEl.classList.add("empty");
    }
    if (countEl) countEl.textContent = "0 字";
  } else {
    if (previewEl) {
      previewEl.textContent = text;
      previewEl.classList.remove("empty");
    }
    if (countEl) countEl.textContent = `${text.length} 字`;
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

  // 选中文本快捷动作与迷你条芯片动作
  document.querySelectorAll(".action-btn, .mini-chip").forEach((btn) => {
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

  const btnTrackChanges = document.getElementById("btnTrackChangesReview");
  if (btnTrackChanges) {
    btnTrackChanges.onclick = () => {
      const text = getCurrentActiveOutputText();
      if (text) insertAsTrackedChanges(text);
    };
  }

  const btnParagraphDiff = document.getElementById("btnParagraphDiffReview");
  if (btnParagraphDiff) {
    btnParagraphDiff.onclick = () => {
      const text = getCurrentActiveOutputText();
      if (text) insertAsParagraphDiff(text);
    };
  }

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

  const btnInsertComment = document.getElementById("btnInsertComment");
  if (btnInsertComment) {
    btnInsertComment.onclick = () => {
      const text = getCurrentActiveOutputText();
      if (text) insertAsWordComment(text);
    };
  }

  document.getElementById("btnCopyResult").onclick = () => {
    const text = getCurrentActiveOutputText();
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
      showToast("已成功复制到剪贴板！");
    });
  };

  // 二次微调芯片与追问输入框
  document.querySelectorAll(".refine-chip").forEach(chip => {
    chip.onclick = () => {
      const refineKey = chip.getAttribute("data-refine");
      handleRefinement(refineKey);
    };
  });
  const btnRefineSubmit = document.getElementById("btnRefineSubmit");
  const refineInput = document.getElementById("refineInput");
  if (btnRefineSubmit && refineInput) {
    btnRefineSubmit.onclick = () => {
      const val = refineInput.value.trim();
      if (val) {
        handleRefinement(val, true);
        refineInput.value = "";
      } else {
        showToast("请输入您的微调修改意见！");
      }
    };
    refineInput.onkeydown = (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        btnRefineSubmit.click();
      }
    };
  }

  // 顶部模型徽章下拉秒切
  const modelBadgeBtn = document.getElementById("currentModelBadge");
  if (modelBadgeBtn) {
    modelBadgeBtn.onclick = (e) => {
      e.stopPropagation();
      toggleQuickModelDropdown();
    };
  }
  document.addEventListener("click", (e) => {
    const dropdown = document.getElementById("quickModelDropdown");
    if (dropdown && dropdown.style.display !== "none" && !dropdown.contains(e.target)) {
      dropdown.style.display = "none";
    }
    // 全局点击关闭设置抽屉里的自定义模型下拉
    const customMenu = document.getElementById("customModelMenu");
    const customBtn = document.getElementById("customModelSelectBtn");
    if (customMenu && customMenu.style.display !== "none" && !customMenu.contains(e.target) && !customBtn.contains(e.target)) {
      customMenu.style.display = "none";
      if (customBtn) customBtn.classList.remove("active");
    }
  });

  // 差异对比（支持字词级红绿高亮合并对比 & 上下分栏对比）
  document.getElementById("btnToggleDiff").onclick = () => {
    const diffContainer = document.getElementById("diffContainer");
    const isHidden = diffContainer.style.display === "none";
    diffContainer.style.display = isHidden ? "flex" : "none";
    document.getElementById("btnToggleDiff").textContent = isHidden ? "收起对比 (Diff)" : "查看与原文红绿字级对比 (Diff)";

    if (isHidden) {
      refreshDiffView();
    }
  };

  const btnDiffInline = document.getElementById("btnDiffModeInline");
  const btnDiffSplit = document.getElementById("btnDiffModeSplit");
  if (btnDiffInline && btnDiffSplit) {
    btnDiffInline.onclick = () => {
      btnDiffInline.classList.add("active");
      btnDiffSplit.classList.remove("active");
      document.getElementById("diffInlineBox").style.display = "block";
      document.getElementById("diffSplitBox").style.display = "none";
    };
    btnDiffSplit.onclick = () => {
      btnDiffSplit.classList.add("active");
      btnDiffInline.classList.remove("active");
      document.getElementById("diffInlineBox").style.display = "none";
      document.getElementById("diffSplitBox").style.display = "flex";
    };
  }

  // 历史记录抽屉
  document.getElementById("btnOpenHistory").onclick = openHistoryDrawer;
  document.getElementById("btnCloseHistory").onclick = closeHistoryDrawer;
  document.getElementById("btnClearHistory").onclick = clearHistory;

  // Prompt 收藏抽屉
  document.getElementById("btnOpenPrompts").onclick = openPromptsDrawer;
  document.getElementById("btnClosePrompts").onclick = closePromptsDrawer;
  document.getElementById("btnAddPromptPreset").onclick = addNewPromptPreset;
  const btnExport = document.getElementById("btnExportPrompts");
  const btnImport = document.getElementById("btnImportPrompts");
  const fileInput = document.getElementById("promptFileInput");
  if (btnExport) btnExport.onclick = exportPromptsToFile;
  if (btnImport && fileInput) {
    btnImport.onclick = () => fileInput.click();
    fileInput.onchange = importPromptsFromFile;
  }

  // 设置抽屉
  document.getElementById("btnSettings").onclick = openSettingsDrawer;
  document.getElementById("btnCloseSettings").onclick = closeSettingsDrawer;
  document.getElementById("btnSaveSettings").onclick = saveSettingsFromDrawer;
  document.getElementById("providerPreset").onchange = (e) => applyProviderPreset(e.target.value);
  document.getElementById("btnTestConnection").onclick = testAPIConnection;
  document.getElementById("btnFetchModels").onclick = () => fetchAvailableModels(false);
  document.getElementById("modelSelect").onchange = (e) => {
    document.getElementById("modelNameInput").value = e.target.value;
  };

  // 自定义模型下拉按钮点击展开/收起
  const customModelBtn = document.getElementById("customModelSelectBtn");
  if (customModelBtn) {
    customModelBtn.onclick = (e) => {
      e.stopPropagation();
      toggleCustomModelMenu();
    };
  }

  document.getElementById("btnToggleManualModel").onclick = toggleManualModelInput;
  document.getElementById("apiUrlInput").onblur = () => {
    if (document.getElementById("apiUrlInput").value.trim()) {
      fetchAvailableModels(true);
    }
  };
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

  // 多版本 Tab 点击委托（事件代理，支持动态渲染的方案 Tab）
  const draftTabsList = document.getElementById("draftTabsList");
  if (draftTabsList) {
    draftTabsList.onclick = (e) => {
      const tab = e.target.closest(".draft-tab");
      if (tab) {
        const idx = parseInt(tab.getAttribute("data-index"), 10);
        switchDraft(idx);
      }
    };
  }
}

/**
 * 计算字/词级红绿高亮 Diff 视图
 */
function refreshDiffView() {
  const origText = (AppState.activeMode === "selection" ? AppState.currentSelection : AppState.fullDocText) || "";
  const newText = getCurrentActiveOutputText() || "";

  document.getElementById("diffOriginalText").textContent = origText || "（无原文字）";
  document.getElementById("diffNewText").textContent = newText || "（无生成结果）";

  const { html, addedCount, deletedCount } = computeInlineDiffHtml(origText, newText);
  const inlineBody = document.getElementById("diffInlineBody");
  const statsBadge = document.getElementById("diffStatsBadge");

  if (inlineBody) inlineBody.innerHTML = html;
  if (statsBadge) {
    statsBadge.innerHTML = `<span class="diff-stat-add">+新增 ${addedCount} 字</span><span class="diff-stat-del">-删除 ${deletedCount} 字</span>`;
  }
}

function tokenizeForDiff(str) {
  if (!str) return [];
  return str.match(/[\u4e00-\u9fa5]|[a-zA-Z0-9_]+|\s+|[^\s\w\u4e00-\u9fa5]/g) || [];
}

function computeInlineDiffHtml(oldStr, newStr) {
  const oldTokens = tokenizeForDiff(oldStr);
  const newTokens = tokenizeForDiff(newStr);

  // 针对极长文本做上限保护以免 LCS 矩阵过大
  if (oldTokens.length * newTokens.length > 400000) {
    return {
      html: `<del class="diff-del">${escapeHtml(oldStr)}</del>\n<ins class="diff-ins">${escapeHtml(newStr)}</ins>`,
      addedCount: newStr.length,
      deletedCount: oldStr.length
    };
  }

  const m = oldTokens.length;
  const n = newTokens.length;
  const dp = Array.from({ length: m + 1 }, () => new Uint16Array(n + 1));

  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      if (oldTokens[i] === newTokens[j]) {
        dp[i][j] = dp[i + 1][j + 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i + 1][j], dp[i][j + 1]);
      }
    }
  }

  let i = 0, j = 0;
  const ops = [];
  let addedCount = 0;
  let deletedCount = 0;

  while (i < m && j < n) {
    if (oldTokens[i] === newTokens[j]) {
      ops.push({ type: "eq", val: oldTokens[i] });
      i++; j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      ops.push({ type: "del", val: oldTokens[i] });
      deletedCount += oldTokens[i].trim().length;
      i++;
    } else {
      ops.push({ type: "ins", val: newTokens[j] });
      addedCount += newTokens[j].trim().length;
      j++;
    }
  }
  while (i < m) {
    ops.push({ type: "del", val: oldTokens[i] });
    deletedCount += oldTokens[i].trim().length;
    i++;
  }
  while (j < n) {
    ops.push({ type: "ins", val: newTokens[j] });
    addedCount += newTokens[j].trim().length;
    j++;
  }

  // 合并相邻同类标签
  let html = "";
  let idx = 0;
  while (idx < ops.length) {
    const curType = ops[idx].type;
    let buf = "";
    while (idx < ops.length && ops[idx].type === curType) {
      buf += ops[idx].val;
      idx++;
    }
    const escaped = escapeHtml(buf);
    if (curType === "eq") html += escaped;
    else if (curType === "del") html += `<del class="diff-del">${escaped}</del>`;
    else if (curType === "ins") html += `<ins class="diff-ins">${escaped}</ins>`;
  }

  return { html: html || "（无变化）", addedCount, deletedCount };
}

function getCurrentActiveOutputText() {
  // 1. 若用户在生成结果区域划选了局部文本，优先写回划选文本
  if (AppState.selectedResultText && AppState.selectedResultText.trim().length > 0) {
    return AppState.selectedResultText.trim();
  }
  // 2. 否则按当前激活的草稿或全局文本
  if (AppState.drafts.length > 0 && AppState.drafts[AppState.activeDraftIndex]) {
    return AppState.drafts[AppState.activeDraftIndex];
  }
  return AppState.lastGeneratedText;
}

function switchDraft(index) {
  AppState.activeDraftIndex = index;
  AppState.selectedResultText = ""; // 切换方案时重置局部划选
  document.querySelectorAll(".draft-tab").forEach(tab => {
    const i = parseInt(tab.getAttribute("data-index"), 10);
    tab.classList.toggle("active", i === index);
  });
  const currentText = AppState.drafts[index] || "";
  renderMarkdownResult(currentText);
  if (document.getElementById("diffContainer")?.style.display !== "none") {
    refreshDiffView();
  }
}

function parseMultipleDrafts(cleanText) {
  if (!cleanText) return null;

  // 1. 显式标记 ===版本1:xxx===
  const explicitDrafts = [];
  const v1 = cleanText.match(/===版本1[^=]*===([\s\S]*?)(?:===版本2|$)/);
  const v2 = cleanText.match(/===版本2[^=]*===([\s\S]*?)(?:===版本3|$)/);
  const v3 = cleanText.match(/===版本3[^=]*===([\s\S]*?)$/);
  if (v1 && v2) {
    explicitDrafts.push({ title: "版本 1", content: v1[1].trim() });
    explicitDrafts.push({ title: "版本 2", content: v2[1].trim() });
    if (v3) explicitDrafts.push({ title: "版本 3", content: v3[1].trim() });
    return explicitDrafts;
  }

  // 2. 自适应探测常见的方案/版本标记 (如 方案1/方案一/【方案一】/### 方案1/Option 1等)
  const pattern = /(?:^|\n)(?:[#*>\s-]*)(?:【)?(?:方案|版本|Option|Draft)\s*([1-9一二三四五ABCabc])[:：\s\-—]*(.*?)(?:】|\*\*)?(?:\n|$)/gi;
  const matches = [];
  let m;
  while ((m = pattern.exec(cleanText)) !== null) {
    matches.push({
      index: m.index,
      fullLength: m[0].length,
      numStr: m[1],
      subTitle: (m[2] || "").trim()
    });
  }

  if (matches.length >= 2) {
    const results = [];
    for (let i = 0; i < matches.length; i++) {
      const cur = matches[i];
      const startPos = cur.index + cur.fullLength;
      const endPos = (i + 1 < matches.length) ? matches[i + 1].index : cleanText.length;
      let body = cleanText.substring(startPos, endPos).trim();

      // 清理末尾可能包含的总结寒暄废话
      if (i === matches.length - 1) {
        body = body.replace(/(?:\n|^)(?:以上|希望|如需|如有|若有|祝您)[\s\S]*$/, "").trim();
      }

      let tabTitle = `方案 ${cur.numStr}`;
      if (cur.subTitle && cur.subTitle.length <= 8) {
        tabTitle = cur.subTitle.replace(/^[（(【]+|[）)】]+$/g, "");
      }

      if (body.length > 0) {
        results.push({
          title: tabTitle,
          content: body
        });
      }
    }

    if (results.length >= 2) {
      return results;
    }
  }

  return null;
}

function renderDraftTabs(draftList, fullText) {
  const container = document.getElementById("draftTabsContainer");
  const listEl = document.getElementById("draftTabsList");
  if (!container || !listEl) return;

  if (!draftList || draftList.length < 2) {
    container.style.display = "none";
    AppState.drafts = [];
    AppState.activeDraftIndex = 0;
    return;
  }

  AppState.drafts = draftList.map(d => d.content);
  // 提供一个看全量的末尾 Tab
  AppState.drafts.push(fullText);

  listEl.innerHTML = "";
  draftList.forEach((d, idx) => {
    const btn = document.createElement("button");
    btn.className = `draft-tab ${idx === 0 ? "active" : ""}`;
    btn.setAttribute("data-index", idx);
    btn.textContent = d.title;
    btn.onclick = () => switchDraft(idx);
    listEl.appendChild(btn);
  });

  const allBtn = document.createElement("button");
  const allIdx = draftList.length;
  allBtn.className = "draft-tab";
  allBtn.setAttribute("data-index", allIdx);
  allBtn.textContent = "完整全部";
  allBtn.onclick = () => switchDraft(allIdx);
  listEl.appendChild(allBtn);

  container.style.display = "block";
  switchDraft(0);
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

  AppState.lastActionType = actionKey;

  // 特权动作：智能往下续写（无需强制选中文字，自动读取光标上文）
  if (actionKey === "continue_write") {
    const precedingText = await getCursorPrecedingText();
    if (!precedingText) {
      showToast("未读取到光标前的上文内容，请先输入或选中一段文字作为引子！");
      return;
    }
    const prompt = `${action.prompt}\n"""\n${precedingText}\n"""\n\n【严格输出规范】：直接输出紧接着上文往下续写的新段落正文本身，不要重复上文，严禁任何寒暄或解释说明。`;
    runAIGeneration(prompt, "正在顺着上文智能续写...", false);
    return;
  }

  if (!AppState.currentSelection) {
    showToast("请先在 Word 文档中选中需要处理的文字！");
    return;
  }

  // 构建上下文语境参考块（若开启了“感知上下文”且存在更长的前后文）
  const isContextAware = document.getElementById("contextAwareCheck")?.checked;
  let contextBlock = "";
  if (isContextAware && AppState.surroundingContext && AppState.surroundingContext.length > AppState.currentSelection.length + 10) {
    contextBlock = `\n\n【前后文语境参考（仅供理解专有名词与行文背景，请勿处理或输出此背景段落）】：\n"""\n${AppState.surroundingContext}\n"""`;
  }

  const isMultiDraft = document.getElementById("multiDraftCheck").checked;
  if (isMultiDraft) {
    const prompt = `请针对以下选中文档内容执行【${action.name}】。${contextBlock}\n\n为了方便用户对比挑选，请同时生成 3 种不同风格的独立候选版本，并使用严格的分隔符分隔：\n\n===版本1:标准平衡===\n(版本1内容)\n===版本2:精简干练===\n(版本2内容)\n===版本3:文雅生动===\n(版本3内容)\n\n待处理文本：\n${AppState.currentSelection}`;
    runAIGeneration(prompt, `正在生成【${action.name}】3 个候选版本...`, true);
  } else {
    const prompt = `${action.prompt}\n\n待处理文本：\n${AppState.currentSelection}${contextBlock}\n\n【严格输出规范】：\n1. 仅输出针对“待处理文本”修改后的单一一版最终正文结果，严禁输出方案一/方案二等多个候选方案；\n2. 严禁包含任何前缀引言、解释说明、提示语（如“修改如下：”、“好的”、“润色结果：”等）或寒暄客套；\n3. 直接输出可直接写回 Word 的纯净最终正文内容。`;
    runAIGeneration(prompt, `正在执行【${action.name}】...`, false);
  }
}

/**
 * 结果二次打磨与多轮追问迭代
 */
function handleRefinement(refineKeyOrCustomText, isCustomInput = false) {
  if (!isCustomInput && refineKeyOrCustomText === "regenerate") {
    if (AppState.lastUserPrompt) {
      runAIGeneration(AppState.lastUserPrompt, "正在重新生成...", document.getElementById("multiDraftCheck")?.checked);
    } else {
      showToast("暂无可重新生成的任务");
    }
    return;
  }

  const currentResult = getCurrentActiveOutputText();
  if (!currentResult) {
    showToast("当前没有可微调的生成结果！");
    return;
  }

  const presetInstructions = {
    shorter: "在保持核心意思完整的前提下，将当前版本进一步精简压缩约 30%，删去冗余修饰，使行文更干练紧凑",
    longer: "在当前版本基础上进一步丰富细节、论据与过渡衔接，使内容更充实饱满",
    more_formal: "将当前版本的措辞调整得更加正式、庄重、严谨，符合高规格书面公文或专业报告规范",
    more_casual: "将当前版本调整得更加通俗自然、平易近人、流畅易读，去除生硬的套话与官腔"
  };

  const instruction = isCustomInput ? refineKeyOrCustomText : (presetInstructions[refineKeyOrCustomText] || refineKeyOrCustomText);
  const origRef = AppState.currentSelection ? `原始文档选区参考：\n"""\n${AppState.currentSelection}\n"""\n\n` : "";

  const refinePrompt = `${origRef}当前已生成的文本版本：\n"""\n${currentResult}\n"""\n\n用户的进一步微调修改要求：【${instruction}】\n\n请基于“当前已生成的文本版本”按照上述微调要求进行针对性改写，直接输出修改后的完整最终正文，严禁输出任何解释说明、前言或寒暄。`;

  runAIGeneration(refinePrompt, "正在根据要求二次微调...", false);
}

async function handleDocAction(docActionKey) {
  const action = DOC_ACTION_PROMPTS[docActionKey];
  if (!action) return;

  AppState.lastActionType = docActionKey;

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

  AppState.lastActionType = "custom";
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

  AppState.lastUserPrompt = userPrompt;

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

    let lastDocStreamSyncTime = 0;
    const directDocStreamCheck = document.getElementById("directStreamToDocCheck");
    const isDirectDocStream = directDocStreamCheck && directDocStreamCheck.checked;

    // 若开启直接打字注入，提前抓取并锁定初始选区 Range
    let targetDocBinding = null;
    if (isDirectDocStream && AppState.isOfficeInitialized) {
      try {
        await Word.run(async (context) => {
          const sel = context.document.getSelection();
          // 用内容控件 ContentControl 或临时标记绑定选区，确保后续连续替换不随光标漂移
          const cc = sel.insertContentControl();
          cc.tag = "ai_stream_target";
          cc.title = "AI 流式生成目标区";
          await context.sync();
          targetDocBinding = cc;
        });
      } catch (bindErr) {
        console.warn("选区 ContentControl 绑定失败，回退到 Selection 替换:", bindErr);
      }
    }

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

            // 流式打字机直接注入 Word (带节流 120ms，避免过度高频调用 Word.run 阻塞)
            if (isDirectDocStream) {
              const now = Date.now();
              if (now - lastDocStreamSyncTime > 120) {
                lastDocStreamSyncTime = now;
                streamChunkDirectToWord(accumulatedRaw);
              }
            }
          }
        } catch (e) {
          // ignore chunk parse
        }
      }
    }

    // 若开启了直接打字注入，生成结束后做一次最终完整同步
    if (isDirectDocStream) {
      streamChunkDirectToWord(accumulatedRaw, true);
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
 * 流式打字机直接注入 Word 正文（提取 think 后的正文流式更新到 Word 选区）
 */
let isStreamingToWordBusy = false;
async function streamChunkDirectToWord(accumulatedRaw, isFinal = false) {
  if (!AppState.isOfficeInitialized || isStreamingToWordBusy) return;

  let textToInsert = accumulatedRaw;
  if (accumulatedRaw.includes("</think>")) {
    textToInsert = accumulatedRaw.split("</think>")[1].trim();
  } else if (accumulatedRaw.includes("<think>")) {
    // 仍在思考中，正文未开始输出
    return;
  }

  if (!textToInsert) return;

  isStreamingToWordBusy = true;
  try {
    await Word.run(async (context) => {
      // 1. 若创建了绑定内容控件，则直接替换其内部文本
      const ccs = context.document.contentControls.getByTag("ai_stream_target");
      ccs.load("items");
      await context.sync();

      if (ccs.items && ccs.items.length > 0) {
        const targetCc = ccs.items[0];
        targetCc.insertText(textToInsert, Word.InsertLocation.replace);
        if (isFinal) {
          // 最终成稿时，解除内容控件包裹，保留纯正文
          targetCc.delete(true);
        }
        await context.sync();
      } else {
        // 2. 回退模式：直接写回当前选区
        const selection = context.document.getSelection();
        selection.insertText(textToInsert, Word.InsertLocation.replace);
        await context.sync();
      }
    });
  } catch (err) {
    console.warn("流式注入 Word 失败/跳过:", err);
  } finally {
    isStreamingToWordBusy = false;
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

  // 自适应多方案解析（不论是否勾选多版本，一旦检测到多个方案便自动拆分选项卡）
  const detectedDrafts = parseMultipleDrafts(cleanText);

  if (detectedDrafts && detectedDrafts.length >= 2) {
    renderDraftTabs(detectedDrafts, cleanText);
  } else {
    document.getElementById("draftTabsContainer").style.display = "none";
    AppState.drafts = [];
    AppState.activeDraftIndex = 0;
    AppState.lastGeneratedText = cleanText;
    renderMarkdownResult(cleanText, true);
  }

  // 若 Diff 面板已展开，自动同步刷新最新对比
  if (document.getElementById("diffContainer")?.style.display !== "none") {
    refreshDiffView();
  }

  // 记录到历史记录
  addHistoryRecord({
    prompt: userPrompt.substring(0, 80),
    result: getCurrentActiveOutputText(),
    mode: AppState.activeMode
  });
}

function renderMarkdownResult(text, isFinished = false) {
  const resultContent = document.getElementById("aiResultContent");
  if (window.marked) {
    resultContent.innerHTML = marked.parse(text);
  } else {
    resultContent.textContent = text;
  }

  // 生成完成时，若包含「原句」->「修改版」校对结构，自动注入“定位原文”与“一键替换”交互按钮
  if (isFinished || !AppState.isGenerating) {
    enhanceAuditReportWithActions(resultContent);
  }
}

/**
 * 扫描审查报告中的「原句」与「修改建议」，自动挂载【📍定位原文】与【⚡一键替换】按钮
 */
function enhanceAuditReportWithActions(containerEl) {
  if (!containerEl) return;
  const items = containerEl.querySelectorAll("li, p");
  items.forEach(el => {
    if (el.querySelector(".audit-inline-actions")) return;
    const text = el.textContent || "";
    const quotes = [];
    const regex = /「([^」]{2,150})」/g;
    let m;
    while ((m = regex.exec(text)) !== null) {
      quotes.push(m[1]);
    }
    if (quotes.length >= 1) {
      const origSnippet = quotes[0];
      const replacementSnippet = quotes.length >= 2 ? quotes[1] : null;

      const actionSpan = document.createElement("span");
      actionSpan.className = "audit-inline-actions";

      const locateBtn = document.createElement("button");
      locateBtn.type = "button";
      locateBtn.className = "btn-audit-chip";
      locateBtn.textContent = "📍 定位原文";
      locateBtn.title = `在 Word 文档中查找并高亮：「${origSnippet}」`;
      locateBtn.onclick = (e) => {
        e.stopPropagation();
        window.locateTextInWord(origSnippet);
      };
      actionSpan.appendChild(locateBtn);

      if (replacementSnippet && replacementSnippet !== origSnippet) {
        const fixBtn = document.createElement("button");
        fixBtn.type = "button";
        fixBtn.className = "btn-audit-chip fix";
        fixBtn.textContent = "⚡ 一键替换";
        fixBtn.title = `将原文「${origSnippet}」直接替换为「${replacementSnippet}」`;
        fixBtn.onclick = (e) => {
          e.stopPropagation();
          window.replaceSpecificTextInWord(origSnippet, replacementSnippet, fixBtn);
        };
        actionSpan.appendChild(fixBtn);
      }

      el.appendChild(actionSpan);
    }
  });
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
  } catch (e) { }
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

window.applyHistoryItem = function (id) {
  const item = AppState.history.find(h => h.id === id);
  if (!item) return;
  AppState.lastGeneratedText = item.result;
  document.getElementById("resultSection").style.display = "block";
  renderMarkdownResult(item.result);
  closeHistoryDrawer();
  showToast("已载入该条历史记录，可直接点击写回");
};

window.copyHistoryItem = function (id) {
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
      const parsed = JSON.parse(saved);
      // 兼容历史老数据，若缺少 showInChips 默认前 5 个为 true
      AppState.promptPresets = parsed.map((p, idx) => ({
        ...p,
        showInChips: typeof p.showInChips === "boolean" ? p.showInChips : (idx < 5)
      }));
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
  } catch (e) { }
  renderPromptChips();
}

function renderPromptChips() {
  const container = document.getElementById("promptChipsList");
  if (!container) return;
  // 筛选出勾选了 showInChips 的提示词胶囊
  const visiblePresets = AppState.promptPresets.filter(p => p.showInChips !== false);
  container.innerHTML = visiblePresets.map(p => `
    <button class="chip" data-text="${escapeHtml(p.content)}" title="${escapeHtml(p.content)}">${escapeHtml(p.title)}</button>
  `).join("");

  // 绑定点击填入指令输入框
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
  AppState.editingPresetId = null;
  document.getElementById("promptsDrawer").style.display = "none";
}

function renderPromptsDrawerList() {
  const container = document.getElementById("promptsPresetList");
  if (!container) return;

  if (!AppState.promptPresets || AppState.promptPresets.length === 0) {
    container.innerHTML = `<div class="empty-state" style="padding:20px; font-size:12px; color:var(--text-3); text-align:center;">暂无收藏提示词</div>`;
    return;
  }

  container.innerHTML = AppState.promptPresets.map(p => {
    // 处于编辑模式的条目
    if (AppState.editingPresetId === p.id) {
      return `
        <div class="preset-item editing" onclick="event.stopPropagation();">
          <div class="preset-edit-box">
            <div class="preset-edit-header">
              <span class="preset-edit-title-label">编辑提示词</span>
              <label class="preset-check-label" title="勾选后在自定义指令下方作为快捷胶囊显示">
                <input type="checkbox" id="editPromptChips_${p.id}" class="preset-chip-toggle" ${p.showInChips !== false ? "checked" : ""}>
                <span class="preset-check-text">在快捷栏显示</span>
              </label>
            </div>
            <input type="text" id="editPromptTitle_${p.id}" class="form-input" value="${escapeHtml(p.title)}" placeholder="标题，例如：公文改写">
            <textarea id="editPromptContent_${p.id}" rows="4" class="form-input" placeholder="具体提示词要求…">${escapeHtml(p.content)}</textarea>
            <div class="preset-edit-actions">
              <button class="btn-subtle" onclick="cancelEditPreset()">取消</button>
              <button class="btn-primary" onclick="saveEditPreset('${p.id}')">保存修改</button>
            </div>
          </div>
        </div>
      `;
    }

    // 默认展示模式
    return `
      <div class="preset-item" onclick="selectPresetPrompt('${p.id}')">
        <div class="preset-info">
          <div class="preset-title">${escapeHtml(p.title)}</div>
          <div class="preset-desc">${escapeHtml(p.content)}</div>
        </div>
        <div class="preset-controls" onclick="event.stopPropagation();">
          <label class="preset-check-label" title="勾选后在自定义指令下方作为快捷胶囊显示">
            <input type="checkbox" class="preset-chip-toggle" ${p.showInChips !== false ? "checked" : ""} onchange="togglePromptInChips('${p.id}', this.checked)">
            <span class="preset-check-text">快捷显示</span>
          </label>
          <button class="btn-text" onclick="startEditPreset('${p.id}')">编辑</button>
          <button class="btn-text danger" onclick="deletePresetPrompt('${p.id}')">删除</button>
        </div>
      </div>
    `;
  }).join("");
}

window.startEditPreset = function (id) {
  AppState.editingPresetId = id;
  renderPromptsDrawerList();
  setTimeout(() => {
    const input = document.getElementById(`editPromptTitle_${id}`);
    if (input) {
      input.focus();
      input.select();
    }
  }, 50);
};

window.cancelEditPreset = function () {
  AppState.editingPresetId = null;
  renderPromptsDrawerList();
};

window.saveEditPreset = function (id) {
  const target = AppState.promptPresets.find(x => x.id === id);
  if (!target) return;
  const titleEl = document.getElementById(`editPromptTitle_${id}`);
  const contentEl = document.getElementById(`editPromptContent_${id}`);
  const chipEl = document.getElementById(`editPromptChips_${id}`);

  const newTitle = titleEl ? titleEl.value.trim() : "";
  const newContent = contentEl ? contentEl.value.trim() : "";

  if (!newTitle || !newContent) {
    showToast("标题和提示词内容均不能为空！");
    return;
  }

  target.title = newTitle;
  target.content = newContent;
  if (chipEl) {
    target.showInChips = chipEl.checked;
  }

  AppState.editingPresetId = null;
  savePromptsToStorage();
  renderPromptsDrawerList();
  renderPromptChips();
  showToast(`已成功修改「${newTitle}」`);
};

window.togglePromptInChips = function (id, isChecked) {
  const target = AppState.promptPresets.find(x => x.id === id);
  if (target) {
    target.showInChips = isChecked;
    savePromptsToStorage();
    renderPromptChips();
    showToast(isChecked ? `已设为在快捷指令栏显示` : `已取消在快捷指令栏显示`);
  }
};

window.selectPresetPrompt = function (id) {
  const p = AppState.promptPresets.find(x => x.id === id);
  if (!p) return;
  document.getElementById("customPromptInput").value = p.content;
  closePromptsDrawer();
  document.getElementById("customPromptInput").focus();
};

window.deletePresetPrompt = function (id) {
  const p = AppState.promptPresets.find(x => x.id === id);
  if (!p) return;
  if (!confirm(`确定要删除提示词「${p.title}」吗？`)) return;
  if (AppState.editingPresetId === id) {
    AppState.editingPresetId = null;
  }
  AppState.promptPresets = AppState.promptPresets.filter(x => x.id !== id);
  savePromptsToStorage();
  renderPromptsDrawerList();
  renderPromptChips();
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
    content,
    showInChips: true
  });
  savePromptsToStorage();
  renderPromptsDrawerList();
  renderPromptChips();
  titleInput.value = "";
  contentInput.value = "";
  showToast("已成功添加到 Prompt 收藏库！");
}

function exportPromptsToFile() {
  try {
    const jsonStr = JSON.stringify(AppState.promptPresets, null, 2);
    const blob = new Blob([jsonStr], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `word-ai-prompts-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast("预设已成功导出为 JSON 文件！");
  } catch (e) {
    showToast("导出失败: " + e.message);
  }
}

function importPromptsFromFile(e) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const imported = JSON.parse(event.target.result);
      if (Array.isArray(imported)) {
        // 合并去重
        let count = 0;
        imported.forEach(item => {
          if (item && item.title && item.content) {
            const exists = AppState.promptPresets.some(p => p.title === item.title);
            if (!exists) {
              AppState.promptPresets.push({
                id: item.id || Date.now().toString() + Math.random().toString(36).substring(2, 5),
                title: item.title,
                content: item.content,
                showInChips: typeof item.showInChips === "boolean" ? item.showInChips : true
              });
              count++;
            }
          }
        });
        savePromptsToStorage();
        renderPromptsDrawerList();
        renderPromptChips();
        showToast(`成功导入 ${count} 条新预设！`);
      } else {
        showToast("导入失败：JSON 格式不符");
      }
    } catch (err) {
      showToast("解析 JSON 失败: " + err.message);
    }
    e.target.value = "";
  };
  reader.readAsText(file);
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
  } catch (e) { }
  updateHeaderModelBadge();
}

function saveConfigToStorage() {
  try {
    localStorage.setItem("word_ai_copilot_config", JSON.stringify(AppState.config));
    updateHeaderModelBadge();
  } catch (e) { }
}

function updateHeaderModelBadge() {
  const badge = document.getElementById("currentModelBadge");
  const badgeText = document.getElementById("currentModelBadgeText");
  const label = AppState.config.model || "未配置模型";
  if (badgeText) {
    badgeText.textContent = label;
  } else if (badge) {
    badge.textContent = label;
  }
  if (badge) {
    badge.title = `当前模型：${label}\n接口：${AppState.config.baseUrl}\n（点击可快速切换模型）`;
  }
}

/**
 * 顶部模型徽章下拉秒切菜单
 */
async function toggleQuickModelDropdown() {
  const dropdown = document.getElementById("quickModelDropdown");
  if (!dropdown) return;

  if (dropdown.style.display !== "none") {
    dropdown.style.display = "none";
    return;
  }

  dropdown.style.display = "block";
  dropdown.innerHTML = `<div style="padding:8px 10px; font-size:11px; color:var(--text-3);">正在加载模型列表…</div>`;

  // 若尚未缓存模型列表，从当前配置的接口拉取
  if (!AppState.cachedModels || AppState.cachedModels.length === 0) {
    try {
      const baseUrl = (AppState.config.baseUrl || "").trim();
      const apiKey = (AppState.config.apiKey || "").trim();
      if (baseUrl) {
        const rawEndpoint = `${baseUrl.replace(/\/+$/, "")}/models`;
        const { fetchUrl, fetchHeaders } = resolveFetchEndpointAndHeaders(rawEndpoint, {
          "Authorization": `Bearer ${apiKey || "dummy-key"}`
        });
        const res = await fetch(fetchUrl, { method: "GET", headers: fetchHeaders });
        if (res.ok) {
          const data = await res.json();
          let list = [];
          if (Array.isArray(data)) list = data.map(x => typeof x === "string" ? x : (x.id || x.name)).filter(Boolean);
          else if (data && Array.isArray(data.data)) list = data.data.map(x => typeof x === "string" ? x : (x.id || x.name)).filter(Boolean);
          else if (data && Array.isArray(data.models)) list = data.models.map(x => typeof x === "string" ? x : (x.name || x.id)).filter(Boolean);
          if (list.length > 0) AppState.cachedModels = list;
        }
      }
    } catch (e) {
      console.warn("顶部快捷拉取模型失败:", e);
    }
  }

  const models = AppState.cachedModels && AppState.cachedModels.length > 0
    ? AppState.cachedModels
    : (AppState.config.model ? [AppState.config.model] : []);

  if (models.length === 0) {
    dropdown.innerHTML = `<button type="button" class="quick-model-item" id="btnQuickOpenSettings">⚙️ 前往设置配置接口与模型…</button>`;
    document.getElementById("btnQuickOpenSettings").onclick = () => {
      dropdown.style.display = "none";
      openSettingsDrawer();
    };
    return;
  }

  dropdown.innerHTML = models.map(m => `
    <button type="button" class="quick-model-item ${m === AppState.config.model ? "active" : ""}" data-model="${escapeHtml(m)}">
      <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${escapeHtml(m)}</span>
      ${m === AppState.config.model ? "<span>✓</span>" : ""}
    </button>
  `).join("") + `
    <div style="border-top:1px solid var(--stroke); margin:4px 0;"></div>
    <button type="button" class="quick-model-item" id="btnQuickRefreshModels" style="color:var(--brand); font-size:11px;">🔄 重新拉取模型列表</button>
  `;

  dropdown.querySelectorAll(".quick-model-item[data-model]").forEach(btn => {
    btn.onclick = () => {
      const chosen = btn.getAttribute("data-model");
      AppState.config.model = chosen;
      saveConfigToStorage();
      dropdown.style.display = "none";
      showToast(`⚡ 已切换模型为：${chosen}`);
    };
  });

  const refreshBtn = document.getElementById("btnQuickRefreshModels");
  if (refreshBtn) {
    refreshBtn.onclick = async (e) => {
      e.stopPropagation();
      AppState.cachedModels = [];
      dropdown.style.display = "none";
      await toggleQuickModelDropdown();
    };
  }
}

function openSettingsDrawer() {
  const drawer = document.getElementById("settingsDrawer");
  document.getElementById("providerPreset").value = AppState.config.provider || "local_custom";
  document.getElementById("apiUrlInput").value = AppState.config.baseUrl || "";
  document.getElementById("apiKeyInput").value = AppState.config.apiKey || "";

  const currentModel = AppState.config.model || "";
  document.getElementById("modelNameInput").value = currentModel;

  const modelSelect = document.getElementById("modelSelect");
  modelSelect.innerHTML = currentModel
    ? `<option value="${currentModel}">${currentModel}</option>`
    : `<option value="">(点击刷新模型列表)</option>`;
  modelSelect.value = currentModel;

  updateCustomModelSelectUI(currentModel || "选择模型...");

  document.getElementById("temperatureSlider").value = AppState.config.temperature || 0.7;
  document.getElementById("temperatureVal").textContent = AppState.config.temperature || 0.7;
  document.getElementById("systemPromptInput").value = AppState.config.systemPrompt || "";
  document.getElementById("testConnectionStatus").textContent = "";
  document.getElementById("modelFetchStatus").textContent = "";

  drawer.style.display = "flex";

  // 自动尝试静默拉取可用模型列表
  if (AppState.config.baseUrl) {
    fetchAvailableModels(true);
  }
}

function updateCustomModelSelectUI(label) {
  const labelEl = document.getElementById("customModelSelectLabel");
  if (labelEl) {
    labelEl.textContent = label || "选择模型...";
  }
}

function toggleCustomModelMenu() {
  const menu = document.getElementById("customModelMenu");
  const btn = document.getElementById("customModelSelectBtn");
  if (!menu) return;

  const isHidden = menu.style.display === "none";
  if (isHidden) {
    renderCustomModelMenuItems();
    menu.style.display = "block";
    if (btn) btn.classList.add("active");
  } else {
    menu.style.display = "none";
    if (btn) btn.classList.remove("active");
  }
}

function renderCustomModelMenuItems() {
  const menu = document.getElementById("customModelMenu");
  if (!menu) return;

  const models = AppState.cachedModels && AppState.cachedModels.length > 0
    ? AppState.cachedModels
    : (AppState.config.model ? [AppState.config.model] : []);

  if (models.length === 0) {
    menu.innerHTML = `<div style="padding:8px 10px; font-size:11.5px; color:var(--text-3); text-align:center;">暂无模型，请点击右上角【获取/刷新】</div>`;
    return;
  }

  const currentVal = document.getElementById("modelNameInput").value.trim() || AppState.config.model;

  menu.innerHTML = models.map(m => `
    <button type="button" class="custom-select-item ${m === currentVal ? "active" : ""}" data-val="${escapeHtml(m)}">
      <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${escapeHtml(m)}</span>
      ${m === currentVal ? "<span style='font-weight:bold;'>✓</span>" : ""}
    </button>
  `).join("");

  menu.querySelectorAll(".custom-select-item").forEach(item => {
    item.onclick = (e) => {
      e.stopPropagation();
      const val = item.getAttribute("data-val");
      selectCustomModel(val);
    };
  });
}

function selectCustomModel(val) {
  const inputEl = document.getElementById("modelNameInput");
  const selectEl = document.getElementById("modelSelect");
  const menu = document.getElementById("customModelMenu");
  const btn = document.getElementById("customModelSelectBtn");

  if (inputEl) inputEl.value = val;
  if (selectEl) {
    let exists = false;
    for (let opt of selectEl.options) {
      if (opt.value === val) {
        selectEl.value = val;
        exists = true;
        break;
      }
    }
    if (!exists) {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = val;
      selectEl.appendChild(opt);
      selectEl.value = val;
    }
  }

  updateCustomModelSelectUI(val);
  if (menu) menu.style.display = "none";
  if (btn) btn.classList.remove("active");
}

function closeSettingsDrawer() {
  document.getElementById("settingsDrawer").style.display = "none";
}

function toggleManualModelInput() {
  const inputEl = document.getElementById("modelNameInput");
  const customWrapper = document.getElementById("customModelSelectBtn");
  const toggleBtn = document.getElementById("btnToggleManualModel");
  const isHidden = inputEl.style.display === "none";
  if (isHidden) {
    inputEl.style.display = "block";
    if (customWrapper) customWrapper.style.display = "none";
    toggleBtn.textContent = "切换为下拉选择";
    inputEl.focus();
  } else {
    inputEl.style.display = "none";
    if (customWrapper) customWrapper.style.display = "flex";
    toggleBtn.textContent = "切换为手动输入";
  }
}

async function fetchAvailableModels(silent = false) {
  const baseUrl = document.getElementById("apiUrlInput").value.trim();
  const apiKey = document.getElementById("apiKeyInput").value.trim();
  const selectEl = document.getElementById("modelSelect");
  const statusEl = document.getElementById("modelFetchStatus");

  if (!baseUrl) {
    if (!silent) showToast("请先填写 API Base URL！");
    return;
  }

  if (statusEl) statusEl.textContent = "正在拉取模型...";

  const rawEndpoint = `${baseUrl.replace(/\/+$/, "")}/models`;
  const defaultHeaders = {
    "Authorization": `Bearer ${apiKey || "dummy-key"}`
  };
  const { fetchUrl, fetchHeaders } = resolveFetchEndpointAndHeaders(rawEndpoint, defaultHeaders);

  try {
    const res = await fetch(fetchUrl, {
      method: "GET",
      headers: fetchHeaders
    });

    if (!res.ok) {
      throw new Error(`HTTP ${res.status}`);
    }

    const data = await res.json();
    let modelList = [];

    if (Array.isArray(data)) {
      modelList = data.map(item => (typeof item === "string" ? item : (item.id || item.name))).filter(Boolean);
    } else if (data && Array.isArray(data.data)) {
      modelList = data.data.map(item => (typeof item === "string" ? item : (item.id || item.name))).filter(Boolean);
    } else if (data && Array.isArray(data.models)) {
      modelList = data.models.map(item => (typeof item === "string" ? item : (item.name || item.id))).filter(Boolean);
    }

    if (modelList.length === 0) {
      throw new Error("接口返回列表中无模型");
    }

    AppState.cachedModels = modelList;

    // 填充下拉选项
    selectEl.innerHTML = "";
    modelList.forEach(m => {
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = m;
      selectEl.appendChild(opt);
    });

    const targetModel = AppState.config.model || document.getElementById("modelNameInput").value.trim();
    if (targetModel && modelList.includes(targetModel)) {
      selectEl.value = targetModel;
    } else {
      selectEl.value = modelList[0];
    }
    document.getElementById("modelNameInput").value = selectEl.value;
    updateCustomModelSelectUI(selectEl.value);

    if (statusEl) {
      statusEl.textContent = `已获取 ${modelList.length} 个模型`;
      setTimeout(() => { if (statusEl) statusEl.textContent = ""; }, 3000);
    }
    if (!silent) showToast(`成功获取到 ${modelList.length} 个可用模型！`);
  } catch (err) {
    console.warn("拉取模型列表失败:", err);
    if (statusEl) statusEl.textContent = "未获取到，可手动输入";
    if (!silent) showToast("获取模型列表失败: " + err.message);
  }
}

function getSelectedModelName() {
  const inputEl = document.getElementById("modelNameInput");
  const selectEl = document.getElementById("modelSelect");
  if (inputEl.style.display !== "none" && inputEl.value.trim()) {
    return inputEl.value.trim();
  }
  if (selectEl.value) {
    return selectEl.value;
  }
  return inputEl.value.trim();
}

function applyProviderPreset(providerKey) {
  const preset = PROVIDER_PRESETS[providerKey];
  if (!preset) return;
  if (preset.baseUrl) {
    document.getElementById("apiUrlInput").value = preset.baseUrl;
    fetchAvailableModels(true);
  }
  if (preset.model) {
    document.getElementById("modelNameInput").value = preset.model;
    const selectEl = document.getElementById("modelSelect");
    let found = false;
    for (let opt of selectEl.options) {
      if (opt.value === preset.model) {
        selectEl.value = preset.model;
        found = true;
        break;
      }
    }
    if (!found) {
      const opt = document.createElement("option");
      opt.value = preset.model;
      opt.textContent = preset.model;
      selectEl.prepend(opt);
      selectEl.value = preset.model;
    }
    updateCustomModelSelectUI(preset.model);
  }
}

function saveSettingsFromDrawer() {
  AppState.config.provider = document.getElementById("providerPreset").value;
  AppState.config.baseUrl = document.getElementById("apiUrlInput").value.trim();
  AppState.config.apiKey = document.getElementById("apiKeyInput").value.trim();
  AppState.config.model = getSelectedModelName();
  AppState.config.temperature = parseFloat(document.getElementById("temperatureSlider").value);
  AppState.config.systemPrompt = document.getElementById("systemPromptInput").value.trim();

  saveConfigToStorage();
  updateHeaderModelBadge();
  closeSettingsDrawer();
  showToast("设置已保存并生效！");
}

async function testAPIConnection() {
  const statusEl = document.getElementById("testConnectionStatus");
  statusEl.className = "test-status";
  statusEl.textContent = "正在连接中...";

  const baseUrl = document.getElementById("apiUrlInput").value.trim();
  const apiKey = document.getElementById("apiKeyInput").value.trim();
  const model = getSelectedModelName();

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
