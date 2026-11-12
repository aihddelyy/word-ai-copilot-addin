// ==========================================================================
// Word AI 创作助手 - WPS JS 加载项主进程脚本
// ==========================================================================

let wpsRibbonUI = null;

// 在 WPS 启动加载 Ribbon 时回调
function OnWPSWorkTabLoad(ribbon) {
  wpsRibbonUI = ribbon;
  return true;
}

// 获取侧边栏任务窗格地址（优先读取本地配置，默认连接极速云端）
function getTaskpaneUrl() {
  const cachedUrl = wps.PluginStorage.getItem("WordAICopilotCustomUrl");
  if (cachedUrl && cachedUrl.trim() !== "") {
    return cachedUrl.trim();
  }
  // 默认由安装脚本自动配置为 Cloudflare Pages 或 GitHub Pages，并带上 ?host=wps 标识
  return "https://aihddelyy.github.io/word-ai-copilot-addin/src/taskpane.html?host=wps";
}

// 统一响应功能区按钮点击事件
function OnAction(control) {
  const btnId = control.Id;
  const storageKey = "WordAICopilotTaskpaneId";
  let tskpane = null;
  const existingId = wps.PluginStorage.getItem(storageKey);

  if (existingId) {
    try {
      tskpane = wps.GetTaskPane(existingId);
    } catch (e) {
      tskpane = null;
    }
  }

  if (!tskpane) {
    const url = getTaskpaneUrl();
    try {
      tskpane = wps.CreateTaskPane(url, "Word AI 创作助手");
    } catch (err) {
      alert("创建 AI 创作助手任务窗格失败: " + err.message);
      return false;
    }

    if (tskpane) {
      wps.PluginStorage.setItem(storageKey, tskpane.ID);
      try {
        // 设置默认宽度 380px，保证双列按钮与开关整齐舒展
        tskpane.Width = 380;
        // 停靠在右侧
        if (wps.Enum && wps.Enum.JSKsoEnum_msoCTPDockPositionRight !== undefined) {
          tskpane.DockPosition = wps.Enum.JSKsoEnum_msoCTPDockPositionRight;
        }
      } catch (e) {
        console.warn("设置任务窗格属性异常", e);
      }
      tskpane.Visible = true;
    }
  } else {
    // 若点击的是主打开按钮，则切换显隐状态
    if (btnId === "btnOpenCopilotPane") {
      tskpane.Visible = !tskpane.Visible;
    } else {
      // 快捷动作按钮确保面板可见
      tskpane.Visible = true;
    }
  }

  return true;
}

// 图标回调（备用）
function GetImage(control) {
  return "icon-32.png";
}






