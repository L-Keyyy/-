(() => {
  const INSTANCE_KEY = "__WEB_AI_VISUAL_EDITOR__";
  const existing = window[INSTANCE_KEY];
  if (existing?.toggle) {
    existing.toggle();
    return;
  }

  const sessionId = `wae-${Date.now().toString(36)}-${Math.random()
    .toString(36)
    .slice(2, 7)}`;
  const actionLabels = {
    remove: "删除模块",
    edit_text: "修改文字",
    move_up: "向上移动",
    move_down: "向下移动",
    duplicate: "复制模块",
    insert_before: "上方新增",
    insert_after: "下方新增",
    resize: "调整尺寸",
    annotate: "添加AI说明",
  };

  const state = {
    enabled: true,
    selected: null,
    hovered: null,
    history: [],
    redo: [],
    changes: [],
    sequence: 0,
    toastTimer: null,
    promptResolver: null,
  };

  const root = document.createElement("div");
  root.id = "wae-root";
  root.dataset.waeUi = "true";
  root.innerHTML = `
    <div class="wae-toolbar" data-wae-ui="true" role="toolbar" aria-label="Web AI 页面修改器">
      <div class="wae-brand"><span class="wae-brand-mark">AI</span><span>页面修改器</span></div>
      <div class="wae-divider"></div>
      <button class="wae-tool-button is-active" data-action="toggle-select" title="暂停后可以操作原页面">◉ 选择</button>
      <button class="wae-tool-button is-danger" data-action="remove" data-needs-selection>删除</button>
      <button class="wae-tool-button" data-action="edit-text" data-needs-selection>改文字</button>
      <button class="wae-tool-button" data-action="move-up" data-needs-selection>上移</button>
      <button class="wae-tool-button" data-action="move-down" data-needs-selection>下移</button>
      <button class="wae-tool-button" data-action="duplicate" data-needs-selection>复制</button>
      <button class="wae-tool-button" data-action="insert-before" data-needs-selection>新增↑</button>
      <button class="wae-tool-button" data-action="insert-after" data-needs-selection>新增↓</button>
      <button class="wae-tool-button" data-action="resize" data-needs-selection>尺寸</button>
      <button class="wae-tool-button" data-action="annotate" data-needs-selection>备注</button>
      <div class="wae-divider"></div>
      <button class="wae-tool-button" data-action="undo" title="撤回（⌘/Ctrl+Z）">↶</button>
      <button class="wae-tool-button" data-action="redo" title="重做（⌘/Ctrl+Shift+Z）">↷</button>
      <button class="wae-tool-button is-primary" data-action="copy">复制AI指令</button>
      <button class="wae-tool-button is-primary" data-action="export">导出修改单</button>
    </div>
    <aside class="wae-panel" data-wae-ui="true" aria-label="修改记录">
      <div class="wae-panel-head">
        <div class="wae-panel-head-row">
          <div class="wae-panel-title">AI修改工作台</div>
          <button class="wae-close-button" data-action="close" aria-label="关闭修改模式">×</button>
        </div>
        <div class="wae-selection-card">
          <div class="wae-selection-label">当前选中</div>
          <div class="wae-selection-name" data-role="selection-name">移动鼠标并点击页面模块</div>
          <div class="wae-selection-meta" data-role="selection-meta"><span>按住 Alt / Option 可精确选择元素</span></div>
        </div>
      </div>
      <div class="wae-panel-actions">
        <button data-action="undo">撤回</button>
        <button data-action="redo">重做</button>
        <button data-action="clear">清空记录</button>
      </div>
      <div class="wae-change-list" data-role="change-list"></div>
      <div class="wae-status-bar">
        <span>修改记录 <span class="wae-status-count" data-role="change-count">0</span> 条</span>
        <span data-role="mode-status">选择模式</span>
      </div>
    </aside>
    <div class="wae-modal" data-role="modal" data-wae-ui="true" aria-hidden="true">
      <div class="wae-modal-dialog" role="dialog" aria-modal="true" aria-labelledby="wae-modal-title">
        <div class="wae-modal-title" id="wae-modal-title" data-role="modal-title">输入修改内容</div>
        <label class="wae-modal-label" data-role="modal-label" for="wae-modal-input">修改要求</label>
        <textarea id="wae-modal-input" data-role="modal-input" rows="4"></textarea>
        <div class="wae-modal-actions">
          <button data-action="modal-cancel">取消</button>
          <button class="is-primary" data-action="modal-confirm">确认</button>
        </div>
      </div>
    </div>
    <div class="wae-toast" data-role="toast"></div>
  `;
  document.documentElement.appendChild(root);

  const $ = (selector) => root.querySelector(selector);
  const $$ = (selector) => [...root.querySelectorAll(selector)];
  const selectionName = $('[data-role="selection-name"]');
  const selectionMeta = $('[data-role="selection-meta"]');
  const changeList = $('[data-role="change-list"]');
  const changeCount = $('[data-role="change-count"]');
  const modeStatus = $('[data-role="mode-status"]');
  const toast = $('[data-role="toast"]');
  const modal = $('[data-role="modal"]');
  const modalTitle = $('[data-role="modal-title"]');
  const modalLabel = $('[data-role="modal-label"]');
  const modalInput = $('[data-role="modal-input"]');

  function normalizeText(value, limit = 180) {
    return String(value ?? "")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, limit);
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function cssEscape(value) {
    if (window.CSS?.escape) return window.CSS.escape(value);
    return String(value).replace(/[^a-zA-Z0-9_-]/g, "\\$&");
  }

  function isEditorElement(element) {
    return Boolean(element?.closest?.("#wae-root, [data-wae-ui='true']"));
  }

  function resolveSelectionTarget(rawTarget, event) {
    if (event?.altKey) return rawTarget;
    const semanticSelector = [
      "[data-module-id]",
      "[data-section]",
      "[data-component]",
      "[role='region']",
      "[role='dialog']",
      "section",
      "article",
      "aside",
      "nav",
      "header",
      "footer",
      "form",
      "table",
    ].join(",");
    const semantic = rawTarget.closest(semanticSelector);
    if (semantic && !semantic.matches("html,body,main")) return semantic;

    let cursor = rawTarget;
    for (let depth = 0; cursor && depth < 5; depth += 1) {
      const className = [...cursor.classList].join(" ").toLowerCase();
      const looksLikeModule = /(card|panel|module|widget|block|section)/.test(
        className,
      );
      const looksLikeLabel = /(title|heading|header|label|icon|button)/.test(
        className,
      );
      if (looksLikeModule && !looksLikeLabel) return cursor;
      cursor = cursor.parentElement;
    }
    return rawTarget;
  }

  function selectorIsUnique(selector) {
    try {
      return document.querySelectorAll(selector).length === 1;
    } catch {
      return false;
    }
  }

  function buildSelector(element) {
    if (element.id) {
      const selector = `#${cssEscape(element.id)}`;
      if (selectorIsUnique(selector)) return selector;
    }

    const stableAttributes = [
      "data-testid",
      "data-module-id",
      "data-section",
      "data-component",
      "aria-label",
      "name",
    ];
    for (const attribute of stableAttributes) {
      const value = element.getAttribute(attribute);
      if (!value || value.length > 100) continue;
      const selector = `${element.tagName.toLowerCase()}[${attribute}="${cssEscape(value)}"]`;
      if (selectorIsUnique(selector)) return selector;
    }

    const parts = [];
    let cursor = element;
    while (cursor && cursor !== document.body && parts.length < 7) {
      let part = cursor.tagName.toLowerCase();
      const classes = [...cursor.classList]
        .filter(
          (name) =>
            !name.startsWith("wae-") &&
            name.length < 50 &&
            !/^[a-z0-9_-]{18,}$/i.test(name),
        )
        .slice(0, 2);
      if (classes.length) part += `.${classes.map(cssEscape).join(".")}`;
      const sameTagSiblings = cursor.parentElement
        ? [...cursor.parentElement.children].filter(
            (item) => item.tagName === cursor.tagName,
          )
        : [];
      if (sameTagSiblings.length > 1) {
        part += `:nth-of-type(${sameTagSiblings.indexOf(cursor) + 1})`;
      }
      parts.unshift(part);
      const selector = parts.join(" > ");
      if (selectorIsUnique(selector)) return selector;
      cursor = cursor.parentElement;
    }
    return `body > ${parts.join(" > ")}`;
  }

  function getModuleName(element) {
    const directLabel =
      element.getAttribute("aria-label") ||
      element.getAttribute("data-title") ||
      element.getAttribute("title");
    if (directLabel) return normalizeText(directLabel, 90);

    const headingSelector =
      "h1,h2,h3,h4,h5,h6,[role='heading'],.title,.card-title,.section-title";
    const heading = element.matches(headingSelector)
      ? element
      : element.querySelector(headingSelector);
    const headingText = normalizeText(heading?.textContent, 90);
    if (headingText) return headingText;

    const text = normalizeText(element.textContent, 90);
    return text || `${element.tagName.toLowerCase()} 模块`;
  }

  function getDescriptor(element) {
    const rect = element.getBoundingClientRect();
    const parent = element.parentElement;
    const siblings = parent ? [...parent.children] : [];
    const index = siblings.indexOf(element);
    const previous = index > 0 ? siblings[index - 1] : null;
    const next = index >= 0 ? siblings[index + 1] : null;
    const html = element.outerHTML
      .replace(/\sdata-web-ai-id="[^"]*"/g, "")
      .replace(/\sclass="([^"]*)wae-page-(hovered|selected)([^"]*)"/g, ' class="$1$3"');

    return {
      moduleName: getModuleName(element),
      selector: buildSelector(element),
      tag: element.tagName.toLowerCase(),
      id: element.id || "",
      classes: [...element.classList]
        .filter((name) => !name.startsWith("wae-"))
        .slice(0, 8),
      text: normalizeText(element.textContent),
      htmlSnippet: html.slice(0, 1600),
      position: {
        x: Math.round(rect.left + window.scrollX),
        y: Math.round(rect.top + window.scrollY),
        width: Math.round(rect.width),
        height: Math.round(rect.height),
      },
      context: {
        parent: parent ? getModuleName(parent) : "",
        previous: normalizeText(previous?.textContent, 80),
        next: normalizeText(next?.textContent, 80),
      },
    };
  }

  function showToast(message) {
    toast.textContent = message;
    toast.classList.add("is-visible");
    window.clearTimeout(state.toastTimer);
    state.toastTimer = window.setTimeout(
      () => toast.classList.remove("is-visible"),
      1800,
    );
  }

  function askUser({ title, label, defaultValue = "", multiline = true }) {
    if (state.promptResolver) state.promptResolver(null);
    modalTitle.textContent = title;
    modalLabel.textContent = label;
    modalInput.value = defaultValue;
    modalInput.rows = multiline ? 5 : 2;
    modal.classList.add("is-open");
    modal.setAttribute("aria-hidden", "false");
    window.setTimeout(() => {
      modalInput.focus();
      modalInput.select();
    }, 0);
    return new Promise((resolve) => {
      state.promptResolver = resolve;
    });
  }

  function closePrompt(value) {
    modal.classList.remove("is-open");
    modal.setAttribute("aria-hidden", "true");
    const resolve = state.promptResolver;
    state.promptResolver = null;
    resolve?.(value);
  }

  function setHovered(element) {
    if (state.hovered === element) return;
    state.hovered?.classList.remove("wae-page-hovered");
    state.hovered = element;
    if (element && element !== state.selected) {
      element.classList.add("wae-page-hovered");
    }
  }

  function setSelected(element) {
    state.selected?.classList.remove("wae-page-selected");
    if (!element || !element.isConnected) {
      state.selected = null;
      render();
      return;
    }
    state.selected = element;
    element.dataset.webAiId ||= `${sessionId}-${++state.sequence}`;
    element.classList.remove("wae-page-hovered");
    element.classList.add("wae-page-selected");
    render();
  }

  function makeChange(action, target, detail, data = {}) {
    return {
      id: `${sessionId}-change-${++state.sequence}`,
      action,
      actionLabel: actionLabels[action] || action,
      target,
      detail,
      data,
      createdAt: new Date().toISOString(),
    };
  }

  function executeCommand(command) {
    command.redo();
    state.history.push(command);
    state.redo = [];
    state.changes.push(command.change);
    render();
    autosave();
    showToast(`${command.change.actionLabel}已记录`);
  }

  function undo() {
    const command = state.history.pop();
    if (!command) return showToast("当前没有可撤回的操作");
    command.undo();
    state.redo.push(command);
    state.changes = state.changes.filter(
      (item) => item.id !== command.change.id,
    );
    const nextSelection = command.selectionAfterUndo
      ? command.selectionAfterUndo()
      : command.element ?? null;
    setSelected(nextSelection);
    render();
    autosave();
    showToast(`已撤回：${command.change.actionLabel}`);
  }

  function redo() {
    const command = state.redo.pop();
    if (!command) return showToast("当前没有可重做的操作");
    command.redo();
    state.history.push(command);
    state.changes.push(command.change);
    const nextSelection = command.selectionAfterRedo
      ? command.selectionAfterRedo()
      : command.element ?? null;
    setSelected(nextSelection);
    render();
    autosave();
    showToast(`已重做：${command.change.actionLabel}`);
  }

  function selectedOrNotify() {
    if (state.selected?.isConnected) return state.selected;
    showToast("请先点击选择页面模块");
    return null;
  }

  function removeSelected() {
    const element = selectedOrNotify();
    if (!element) return;
    const target = getDescriptor(element);
    const previousDisplay = element.style.display;
    const previousAriaHidden = element.getAttribute("aria-hidden");
    const change = makeChange(
      "remove",
      target,
      "从页面中删除整个模块；同时清理相关样式、交互和无用数据计算。",
    );
    executeCommand({
      element,
      change,
      redo: () => {
        element.classList.remove("wae-page-selected", "wae-page-hovered");
        element.style.display = "none";
        element.setAttribute("aria-hidden", "true");
        state.selected = null;
      },
      undo: () => {
        element.style.display = previousDisplay;
        if (previousAriaHidden === null) element.removeAttribute("aria-hidden");
        else element.setAttribute("aria-hidden", previousAriaHidden);
      },
      selectionAfterUndo: () => element,
      selectionAfterRedo: () => null,
    });
  }

  async function editSelectedText() {
    const selected = selectedOrNotify();
    if (!selected) return;
    const textSelector =
      "h1,h2,h3,h4,h5,h6,[role='heading'],button,label,span,p,strong,small";
    const element =
      selected.childElementCount === 0 || selected.matches(textSelector)
        ? selected
        : selected.querySelector(textSelector);
    if (!element) {
      showToast("该模块没有可直接修改的文字，请使用“备注”说明要求");
      return;
    }
    const before = element.textContent ?? "";
    const after = await askUser({
      title: "修改文字",
      label: "输入修改后的文字",
      defaultValue: before,
      multiline: false,
    });
    if (after === null || after === before) return;
    const target = getDescriptor(element);
    const change = makeChange("edit_text", target, `“${normalizeText(before)}” → “${normalizeText(after)}”`, {
      before,
      after,
    });
    executeCommand({
      element,
      change,
      redo: () => {
        element.textContent = after;
      },
      undo: () => {
        element.textContent = before;
      },
    });
  }

  function moveSelected(direction) {
    const element = selectedOrNotify();
    if (!element) return;
    const parent = element.parentElement;
    if (!parent) return;
    const sibling =
      direction === "up"
        ? element.previousElementSibling
        : element.nextElementSibling;
    if (!sibling) return showToast(direction === "up" ? "已经在最上方" : "已经在最下方");
    const originalNext = element.nextSibling;
    const target = getDescriptor(element);
    const siblingTarget = getDescriptor(sibling);
    const action = direction === "up" ? "move_up" : "move_down";
    const change = makeChange(
      action,
      target,
      `${direction === "up" ? "移动到" : "移动到"}“${siblingTarget.moduleName}”${direction === "up" ? "之前" : "之后"}`,
      { relativeTarget: siblingTarget },
    );
    executeCommand({
      element,
      change,
      redo: () => {
        if (direction === "up") parent.insertBefore(element, sibling);
        else parent.insertBefore(element, sibling.nextSibling);
      },
      undo: () => {
        if (originalNext?.parentNode === parent) parent.insertBefore(element, originalNext);
        else parent.appendChild(element);
      },
    });
  }

  function stripDuplicateIdentifiers(element) {
    element.removeAttribute("id");
    element.removeAttribute("data-web-ai-id");
    element.classList.remove("wae-page-selected", "wae-page-hovered");
    for (const child of element.querySelectorAll("[id], [data-web-ai-id]")) {
      child.removeAttribute("id");
      child.removeAttribute("data-web-ai-id");
    }
  }

  function duplicateSelected() {
    const element = selectedOrNotify();
    if (!element || !element.parentElement) return;
    const parent = element.parentElement;
    const reference = element.nextSibling;
    const clone = element.cloneNode(true);
    stripDuplicateIdentifiers(clone);
    clone.dataset.webAiGenerated = "duplicate";
    const target = getDescriptor(element);
    const change = makeChange(
      "duplicate",
      target,
      `复制“${target.moduleName}”模块，并在原模块下方新增一份。`,
      { htmlSnippet: clone.outerHTML.slice(0, 1600) },
    );
    executeCommand({
      element: clone,
      change,
      redo: () => {
        if (reference?.parentNode === parent) parent.insertBefore(clone, reference);
        else parent.appendChild(clone);
      },
      undo: () => clone.remove(),
      selectionAfterUndo: () => element,
      selectionAfterRedo: () => clone,
    });
  }

  function createPlaceholder(parent, instruction) {
    if (parent.tagName === "TBODY" || parent.tagName === "THEAD") {
      const row = document.createElement("tr");
      const cell = document.createElement("td");
      const columnCount = parent.querySelector("tr")?.children.length || 1;
      cell.colSpan = columnCount;
      cell.textContent = `新增模块：${instruction}`;
      cell.className = "wae-page-placeholder";
      row.appendChild(cell);
      return row;
    }
    const element = document.createElement(
      parent.matches("ul,ol") ? "li" : "div",
    );
    element.className = "wae-page-placeholder";
    element.textContent = `新增模块：${instruction}`;
    return element;
  }

  async function insertPlaceholder(position) {
    const element = selectedOrNotify();
    if (!element || !element.parentElement) return;
    const instruction = await askUser({
      title: position === "before" ? "在上方新增模块" : "在下方新增模块",
      label: "描述需要新增的内容、数据和交互",
      defaultValue: "请在这里增加一个新模块",
      multiline: true,
    });
    if (!instruction?.trim()) return;
    const parent = element.parentElement;
    const placeholder = createPlaceholder(parent, instruction.trim());
    placeholder.dataset.webAiGenerated = "placeholder";
    const target = getDescriptor(element);
    const action = position === "before" ? "insert_before" : "insert_after";
    const change = makeChange(
      action,
      target,
      `${position === "before" ? "在上方" : "在下方"}新增：${instruction.trim()}`,
      { instruction: instruction.trim(), position },
    );
    executeCommand({
      element: placeholder,
      change,
      redo: () => {
        if (position === "before") parent.insertBefore(placeholder, element);
        else parent.insertBefore(placeholder, element.nextSibling);
      },
      undo: () => placeholder.remove(),
      selectionAfterUndo: () => element,
      selectionAfterRedo: () => placeholder,
    });
  }

  async function resizeSelected() {
    const element = selectedOrNotify();
    if (!element) return;
    const before = {
      width: element.style.width,
      minHeight: element.style.minHeight,
    };
    const width = await askUser({
      title: "调整模块宽度",
      label: "输入宽度，例如 100%、50%、600px 或 auto",
      defaultValue: element.style.width || "100%",
      multiline: false,
    });
    if (width === null) return;
    const minHeight = await askUser({
      title: "调整模块高度",
      label: "输入最小高度，例如 240px 或 auto",
      defaultValue: element.style.minHeight || "auto",
      multiline: false,
    });
    if (minHeight === null) return;
    const after = { width: width.trim(), minHeight: minHeight.trim() };
    const target = getDescriptor(element);
    const change = makeChange(
      "resize",
      target,
      `宽度调整为 ${after.width || "默认"}，最小高度调整为 ${after.minHeight || "默认"}。`,
      { before, after },
    );
    executeCommand({
      element,
      change,
      redo: () => {
        element.style.width = after.width === "auto" ? "" : after.width;
        element.style.minHeight =
          after.minHeight === "auto" ? "" : after.minHeight;
      },
      undo: () => {
        element.style.width = before.width;
        element.style.minHeight = before.minHeight;
      },
    });
  }

  async function annotateSelected() {
    const element = selectedOrNotify();
    if (!element) return;
    const instruction = await askUser({
      title: "添加AI修改说明",
      label: "告诉AI这个模块需要怎样修改",
      defaultValue: "请调整这个模块的内容、样式或交互",
      multiline: true,
    });
    if (!instruction?.trim()) return;
    const target = getDescriptor(element);
    const change = makeChange(
      "annotate",
      target,
      instruction.trim(),
      { instruction: instruction.trim() },
    );
    executeCommand({
      element,
      change,
      redo: () => {},
      undo: () => {},
    });
  }

  function buildChangeSet() {
    const counts = state.changes.reduce((result, change) => {
      result[change.action] = (result[change.action] || 0) + 1;
      return result;
    }, {});
    return {
      format: "web-ai-change-set/v1",
      generatedAt: new Date().toISOString(),
      page: {
        url: location.href,
        title: document.title,
        viewport: {
          width: window.innerWidth,
          height: window.innerHeight,
          scrollX: Math.round(window.scrollX),
          scrollY: Math.round(window.scrollY),
        },
      },
      summary: {
        total: state.changes.length,
        counts,
      },
      instructions: [
        "请根据以下页面修改记录更新真实项目源代码，而不是只修改运行时DOM。",
        "删除模块时同步清理无用样式、事件、状态和数据计算。",
        "涉及接口或数据库时保留现有数据契约，并标明需要调整的字段。",
        "完成后重新构建并逐项验证修改结果。",
      ],
      changes: state.changes.map((change, index) => ({
        order: index + 1,
        ...change,
      })),
    };
  }

  function buildMarkdown() {
    const changeSet = buildChangeSet();
    const lines = [
      "# Web页面AI修改单",
      "",
      `- 页面：${changeSet.page.title}`,
      `- 地址：${changeSet.page.url}`,
      `- 修改数量：${changeSet.summary.total}`,
      "",
      "## 执行要求",
      ...changeSet.instructions.map((item) => `- ${item}`),
      "",
      "## 修改明细",
    ];
    changeSet.changes.forEach((change) => {
      lines.push(
        `### ${change.order}. ${change.actionLabel}`,
        `- 模块：${change.target.moduleName}`,
        `- 定位：${change.target.selector}`,
        `- 要求：${change.detail}`,
        `- 当前文字：${change.target.text || "（无）"}`,
        `- 上下文：上方“${change.target.context.previous || "无"}”；下方“${change.target.context.next || "无"}”`,
        "",
      );
    });
    return lines.join("\n");
  }

  async function copyInstructions() {
    if (!state.changes.length) return showToast("请先记录至少一项修改");
    const text = buildMarkdown();
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const textarea = document.createElement("textarea");
      textarea.value = text;
      textarea.style.position = "fixed";
      textarea.style.opacity = "0";
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand("copy");
      textarea.remove();
    }
    showToast("AI修改指令已复制，可直接粘贴给AI");
  }

  function exportChangeSet() {
    if (!state.changes.length) return showToast("请先记录至少一项修改");
    const payload = JSON.stringify(buildChangeSet(), null, 2);
    const blob = new Blob([payload], { type: "application/json;charset=utf-8" });
    const link = document.createElement("a");
    const host = (location.hostname || "local-page").replace(/[^a-z0-9.-]/gi, "-");
    const date = new Date().toISOString().slice(0, 19).replaceAll(":", "-");
    link.href = URL.createObjectURL(blob);
    link.download = `${host}-AI页面修改单-${date}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    showToast("结构化修改单已导出");
  }

  function autosave() {
    try {
      chrome.storage?.local?.set({
        [`web-ai-editor:${location.href}`]: buildChangeSet(),
      });
    } catch {
      // 页面预览环境没有扩展存储时，继续使用当前会话。
    }
  }

  function clearHistory() {
    if (!state.history.length) return showToast("当前没有修改记录");
    if (!window.confirm("撤回本次会话中的全部页面修改并清空记录？")) return;
    while (state.history.length) {
      const command = state.history.pop();
      command.undo();
    }
    state.redo = [];
    state.changes = [];
    setSelected(null);
    render();
    autosave();
    showToast("已恢复进入修改模式前的页面");
  }

  function render() {
    const descriptor = state.selected?.isConnected
      ? getDescriptor(state.selected)
      : null;
    selectionName.textContent = descriptor?.moduleName || "移动鼠标并点击页面模块";
    selectionMeta.innerHTML = descriptor
      ? `<span>${escapeHtml(descriptor.tag)}</span><span>${descriptor.position.width} × ${descriptor.position.height}</span><span>${escapeHtml(descriptor.selector)}</span>`
      : "<span>按住 Alt / Option 可精确选择元素</span>";
    changeCount.textContent = String(state.changes.length);
    modeStatus.textContent = state.enabled ? "选择模式" : "页面操作模式";

    const selectButton = $('[data-action="toggle-select"]');
    selectButton.classList.toggle("is-active", state.enabled);
    selectButton.textContent = state.enabled ? "◉ 选择" : "○ 已暂停";
    $$('[data-needs-selection]').forEach((button) => {
      button.disabled = !descriptor;
    });
    $$('[data-action="undo"]').forEach((button) => {
      button.disabled = state.history.length === 0;
    });
    $$('[data-action="redo"]').forEach((button) => {
      button.disabled = state.redo.length === 0;
    });

    if (!state.changes.length) {
      changeList.innerHTML = `
        <div class="wae-change-empty">
          <div>点击页面中的模块开始修改。<br>所有操作都会记录，并支持撤回和重做。</div>
        </div>`;
      return;
    }
    changeList.innerHTML = state.changes
      .map(
        (change, index) => `
          <article class="wae-change-item">
            <div class="wae-change-top">
              <span class="wae-change-index">${index + 1}</span>
              <span class="wae-change-action">${escapeHtml(change.actionLabel)}</span>
              <span class="wae-change-time">${new Date(change.createdAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}</span>
            </div>
            <div class="wae-change-target">${escapeHtml(change.target.moduleName)}</div>
            <div class="wae-change-detail">${escapeHtml(change.detail)}</div>
          </article>`,
      )
      .join("");
  }

  function toggleSelectionMode() {
    state.enabled = !state.enabled;
    setHovered(null);
    render();
    showToast(state.enabled ? "已进入选择模式" : "已暂停，可操作原页面");
  }

  function closeEditor(force = false) {
    if (!force && state.changes.length) {
      const confirmed = window.confirm(
        "关闭修改模式？修改记录已自动保存在扩展中，建议先复制AI指令或导出修改单。",
      );
      if (!confirmed) return;
    }
    if (state.promptResolver) closePrompt(null);
    setHovered(null);
    state.selected?.classList.remove("wae-page-selected");
    document.removeEventListener("pointermove", onPointerMove, true);
    document.removeEventListener("click", onPageClick, true);
    document.removeEventListener("keydown", onKeyDown, true);
    root.remove();
    delete window[INSTANCE_KEY];
    try {
      chrome.runtime?.sendMessage?.({ type: "WEB_AI_EDITOR_CLOSED" });
    } catch {
      // 页面预览环境忽略扩展通信。
    }
  }

  function toggleEditor() {
    const hidden = root.style.display === "none";
    root.style.display = hidden ? "" : "none";
    state.selected?.classList.toggle("wae-page-selected", hidden);
    setHovered(null);
  }

  function onPointerMove(event) {
    if (!state.enabled) return;
    const rawTarget = event.target instanceof Element ? event.target : null;
    if (!rawTarget || isEditorElement(rawTarget)) return setHovered(null);
    setHovered(resolveSelectionTarget(rawTarget, event));
  }

  function onPageClick(event) {
    if (!state.enabled) return;
    const rawTarget = event.target instanceof Element ? event.target : null;
    if (!rawTarget || isEditorElement(rawTarget)) return;
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    setSelected(resolveSelectionTarget(rawTarget, event));
  }

  function onKeyDown(event) {
    const active = document.activeElement;
    const isTyping =
      active?.matches?.("input,textarea,select,[contenteditable='true']") ||
      isEditorElement(active);
    if (modal.classList.contains("is-open")) {
      if (event.key === "Escape") {
        event.preventDefault();
        closePrompt(null);
      } else if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
        event.preventDefault();
        closePrompt(modalInput.value);
      }
      return;
    }
    if (event.key === "Escape") {
      setSelected(null);
      setHovered(null);
      return;
    }
    if (isTyping) return;
    const commandKey = event.metaKey || event.ctrlKey;
    if (commandKey && event.key.toLowerCase() === "z") {
      event.preventDefault();
      if (event.shiftKey) redo();
      else undo();
    }
  }

  root.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button || button.disabled) return;
    const action = button.dataset.action;
    const handlers = {
      "toggle-select": toggleSelectionMode,
      remove: removeSelected,
      "edit-text": editSelectedText,
      "move-up": () => moveSelected("up"),
      "move-down": () => moveSelected("down"),
      duplicate: duplicateSelected,
      "insert-before": () => insertPlaceholder("before"),
      "insert-after": () => insertPlaceholder("after"),
      resize: resizeSelected,
      annotate: annotateSelected,
      undo,
      redo,
      copy: copyInstructions,
      export: exportChangeSet,
      clear: clearHistory,
      close: closeEditor,
      "modal-cancel": () => closePrompt(null),
      "modal-confirm": () => closePrompt(modalInput.value),
    };
    handlers[action]?.();
  });

  document.addEventListener("pointermove", onPointerMove, true);
  document.addEventListener("click", onPageClick, true);
  document.addEventListener("keydown", onKeyDown, true);

  window[INSTANCE_KEY] = {
    toggle: toggleEditor,
    close: () => closeEditor(true),
    getChangeSet: buildChangeSet,
    getMarkdown: buildMarkdown,
  };
  render();
  showToast("页面修改器已开启，点击任意模块开始修改");
})();
