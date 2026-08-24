"use strict";

const state = {
  presets: [],
  templateWorkflows: [],
  spec: {},
  selectedPresetId: null,
  file: null,
  analysis: null,
  capabilities: null,
  verificationCapabilities: null,
  citationTask: null,
  citationContext: null,
  citationRequested: false,
  aiSession: null,
  sessionSubmitted: false,
  sessionHandoff: null,
  visualPollTimer: null,
  runtime: null,
  runtimePollTimer: null,
  verification: {
    visual_enabled: false,
    visual_model: "auto",
    custom_model_id: "",
    render_method: "auto",
  },
  selectedHeadingLevel: 1,
  customHeadingLevels: new Set(),
};

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));
const PAGE_SIZE_MM = {
  A4: { width_mm: 210, height_mm: 297 },
  Letter: { width_mm: 215.9, height_mm: 279.4 },
};

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function getPath(object, path) {
  return path.split(".").reduce((value, key) => value?.[key], object);
}

function setPath(object, path, value) {
  const parts = path.split(".");
  let cursor = object;
  parts.slice(0, -1).forEach((part) => {
    cursor[part] ??= {};
    cursor = cursor[part];
  });
  cursor[parts.at(-1)] = value;
}

function merge(base, override) {
  const result = clone(base || {});
  Object.entries(override || {}).forEach(([key, value]) => {
    if (value === null || value === undefined || value === "") return;
    if (Array.isArray(value)) {
      if (value.length) result[key] = clone(value);
    } else if (typeof value === "object") {
      result[key] = merge(result[key] || {}, value);
    } else {
      result[key] = value;
    }
  });
  return result;
}

function setStatus(message, isError = false) {
  const status = $("#statusText");
  status.textContent = message;
  status.style.color = isError ? "#ffb4aa" : "";
}

function selectedPreset() {
  if (state.selectedPresetId === "custom-template" && state.customTemplate) {
    return {
      id: "custom-template",
      name: `📄 已识别模板：${state.customTemplate.name}`,
      mode: "custom-template",
      authority: state.spec.authority,
    };
  }
  return [...state.presets, ...state.templateWorkflows].find((item) => item.id === state.selectedPresetId);
}

function isOfficialTemplateWorkflow() {
  return selectedPreset()?.mode === "official-template";
}

function mergeIntoSelected(patch) {
  const selected = selectedPreset() || {};
  const result = merge(selected, patch || {});
  if (selected.mode === "official-template") {
    ["id", "name", "mode", "template_required", "official_url", "workflow", "notes"].forEach((key) => {
      result[key] = clone(selected[key]);
    });
  }
  return result;
}

function adoptExtractedTemplateSpec(inferredSpec) {
  const extracted = clone(inferredSpec || {});
  extracted.schema_version ??= "1.0";
  extracted.id = "distilled-template";
  extracted.name = "导入模板提取结果";
  extracted.mode = "parameterized";
  extracted.template_required = false;
  extracted.page ??= {};
  extracted.body ??= {};
  extracted.lists ??= {};
  extracted.headings = Array.isArray(extracted.headings) ? extracted.headings : [];
  extracted.captions ??= {};
  extracted.headers_footers ??= {};
  extracted.headers_footers.preserve_existing = inferredSpec?.headers_footers?.preserve_existing ?? true;
  if (inferredSpec?.headers_footers?.header?.enabled && inferredSpec.headers_footers.header.text) {
    extracted.headers_footers.header = { ...inferredSpec.headers_footers.header, enabled: true };
  } else {
    extracted.headers_footers.header = { enabled: false, text: "" };
  }
  if (inferredSpec?.headers_footers?.footer?.enabled && inferredSpec.headers_footers.footer.text) {
    extracted.headers_footers.footer = { ...inferredSpec.headers_footers.footer, enabled: true };
  } else {
    extracted.headers_footers.footer = { enabled: false, text: "" };
  }
  extracted.page_numbers ??= { enabled: false };
  extracted.table_of_contents ??= { enabled: false };
  return extracted;
}

function renderPresetSelect() {
  const select = $("#presetSelect");
  if (!select) return;
  select.innerHTML = "";
  if (state.customTemplate) {
    const customGroup = document.createElement("optgroup");
    customGroup.label = "当前已识别的模板规范";
    const customOpt = document.createElement("option");
    customOpt.value = "custom-template";
    customOpt.textContent = `📄 已识别模板：${state.customTemplate.name}`;
    customGroup.append(customOpt);
    select.append(customGroup);
  }
  const appendGroup = (label, items) => {
    if (!items || !items.length) return;
    const group = document.createElement("optgroup");
    group.label = label;
    items.forEach((preset) => {
      const option = document.createElement("option");
      option.value = preset.id;
      option.textContent = preset.name;
      group.append(option);
    });
    select.append(group);
  };
  appendGroup("可直接应用的格式预设", state.presets);
  if (state.templateWorkflows?.length) appendGroup("官方模板工作流", state.templateWorkflows);
  if (state.selectedPresetId) {
    select.value = state.selectedPresetId;
  }
}

async function loadPresets() {
  const response = await fetch("/api/presets");
  if (!response.ok) throw new Error("预设载入失败");
  const data = await response.json();
  state.presets = data.presets || [];
  state.templateWorkflows = data.template_workflows || [];
  state.selectedPresetId = state.presets[0]?.id;
  renderPresetSelect();
  applyPreset(state.selectedPresetId);
  setStatus(`已载入 ${state.presets.length} 个格式预设`);
}

async function loadCapabilities() {
  const [applicationResponse, verificationResponse] = await Promise.all([
    fetch("/api/capabilities"),
    fetch("/api/verification/capabilities"),
  ]);
  if (!applicationResponse.ok || !verificationResponse.ok) throw new Error("写入或验收能力载入失败");
  state.capabilities = await applicationResponse.json();
  state.verificationCapabilities = await verificationResponse.json();
}

function applyPreset(id) {
  if (id === "custom-template" && state.customTemplate) {
    state.selectedPresetId = "custom-template";
    if (state.customTemplate.result?.inferred_spec) {
      state.spec = adoptExtractedTemplateSpec(state.customTemplate.result.inferred_spec);
    }
    state.spec.name = `模板提取规范：${state.customTemplate.name}`;
    state.spec.template_required = true;
    state.spec.authority = {
      level: "uploaded-template",
      sources: [state.customTemplate.name, state.aiSession?.input?.name || "目标文档"],
      override_note: `已通过 AI 双轨印证从【${state.customTemplate.name}】精确提取格式规范，并将应用于目标文档【${state.aiSession?.input?.name || "目标文档"}】。`,
    };
    renderPresetSelect();
    renderAll();
    setStatus(`已切换至：${state.customTemplate.name}`);
    return;
  }
  const previous = selectedPreset();
  state.selectedPresetId = id;
  const preset = selectedPreset();
  if (!preset) return;
  if (!state.aiSession && previous && previous.id !== preset.id && (previous.mode === "official-template" || preset.mode === "official-template")) {
    state.file = null;
    $("#fileInput").value = "";
    $("#fileName").textContent = "选择 DOCX / DOTX";
  }
  state.spec = clone(preset);
  if (!state.aiSession) {
    state.analysis = null;
    state.citationTask = null;
    state.citationContext = null;
    state.citationRequested = false;
  }
  state.customHeadingLevels.clear();
  renderPresetSelect();
  renderAll();
  setStatus(`已选择：${preset.name}`);
}

const CHINESE_FONT_SIZES = [
  { name: "初号", pt: 42 },
  { name: "小初", pt: 36 },
  { name: "一号", pt: 26 },
  { name: "小一", pt: 24 },
  { name: "二号", pt: 22 },
  { name: "小二", pt: 18 },
  { name: "三号", pt: 16 },
  { name: "小三", pt: 15 },
  { name: "四号", pt: 14 },
  { name: "小四", pt: 12 },
  { name: "五号", pt: 10.5 },
  { name: "小五", pt: 9 },
  { name: "六号", pt: 7.5 },
  { name: "小六", pt: 6.5 },
  { name: "七号", pt: 5.5 },
  { name: "八号", pt: 5 },
];

function findChineseFontSizeOption(pt) {
  if (pt === null || pt === undefined || pt === "") return "12";
  const num = Number(pt);
  const match = CHINESE_FONT_SIZES.find((item) => Math.abs(item.pt - num) < 0.15);
  return match ? String(match.pt) : "custom";
}

function convertFromMm(mm, targetUnit, fontSizePt = 12) {
  const mmVal = Math.round(Number(mm) || 0);
  const pt = Math.round(Number(fontSizePt) || 12);
  const ptVal = Math.round(mmVal * 72 / 25.4);
  if (targetUnit === "chars") {
    return Math.round(ptVal / (pt || 12));
  }
  if (targetUnit === "pt") {
    return ptVal;
  }
  if (targetUnit === "cm") {
    return Math.round(mmVal / 10);
  }
  return mmVal; // mm
}

function convertToMm(value, fromUnit, fontSizePt = 12) {
  const val = Math.round(Number(value) || 0);
  const pt = Math.round(Number(fontSizePt) || 12);
  if (fromUnit === "chars") {
    return Math.round(val * pt * 25.4 / 72);
  }
  if (fromUnit === "pt") {
    return Math.round(val * 25.4 / 72);
  }
  if (fromUnit === "cm") {
    return Math.round(val * 10);
  }
  return val; // mm
}

function formatIndentSummary(mm, fontSizePt = 12) {
  const mmVal = Math.round(Number(mm) || 0);
  const pt = Math.round(Number(fontSizePt) || 12);
  const chars = Math.round(mmVal / ((pt || 12) * 25.4 / 72));
  const ptVal = Math.round(mmVal * 72 / 25.4);
  return `换算：${chars} 字符 = ${ptVal} 磅 (${mmVal} mm)`;
}

function updateIndentUI(prefix, mmValue, fontSizePt, workflowLocked) {
  const unitSelect = prefix === "body" ? $("#bodyFirstLineIndentUnitSelect") : $("#referenceHangingIndentUnitSelect");
  const valueInput = prefix === "body" ? $("#bodyFirstLineIndentValueInput") : $("#referenceHangingIndentValueInput");
  const label = prefix === "body" ? $("#bodyFirstLineIndentValueLabel") : $("#referenceHangingIndentValueLabel");
  const hint = prefix === "body" ? $("#bodyIndentConversionHint") : $("#referenceIndentConversionHint");
  if (!unitSelect || !valueInput) return;
  const currentUnit = unitSelect.value || "chars";
  const unitNames = { chars: "字符", pt: "磅 (pt)", mm: "毫米 (mm)", cm: "厘米 (cm)" };
  if (label) label.textContent = `${prefix === "body" ? "首行缩进" : "悬挂缩进"} (${unitNames[currentUnit] || "字符"})`;
  valueInput.value = convertFromMm(mmValue, currentUnit, fontSizePt);
  valueInput.step = "1";
  valueInput.disabled = workflowLocked;
  unitSelect.disabled = workflowLocked;
  if (hint) hint.textContent = formatIndentSummary(mmValue, fontSizePt);
}

function explicitHeading(level) {
  return (state.spec.headings || []).find((item) => Number(item.level) === Number(level));
}

function nearestHeading(level) {
  return (state.spec.headings || [])
    .filter((item) => Number.isFinite(Number(item.level)))
    .sort((left, right) => {
      const distance = Math.abs(Number(left.level) - level) - Math.abs(Number(right.level) - level);
      return distance || Number(right.level) - Number(left.level);
    })[0];
}

function headingForLevel(level, create = false) {
  const existing = explicitHeading(level);
  if (existing) return existing;
  const source = nearestHeading(level);
  const body = state.spec.body || {};
  const heading = {
    level,
    font_latin: source?.font_latin || body.font_latin || "Times New Roman",
    font_east_asia: source?.font_east_asia || body.font_east_asia || "宋体",
    font_size_pt: source?.font_size_pt ?? body.font_size_pt ?? Math.max(10, 17 - level),
    bold: source?.bold ?? true,
    italic: source?.italic ?? false,
    alignment: source?.alignment || "left",
    space_before_pt: source?.space_before_pt ?? 0,
    space_after_pt: source?.space_after_pt ?? 0,
    keep_with_next: true,
  };
  if (!create) return heading;
  state.spec.headings ??= [];
  state.spec.headings.push(heading);
  state.spec.headings.sort((left, right) => Number(left.level) - Number(right.level));
  return heading;
}

function displayedHeading(level) {
  if (state.spec?.id === "distilled-template") {
    return explicitHeading(level) || { level };
  }
  return headingForLevel(level);
}

function headingDefinitionLabel(level, workflowLocked) {
  if (workflowLocked) return "等待模板";
  if (explicitHeading(level)) {
    if (state.customHeadingLevels.has(level)) return "用户已设置";
    if (state.analysis?.route) return "文字已识别";
    if (state.analysis?.source) return "模板已识别";
    return "预设已定义";
  }
  const source = nearestHeading(level);
  return source ? `继承 H${source.level}` : "继承正文";
}

function updateHeadingEditor(workflowLocked) {
  const level = state.selectedHeadingLevel;
  const heading = displayedHeading(level);
  const fields = {
    headingEastAsiaInput: heading.font_east_asia ?? "",
    headingLatinInput: heading.font_latin ?? "",
    headingSizeInput: heading.font_size_pt !== undefined ? Math.round(Number(heading.font_size_pt)) : "",
    headingAlignInput: (heading.alignment === "both" ? "justify" : heading.alignment) || "left",
    headingSpaceBeforeInput: heading.space_before_pt !== undefined ? Math.round(Number(heading.space_before_pt)) : "",
    headingSpaceAfterInput: heading.space_after_pt !== undefined ? Math.round(Number(heading.space_after_pt)) : "",
    headingFirstLineIndentInput: heading.first_line_indent_chars !== undefined ? heading.first_line_indent_chars : (heading.first_line_indent_mm ? Math.round(Number(heading.first_line_indent_mm) / ((Number(heading.font_size_pt) || 12) * 25.4 / 72) * 10) / 10 : ""),
    headingLeftIndentInput: heading.left_indent_chars !== undefined ? heading.left_indent_chars : (heading.left_indent_mm ? Math.round(Number(heading.left_indent_mm) / ((Number(heading.font_size_pt) || 12) * 25.4 / 72) * 10) / 10 : ""),
  };
  Object.entries(fields).forEach(([id, value]) => {
    const input = $(`#${id}`);
    if (input) {
      input.value = value;
      input.disabled = workflowLocked;
    }
  });
  const headingSizeSelect = $("#headingFontSizeSelect");
  if (headingSizeSelect) {
    const opt = findChineseFontSizeOption(heading.font_size_pt);
    headingSizeSelect.value = opt;
    headingSizeSelect.disabled = workflowLocked;
    $("#headingCustomFontSizeRow").hidden = opt !== "custom";
  }
  $("#headingBoldInput").checked = Boolean(heading.bold);
  $("#headingItalicInput").checked = Boolean(heading.italic);
  $("#headingBoldInput").disabled = workflowLocked;
  $("#headingItalicInput").disabled = workflowLocked;
  $$("[data-heading-level]").forEach((button) => {
    const active = Number(button.dataset.headingLevel) === level;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", String(active));
  });
  const badge = $("#headingDefinitionBadge");
  const isDerived = !explicitHeading(level);
  badge.textContent = headingDefinitionLabel(level, workflowLocked);
  badge.className = `badge ${isDerived ? "derived" : "neutral"}`;
}

function updateBodyRoleSummary() {
  const body = state.spec.body || {};
  const bodyFonts = [body.font_east_asia, body.font_latin].filter(Boolean).join(" / ");
  const ptVal = body.font_size_pt ? Math.round(Number(body.font_size_pt)) : "-";
  const cnMatch = CHINESE_FONT_SIZES.find((item) => Math.abs(item.pt - Number(body.font_size_pt)) < 0.6);
  const sizeText = cnMatch ? `${cnMatch.name} (${ptVal} pt)` : `${ptVal} pt`;
  $("#bodyRoleSummary").textContent = bodyFonts
    ? `${bodyFonts} · ${sizeText}`
    : "由模板决定";
}

function updateLineSpacingUI(prefix, spacingObj, workflowLocked) {
  const kindSelect = $(`#${prefix}LineSpacingKindSelect`);
  const input = prefix === "body" ? $("#lineSpacingInput") : $("#referenceSpacingInput");
  const label = $(`#${prefix}LineSpacingLabel`);
  if (!kindSelect || !input || !label) return;
  if (!spacingObj) {
    kindSelect.value = "";
    label.textContent = "行距";
    input.value = "";
    input.disabled = workflowLocked;
    return;
  }
  const kind = spacingObj.kind;
  const val = Number(spacingObj.value ?? spacingObj.value_pt);
  if (kind === "exact") {
    kindSelect.value = "exact";
    label.textContent = "行距 (磅 pt)";
    input.value = String(spacingObj.value_pt ?? 22);
    input.step = "1";
    input.min = "1";
    input.max = "100";
    input.disabled = workflowLocked;
  } else if (kind === "at_least") {
    kindSelect.value = "at_least";
    label.textContent = "最小值 (磅 pt)";
    input.value = String(spacingObj.value_pt ?? 22);
    input.step = "1";
    input.min = "1";
    input.max = "100";
    input.disabled = workflowLocked;
  } else if (kind === "single" || (kind === "multiple" && Math.abs(val - 1.0) < 0.05)) {
    kindSelect.value = "single";
    label.textContent = "行距 (单倍)";
    input.value = "1";
    input.disabled = true;
  } else if (kind === "1.5" || (kind === "multiple" && Math.abs(val - 1.5) < 0.05)) {
    kindSelect.value = "1.5";
    label.textContent = "行距 (1.5 倍)";
    input.value = "1.5";
    input.disabled = true;
  } else if (kind === "double" || (kind === "multiple" && Math.abs(val - 2.0) < 0.05)) {
    kindSelect.value = "double";
    label.textContent = "行距 (2 倍)";
    input.value = "2";
    input.disabled = true;
  } else {
    kindSelect.value = "multiple";
    label.textContent = "行距 (倍数)";
    input.value = String(spacingObj?.value ?? (prefix === "body" ? 1.5 : 1));
    input.step = "0.1";
    input.min = "0.5";
    input.max = "10";
    input.disabled = workflowLocked;
  }
  kindSelect.disabled = workflowLocked;
}

function visualReviewRuntimeReady() {
  return Boolean(state.runtime?.visual_ready);
}

function visualReviewUnavailableMessage() {
  if (!state.runtime) return "正在检测 Microsoft Word、LibreOffice 和 PDF 转图组件，视觉验收暂不可用";
  if (!state.runtime.renderer_ready) return "未检测到 Microsoft Word 或 LibreOffice，无法启用视觉验收";
  if (!state.runtime.pdf_rasterizer?.available) return "缺少 PDF 逐页转图组件，无法启用视觉验收";
  return "";
}

function updateInputs() {
  const workflowLocked = state.sessionSubmitted || (isOfficialTemplateWorkflow() && !state.analysis?.source);
  const runtimeReady = visualReviewRuntimeReady();
  if (!runtimeReady && !state.sessionSubmitted) state.verification.visual_enabled = false;
  $$("[data-spec-path]").forEach((input) => {
    let value = getPath(state.spec, input.dataset.specPath);
    if (input.dataset.specPath?.endsWith("alignment") && value === "both") {
      value = "justify";
      setPath(state.spec, input.dataset.specPath, "justify");
    }
    if (input.type === "checkbox") {
      input.checked = Boolean(value);
    } else if (input.type === "number") {
      input.value = value !== null && value !== undefined && value !== "" ? Math.round(Number(value)) : "";
      input.step = "1";
    } else {
      input.value = value ?? "";
    }
    input.disabled = workflowLocked;
  });

  // Body line spacing
  updateLineSpacingUI("body", state.spec.body?.line_spacing, workflowLocked);
  // Reference line spacing
  updateLineSpacingUI("reference", state.spec.references?.line_spacing, workflowLocked);

  // Body font size select sync
  const bodySizeOpt = findChineseFontSizeOption(state.spec.body?.font_size_pt);
  const bodySizeSelect = $("#bodyFontSizeSelect");
  if (bodySizeSelect) {
    bodySizeSelect.value = bodySizeOpt;
    bodySizeSelect.disabled = workflowLocked;
    $("#bodyCustomFontSizeRow").hidden = (bodySizeOpt !== "custom");
  }

  // Body indent sync
  updateIndentUI("body", state.spec.body?.first_line_indent_mm ?? 0, state.spec.body?.font_size_pt, workflowLocked);

  // Captions font size selects sync
  const figureSizeOpt = findChineseFontSizeOption(state.spec.captions?.figure?.font_size_pt);
  const figureSelect = $("#captionFigureFontSizeSelect");
  if (figureSelect) {
    figureSelect.value = figureSizeOpt;
    figureSelect.disabled = workflowLocked;
    $("#captionFigureCustomSizeRow").hidden = (figureSizeOpt !== "custom");
  }
  const tableSizeOpt = findChineseFontSizeOption(state.spec.captions?.table?.font_size_pt);
  const tableSelect = $("#captionTableFontSizeSelect");
  if (tableSelect) {
    tableSelect.value = tableSizeOpt;
    tableSelect.disabled = workflowLocked;
    $("#captionTableCustomSizeRow").hidden = (tableSizeOpt !== "custom");
  }

  // References font size select sync
  const refSizeOpt = findChineseFontSizeOption(state.spec.references?.font_size_pt);
  const refSelect = $("#referenceFontSizeSelect");
  if (refSelect) {
    refSelect.value = refSizeOpt;
    refSelect.disabled = workflowLocked;
    $("#referenceCustomSizeRow").hidden = (refSizeOpt !== "custom");
  }

  // References indent sync
  updateIndentUI("reference", state.spec.references?.hanging_indent_mm ?? 0, state.spec.references?.font_size_pt, workflowLocked);

  updateBodyRoleSummary();
  $("#referenceSystemValue").textContent = state.spec.references?.citation_system || "未指定";
  const citationButton = $("#citationRequestButton");
  citationButton.disabled = state.citationContext?.status !== "ready" || state.sessionSubmitted;
  citationButton.textContent = state.citationRequested ? "✓ 已开启在正文中引用" : "在正文中引用";
  citationButton.classList.toggle("is-selected", state.citationRequested);
  const visualToggle = $("#visualReviewToggle");
  visualToggle.checked = state.verification.visual_enabled;
  visualToggle.disabled = !runtimeReady || state.sessionSubmitted;
  visualToggle.title = state.sessionSubmitted
    ? "设置已交给 AI，不能再修改视觉验收选项"
    : runtimeReady
      ? "启用后，AI 将使用检测到的本地渲染器逐页验收"
      : visualReviewUnavailableMessage();
  if ($("#visualModelSelect")) {
    $("#visualModelSelect").value = state.verification.visual_model;
    $("#visualModelSelect").disabled = state.sessionSubmitted;
  }
  if ($("#customVisualModelInput")) {
    $("#customVisualModelInput").value = state.verification.custom_model_id;
    $("#customVisualModelInput").disabled = state.sessionSubmitted;
  }
  if ($("#customVisualModelControl")) {
    $("#customVisualModelControl").hidden = state.verification.visual_model !== "custom";
  }
  $("#renderMethodSelect").value = state.verification.render_method || "auto";
  $("#renderMethodSelect").disabled = state.sessionSubmitted;
  $("#presetSelect").disabled = state.sessionSubmitted;
  $("#requirementsText").disabled = state.sessionSubmitted;
  $("#analyzeTextButton").disabled = state.sessionSubmitted;
  updateHeadingEditor(workflowLocked);
  $("#clearDirectInput").disabled = workflowLocked;
}

function renderAuthority() {
  const authority = state.spec.authority || {};
  const isCustom = state.selectedPresetId === "custom-template" || authority.level === "uploaded-template";
  const workflow = !isCustom && isOfficialTemplateWorkflow() ? selectedPreset() : null;
  $("#authorityNote").textContent = authority.override_note || (isCustom ? `已根据【${state.customTemplate?.name || authority.sources?.[0] || "上传模板"}】完成 AI 深度识别与排版规范配置。` : "当前参数来自用户调整或模板识别。");
  $("#modeBadge").textContent = isCustom ? "模板识别" : (state.spec.template_required ? "模板优先" : "参数预设");
  $("#modeBadge").className = `badge ${isCustom ? "success" : (state.spec.template_required ? "warning" : "neutral")}`;

  const workflowPanel = $("#templateWorkflow");
  workflowPanel.hidden = !workflow;
  if (workflow) {
    $("#officialTemplateLink").href = workflow.official_url;
    $("#templateWorkflowSteps").innerHTML = (workflow.workflow || [])
      .map((step) => `<li>${escapeHtml(step)}</li>`)
      .join("");
    const scope = workflow.processing_scope || {};
    $("#templateScopeNote").textContent = scope.start
      ? `处理范围：${scope.start} 至 ${scope.end || "文档末尾"}；封面、声明和 AI 使用说明表保持母版原样。`
      : "";
    $("#localTemplateLinks").innerHTML = (workflow.local_templates || [])
      .map((item) => `<a class="template-link" href="${escapeHtml(item.url)}" download="${escapeHtml(item.filename)}">下载 ${escapeHtml(item.label || item.paper_size || "本地模板")}</a>`)
      .join("");
  } else {
    $("#templateScopeNote").textContent = "";
    $("#localTemplateLinks").innerHTML = "";
  }
}

function previewCopyForPreset(id) {
  if (id === "apa7-student") {
    return {
      layout: "apa",
      language: "en",
      title: "Urban Green Infrastructure and Summer Heat Exposure",
      author: "Morgan Lee",
      affiliation: "Department of Environmental Studies, Northbridge University",
      abstractLabel: "Abstract. ",
      abstract: "This study combines remote sensing, neighborhood sensors, and land-cover data to estimate how green infrastructure changes summer heat exposure.",
      keywordsLabel: "Keywords: ",
      keywords: "green infrastructure, heat exposure, spatial analysis",
      headings: ["Urban Heat Exposure", "Data and Measures", "Sensor Calibration", "Robustness Checks", "Local sensitivity analysis."],
      body: [
        "Urban development changes surface energy balance and creates unequal heat exposure across neighborhoods.",
        "Hourly observations were aligned with satellite surface temperature and street-level canopy estimates.",
        "The model included canopy cover, sky-view factor, and impervious surface, with grouped cross-validation.",
        "The direction and magnitude of the main effect remained stable across buffer sizes.",
      ],
      list: ["Neighborhood environmental features", "Hourly heat-exposure measures"],
      equation: "HI = 0.62T + 0.38RH",
      tableTitle: "Model Comparison",
      tableHeaders: ["Model", "RMSE"],
      tableRows: [["Baseline", "2.41"], ["Sensor fusion", "1.76"]],
      figureTitle: "Heat exposure by intervention scenario",
      figureLabels: ["Base", "Trees", "Mixed"],
      figureAlt: "Heat exposure under baseline, tree canopy, and mixed intervention scenarios",
      referencesHeading: "References",
      references: [
        "Oke, T. R., Mills, G., Christen, A., & Voogt, J. A. (2017). Urban climates. Cambridge University Press.",
        "Ziter, C. D., Pedersen, E. J., Kucharik, C. J., & Turner, M. G. (2019). Scale-dependent interactions between tree canopy cover and impervious surfaces reduce daytime urban heat. Proceedings of the National Academy of Sciences, 116(15), 7575–7580.",
      ],
      showAbstract: true,
      twoColumn: false,
    };
  }
  if (id === "mla9-research") {
    return {
      layout: "mla",
      language: "en",
      title: "Green Infrastructure and the Unequal Geography of Urban Heat",
      author: "Morgan Lee",
      affiliation: "Environmental Studies 302 · 17 August 2026",
      headings: ["Urban Heat as a Spatial Problem", "Measuring Neighborhood Exposure", "Sensor Calibration", "Robustness Checks", "Local Sensitivity Analysis"],
      body: [
        "Heat is experienced locally, yet its causes extend across the built environment and the distribution of public investment.",
        "This analysis pairs hourly sensor observations with canopy cover and impervious-surface measures.",
        "Grouped validation separates calibration locations from evaluation locations to limit spatial leakage.",
        "The main effect remains stable across buffer sizes and missing-data treatments.",
      ],
      list: ["Neighborhood environmental features", "Hourly heat-exposure measures"],
      equation: "HI = 0.62T + 0.38RH",
      tableTitle: "Model Comparison",
      tableHeaders: ["Model", "RMSE"],
      tableRows: [["Baseline", "2.41"], ["Sensor fusion", "1.76"]],
      figureTitle: "Heat exposure by intervention scenario",
      figureLabels: ["Base", "Trees", "Mixed"],
      figureAlt: "Heat exposure under three neighborhood intervention scenarios",
      referencesHeading: "Works Cited",
      references: [
        "Oke, T. R., et al. Urban Climates. Cambridge UP, 2017.",
        "Ziter, Carly D., et al. “Scale-Dependent Interactions between Tree Canopy Cover and Impervious Surfaces Reduce Daytime Urban Heat.” Proceedings of the National Academy of Sciences, vol. 116, no. 15, 2019, pp. 7575–80.",
      ],
      showAbstract: false,
      twoColumn: false,
    };
  }
  let h1Text = "第一章  绪论";
  let h2Text = "1.1 研究背景";
  let h3Text = "1.1.1 国内外现状";
  if (state.citationContext?.task?.paragraphs?.length) {
    const rawParagraphs = state.citationContext.task.paragraphs
      .map((p) => p.text?.trim())
      .filter((t) => t && !/\t\d+$/.test(t) && !/\.{3,}\s*\d+$/.test(t));
    const l1 = rawParagraphs.find((t) => /^(?:第[一二三四五六七八九十0-9]+章|\d+\s*绪论|绪论|引言)/.test(t) && t.length < 40);
    if (l1) h1Text = l1;
    const l2 = rawParagraphs.find((t) => /^\d+\.\d+(?:\s+|[^\d\.])/.test(t) && !/^\d+\.\d+\.\d+/.test(t) && t.length < 40);
    if (l2) h2Text = l2;
    const l3 = rawParagraphs.find((t) => /^\d+\.\d+\.\d+(?:\s+|[^\d\.])/.test(t) && t.length < 40);
    if (l3) h3Text = l3;
  }
  const headings = [h1Text, h2Text, h3Text];
  return {
    layout: "chinese",
    language: "zh-CN",
    title: "城市绿色基础设施对夏季热暴露的缓解效应",
    author: "张明　李然",
    affiliation: "城市环境与空间分析实验室",
    abstractLabel: "摘要：",
    abstract: "本研究融合遥感、街区传感器与地表覆盖数据，评估不同尺度绿色基础设施对高温暴露的影响。",
    keywordsLabel: "关键词：",
    keywords: "绿色基础设施；热暴露；空间分析",
    headings,
    body: [
      "快速城市化改变了地表能量平衡。连续观测显示，树冠覆盖与遮阴连通性能够降低午后行人层热暴露。",
      "研究在典型居住区布设微气候节点，并以同一时间窗的卫星地表温度校正传感器观测。",
      "模型纳入树冠率、天空可视因子与不透水面比例，采用分层交叉验证比较预测误差。",
      "在不同缓冲半径与缺失值处理方案下，主要效应方向保持一致。",
    ],
    list: ["街区尺度环境特征", "逐小时热暴露指标"],
    equation: "HI = 0.62T + 0.38RH",
    tableTitle: "模型比较结果",
    tableHeaders: ["模型", "RMSE"],
    tableRows: [["基线模型", "2.41"], ["融合模型", "1.76"]],
    figureTitle: "不同情景下的热暴露指数",
    figureLabels: ["基线", "树冠", "复合"],
    figureAlt: "基线、树冠和复合干预情景的热暴露指数比较",
    referencesHeading: "参考文献",
    references: [
      "[1] 中国标准化研究院. 学术论文编写规则: GB/T 7713.2—2022[S]. 北京: 中国标准出版社, 2022.",
      "[2] OKE T R, MILLS G, CHRISTEN A, et al. Urban Climates[M]. Cambridge: Cambridge University Press, 2017.",
    ],
    showAbstract: true,
    twoColumn: false,
  };
}

function renderPreviewCopy(copy) {
  const text = (selector, value) => { $(selector).textContent = value ?? ""; };
  text("#previewTitle", copy.title);
  text("#previewAuthor", copy.author);
  text("#previewAffiliation", copy.affiliation);
  text("#previewAbstractLabel", copy.abstractLabel);
  text("#previewAbstractText", copy.abstract);
  text("#previewKeywordsLabel", copy.keywordsLabel);
  text("#previewKeywordsText", copy.keywords);
  $(".preview-abstract-block").hidden = !copy.showAbstract;
  copy.headings.slice(0, 3).forEach((value, index) => text(`#previewHeading${index + 1}`, value));
  copy.body.forEach((value, index) => text(`#previewBody${index + 1}`, value));
  const listFallback = copy.language === "zh-CN"
    ? ["空间缓冲区设置", "稳健性检验方案", "结果复核记录"]
    : ["Spatial buffer settings", "Robustness checks", "Result verification"];
  const listItems = [...(copy.list || []), ...listFallback];
  $("#previewList").querySelectorAll("li").forEach((item, index) => { item.textContent = listItems[index] || ""; });
  text("#previewEquation", copy.equation);
  text("#previewTableHeader1", copy.tableHeaders[0]);
  text("#previewTableHeader2", copy.tableHeaders[1]);
  text("#previewTableCell11", copy.tableRows[0][0]);
  text("#previewTableCell12", copy.tableRows[0][1]);
  text("#previewTableCell21", copy.tableRows[1][0]);
  text("#previewTableCell22", copy.tableRows[1][1]);
  copy.figureLabels.forEach((value, index) => text(`#previewFigureLabel${index + 1}`, value));
  $("#previewFigure").setAttribute("aria-label", copy.figureAlt);
  text("#previewReferencesHeading", copy.referencesHeading);
  text("#previewReference1", copy.references[0]);
  text("#previewReference2", copy.references[1]);
  $("#previewMainBody").classList.toggle("two-column", copy.twoColumn);
}

function applyHeadingPreview(element, level, scale) {
  const heading = headingForLevel(level);
  element.style.fontFamily = `"${heading.font_latin || "Times New Roman"}", "${heading.font_east_asia || "宋体"}", serif`;
  element.style.fontSize = `${(Number(heading.font_size_pt) || 12) * scale}pt`;
  element.style.fontWeight = heading.bold ? "700" : "400";
  element.style.fontStyle = heading.italic ? "italic" : "normal";
  element.style.textAlign = heading.alignment || "left";
  element.style.marginTop = `${(Number(heading.space_before_pt) || 0) * scale}pt`;
  element.style.marginBottom = `${(Number(heading.space_after_pt) || 0) * scale}pt`;
  const flChars = Number(heading.first_line_indent_chars) || 0;
  const lChars = Number(heading.left_indent_chars) || 0;
  element.style.textIndent = flChars ? `${flChars}em` : "0";
  element.style.marginLeft = lChars ? `${lChars}em` : "0";
}

function flexAlignment(value) {
  return value === "left" ? "flex-start" : value === "right" ? "flex-end" : "center";
}

function numberOr(value, fallback) {
  return value === null || value === undefined || value === "" ? fallback : Number(value);
}

function romanNumber(value) {
  const pairs = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let number = Math.max(1, Math.floor(value));
  let result = "";
  pairs.forEach(([unit, marker]) => {
    while (number >= unit) {
      result += marker;
      number -= unit;
    }
  });
  return result;
}

function chineseNumber(value) {
  const digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
  const number = Math.max(1, Math.min(999, Math.floor(value)));
  if (number < 10) return digits[number];
  if (number < 100) return `${number < 20 ? "" : digits[Math.floor(number / 10)]}十${number % 10 ? digits[number % 10] : ""}`;
  const remainder = number % 100;
  return `${digits[Math.floor(number / 100)]}百${remainder < 10 && remainder ? "零" : ""}${remainder ? chineseNumber(remainder) : ""}`;
}

function alphaNumber(value, upper = true) {
  let number = Math.max(1, Math.floor(value));
  let result = "";
  while (number > 0) {
    number -= 1;
    result = String.fromCharCode(65 + (number % 26)) + result;
    number = Math.floor(number / 26);
  }
  return upper ? result : result.toLowerCase();
}

function listMarker(style, value) {
  const chinese = chineseNumber(value);
  const formats = {
    "decimal-period": `${value}.`,
    "decimal-parenthesis": `${value})`,
    "decimal-parentheses": `(${value})`,
    "decimal-fullwidth-parentheses": `（${value}）`,
    "chinese-period": `${chinese}、`,
    "chinese-parentheses": `（${chinese}）`,
    "upper-alpha-period": `${alphaNumber(value)}.`,
    "lower-alpha-parenthesis": `${alphaNumber(value, false)})`,
    "upper-roman-period": `${romanNumber(value)}.`,
  };
  return formats[style] || formats["decimal-period"];
}

function previewBandAlignment(config, pageNumbers, location) {
  const hasText = Boolean(config.enabled && config.text);
  const hasPageNumber = Boolean(pageNumbers.enabled && pageNumbers.location === location);
  if (hasText && hasPageNumber) return "space-between";
  if (hasPageNumber) return flexAlignment(pageNumbers.alignment);
  if (hasText) return flexAlignment(config.alignment);
  return "center";
}

function renderPreview() {
  const page = state.spec.page || {};
  const copy = previewCopyForPreset(state.selectedPresetId);
  renderPreviewCopy(copy);
  const body = state.spec.body || {};
  const paper = $("#paperPreview");
  const width = Number(page.width_mm) || 210;
  const height = Number(page.height_mm) || 297;
  paper.dataset.previewLayout = copy.layout;
  paper.style.setProperty("--paper-width", width);
  paper.style.setProperty("--paper-height", height);
  paper.style.setProperty("--margin-top", numberOr(page.margin_top_mm, 20));
  paper.style.setProperty("--margin-right", numberOr(page.margin_right_mm, 20));
  paper.style.setProperty("--margin-bottom", numberOr(page.margin_bottom_mm, 20));
  paper.style.setProperty("--margin-left", numberOr(page.margin_left_mm, 20));
  const renderedWidth = paper.getBoundingClientRect().width || 680;
  const scale = Math.min(1, renderedWidth / (width * 96 / 25.4));
  paper.style.setProperty("--preview-scale", scale);
  const content = $("#paperContent");
  content.lang = copy.language;
  const bodySize = Number(body.font_size_pt) || 12;
  content.style.fontFamily = `"${body.font_latin || "Times New Roman"}", "${body.font_east_asia || "宋体"}", serif`;
  content.style.fontSize = `${bodySize * scale}pt`;
  const spacingKind = body.line_spacing?.kind;
  let lineHeightCss = "1.5";
  if (spacingKind === "exact" || spacingKind === "at_least") {
    lineHeightCss = `${(Number(body.line_spacing.value_pt) || 22) * scale}pt`;
  } else if (spacingKind === "single") {
    lineHeightCss = "1.0";
  } else if (spacingKind === "1.5") {
    lineHeightCss = "1.5";
  } else if (spacingKind === "double") {
    lineHeightCss = "2.0";
  } else if (spacingKind === "multiple") {
    lineHeightCss = String(body.line_spacing.value || 1.5);
  } else {
    lineHeightCss = "1.5";
  }
  content.style.lineHeight = lineHeightCss;
  content.querySelectorAll(".preview-body").forEach((paragraph) => {
    paragraph.style.textAlign = body.alignment || (copy.twoColumn ? "justify" : "left");
    paragraph.style.textIndent = `${(Number(body.first_line_indent_mm) || 0) * scale}mm`;
    paragraph.style.marginTop = `${(Number(body.space_before_pt) || 0) * scale}pt`;
    paragraph.style.marginBottom = `${(Number(body.space_after_pt) || 0) * scale}pt`;
  });
  content.querySelectorAll("[data-preview-heading]").forEach((element) => {
    applyHeadingPreview(element, Number(element.dataset.previewHeading), scale);
  });
  const numberedList = state.spec.lists?.numbered || {};
  const listStart = Math.max(1, Number(numberedList.start) || 1);
  $("#previewList").querySelectorAll("li").forEach((item, index) => {
    item.dataset.marker = listMarker(numberedList.style || "decimal-period", listStart + index);
  });

  const captions = state.spec.captions || {};
  const figure = captions.figure || {};
  const table = captions.table || {};
  const figureCaption = $("#previewFigureCaption");
  const tableCaption = $("#previewTableCaption");
  const separator = copy.language === "zh-CN" ? "　" : ". ";
  figureCaption.textContent = `${figure.label || (copy.language === "zh-CN" ? "图" : "Figure")} 1${separator}${copy.figureTitle}`;
  tableCaption.textContent = `${table.label || (copy.language === "zh-CN" ? "表" : "Table")} 1${separator}${copy.tableTitle}`;
  [figureCaption, tableCaption].forEach((element) => {
    const tokens = element === figureCaption ? figure : table;
    element.style.fontFamily = `"${tokens.font_latin || body.font_latin || "Times New Roman"}", "${tokens.font_east_asia || body.font_east_asia || "宋体"}", serif`;
    element.style.fontSize = `${(Number(tokens.font_size_pt) || 9) * scale}pt`;
    element.style.textAlign = tokens.alignment || "center";
    element.style.textIndent = "0";
  });
  const figureBox = $("#previewFigure");
  if (figure.position === "above") figureBox.before(figureCaption);
  else figureBox.after(figureCaption);
  const tableBox = $(".sample-table");
  if (table.position === "below") tableBox.after(tableCaption);
  else tableBox.before(tableCaption);

  const headerConfig = state.spec.headers_footers?.header || {};
  const footerConfig = state.spec.headers_footers?.footer || {};
  $("#previewHeaderText").textContent = headerConfig.enabled && headerConfig.text ? headerConfig.text : "";
  $("#previewFooterText").textContent = footerConfig.enabled && footerConfig.text ? footerConfig.text : "";
  const pageNumbers = state.spec.page_numbers || {};
  let pageText = "";
  if (pageNumbers.enabled) {
    pageText = pageNumbers.format === "page-x-of-y" ? "Page 1 of 8" : pageNumbers.format === "page-number" ? "Page 1" : "1";
  }
  $("#previewHeaderPage").textContent = pageNumbers.location === "header" ? pageText : "";
  $("#previewFooterPage").textContent = pageNumbers.location === "footer" ? pageText : "";
  $(".paper-header").style.justifyContent = previewBandAlignment(headerConfig, pageNumbers, "header");
  $(".paper-footer").style.justifyContent = previewBandAlignment(footerConfig, pageNumbers, "footer");

  const references = state.spec.references || {};
  const refKind = references.line_spacing?.kind;
  let refLineHeightCss = "1.5";
  if (refKind === "exact" || refKind === "at_least") {
    refLineHeightCss = `${(Number(references.line_spacing.value_pt) || 22) * scale}pt`;
  } else if (refKind === "single") {
    refLineHeightCss = "1.0";
  } else if (refKind === "1.5") {
    refLineHeightCss = "1.5";
  } else if (refKind === "double") {
    refLineHeightCss = "2.0";
  } else if (refKind === "multiple") {
    refLineHeightCss = String(references.line_spacing.value || 1);
  } else {
    refLineHeightCss = "1.5";
  }
  content.querySelectorAll(".preview-reference").forEach((reference) => {
    reference.style.fontSize = `${(Number(references.font_size_pt) || 9) * scale}pt`;
    reference.style.textAlign = references.alignment || "left";
    const hanging = Number(references.hanging_indent_mm) || 0;
    reference.style.marginLeft = `${hanging * scale}mm`;
    reference.style.textIndent = `${-hanging * scale}mm`;
    reference.style.lineHeight = refLineHeightCss;
  });
  $("#pageLabel").textContent = isOfficialTemplateWorkflow() && !Object.keys(page).length
    ? "由官方模板决定"
    : `${page.size || "自定义"} · ${width} × ${height} mm`;
}

function renderAll() {
  updateInputs();
  renderAuthority();
  renderPreview();
  renderRuntime();
  const sessionMode = Boolean(state.aiSession);
  const applyButton = $("#applyButton");
  applyButton.textContent = sessionMode
    ? (state.sessionSubmitted ? "已交给 AI 继续修改" : "确认设置并交给 AI")
    : "请先启动目标文档会话";
  applyButton.disabled = sessionMode
    ? !state.file || state.sessionSubmitted
    : true;

  const citationContext = state.citationContext;
  const citationTask = citationContext?.task;
  const paragraphCount = citationContext?.paragraph_count ?? citationTask?.paragraphs?.length ?? 0;
  const referenceCount = citationContext?.reference_count ?? citationTask?.references?.length ?? 0;
  $("#citationParagraphCount").textContent = String(paragraphCount);
  $("#citationReferenceCount").textContent = String(referenceCount);
  const citationBadge = $("#citationReadyBadge");
  if (citationContext?.status === "ready") {
    citationBadge.textContent = "AI 已预处理";
    citationBadge.className = "badge success";
    $("#citationWorkflowStatus").textContent = state.citationRequested
      ? "已请求 AI 分析；AI 只会在段落指纹和文献编号校验通过的位置写入引用。"
      : "正文锚点和文献条目已预处理；点击按钮后才会请求 AI 分析引用位置。";
  } else if (citationContext?.status === "unavailable") {
    citationBadge.textContent = "不可用";
    citationBadge.className = "badge warning";
    $("#citationWorkflowStatus").textContent = citationContext.reason || "当前文件不能生成引用位置上下文。";
  } else {
    citationBadge.textContent = "等待文档";
    citationBadge.className = "badge neutral";
    $("#citationWorkflowStatus").textContent = "载入 DOCX 后可请求 AI 分析正文锚点和参考文献。";
  }
  const visual = state.aiSession?.visual_verification;
  if (!visualReviewRuntimeReady()) {
    $("#visualReviewStatus").textContent = visualReviewUnavailableMessage();
  } else if (!state.verification.visual_enabled) {
    $("#visualReviewStatus").textContent = "视觉验收未开启";
  } else if (visual) {
    $("#visualReviewStatus").textContent = visualReviewMessage(visual);
  } else if (state.runtime && !selectedRuntime().available) {
    const requested = state.verification.render_method || "auto";
    $("#visualReviewStatus").textContent = requested === "auto"
      ? "未检测到 Word 或 LibreOffice；视觉验收将跳过并说明原因。"
      : `已选择 ${runtimeLabel(requested)}，但当前不可用；视觉验收将跳过且不会改用其他方式。`;
  } else if (sessionMode) {
    $("#visualReviewStatus").textContent = state.sessionSubmitted
      ? "视觉验收已请求，等待 AI 渲染并回传最终结果。"
      : "AI 修改后将渲染并逐页验收；未确认图片输入能力时会跳过并说明原因。";
  } else {
    $("#visualReviewStatus").textContent = "本地下载不会调用 AI；应用后会标记为已跳过。请使用 AI 交接流程执行视觉验收。";
  }
  applyButton.title = sessionMode
    ? (state.sessionSubmitted ? "设置已提交，等待 AI 使用交接方案修改文档" : "提交确认的参数、写入方法和验收选项给 AI")
    : (isOfficialTemplateWorkflow() ? `${selectedPreset()?.name || "当前模板"} 必须在母版中排版，不能作为数值预设直接套用` : "");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatRuntimeBytes(value) {
  const bytes = Number(value) || 0;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function runtimeLabel(method) {
  return {
    word: "Microsoft Word",
    libreoffice: "LibreOffice",
  }[method] || "自动选择";
}

function selectedRuntime() {
  const requested = state.verification.render_method || "auto";
  const selected = requested === "auto" ? state.runtime?.preferred : requested;
  const entry = selected ? state.runtime?.[selected] : null;
  return {
    requested,
    selected,
    available: Boolean(entry?.available),
    entry,
  };
}

function renderRuntime() {
  const badge = $("#renderRuntimeBadge");
  const status = $("#renderRuntimeStatus");
  const progress = $("#renderRuntimeProgress");
  const progressText = $("#renderRuntimeProgressText");
  const progressBytes = $("#renderRuntimeProgressBytes");
  const progressBar = progress?.parentElement;
  const actions = $("#renderRuntimeActions");
  const officialLink = $("#libreOfficeOfficialLink");
  const downloadButton = $("#aiDownloadRuntimeButton");
  const methodSelect = $("#renderMethodSelect");
  const wordBadge = $("#wordRuntimeBadge");
  const libreOfficeBadge = $("#libreOfficeRuntimeBadge");
  if (!badge || !status || !progress || !progressText || !progressBytes || !actions || !officialLink || !downloadButton || !methodSelect || !wordBadge || !libreOfficeBadge) return;
  methodSelect.value = state.verification.render_method || "auto";
  if (!state.runtime) {
    badge.textContent = "检测中";
    badge.className = "badge neutral";
    status.textContent = "正在检测 Microsoft Word 和 LibreOffice。";
    wordBadge.textContent = "检测中";
    wordBadge.className = "badge neutral";
    libreOfficeBadge.textContent = "检测中";
    libreOfficeBadge.className = "badge neutral";
    progress.style.width = "0%";
    progressText.textContent = "检测中";
    progressBytes.textContent = "";
    actions.hidden = true;
    return;
  }
  officialLink.href = state.runtime.official_download_page || "https://www.libreoffice.org/download/download-libreoffice/";
  const job = state.runtime.download || {};
  const busy = ["queued", "downloading", "installing"].includes(job.status);
  const libreOfficeAvailable = Boolean(state.runtime.libreoffice?.available);
  const selected = selectedRuntime();
  [["word", wordBadge], ["libreoffice", libreOfficeBadge]].forEach(([method, methodBadge]) => {
    const available = Boolean(state.runtime[method]?.available);
    methodBadge.textContent = available ? "可用" : "未检测到";
    methodBadge.className = `badge ${available ? "success" : "neutral"}`;
  });
  $$('[data-runtime-method]').forEach((row) => {
    row.classList.toggle("is-selected", row.dataset.runtimeMethod === selected.selected);
  });
  const percent = Math.max(0, Math.min(100, Number(job.percent) || 0));
  progress.style.width = `${percent}%`;
  progressText.textContent = busy
    ? `${job.phase || "正在处理"} · ${percent}%`
    : libreOfficeAvailable
      ? "已检测到 LibreOffice"
      : job.status === "error"
        ? "上次下载失败，可重试"
        : "等待安装";
  progressBytes.textContent = job.total
    ? `${formatRuntimeBytes(job.downloaded)} / ${formatRuntimeBytes(job.total)}`
    : "";
  if (progressBar) progressBar.setAttribute("aria-valuenow", String(percent));
  if (selected.available && state.runtime.pdf_rasterizer?.available) {
    badge.textContent = "可用";
    badge.className = "badge success";
    status.textContent = selected.requested === "auto"
      ? `自动选择将使用 ${runtimeLabel(selected.selected)} 进行真实分页渲染，并用 ${state.runtime.pdf_rasterizer.name} 生成逐页图片。`
      : `已选择 ${runtimeLabel(selected.selected)}，并用 ${state.runtime.pdf_rasterizer.name} 生成逐页图片。`;
  } else if (selected.available) {
    badge.textContent = "缺少转图组件";
    badge.className = "badge warning";
    status.textContent = "已检测到文档渲染器，但缺少 PDF 逐页转图组件；视觉验收不可用。";
  } else {
    badge.textContent = "不可用";
    badge.className = "badge warning";
    status.textContent = selected.requested === "auto"
      ? "未检测到 Microsoft Word 或 LibreOffice；请先准备渲染环境。"
      : `已选择 ${runtimeLabel(selected.requested)}，但当前未检测到；不会自动改用其他渲染器。`;
  }
  actions.hidden = libreOfficeAvailable;
  if (busy && selected.requested !== "word") {
    badge.textContent = "处理中";
    badge.className = "badge warning";
    status.textContent = job.phase || "正在准备 LibreOffice 运行包。";
  } else if (job.status === "error" && !selected.available && selected.requested !== "word") {
    status.textContent = job.error || "LibreOffice 下载或安装失败。";
  }
  downloadButton.disabled = busy;
  downloadButton.textContent = job.status === "error" ? "交给 Agent 重试" : "交给 Agent 安装";
}

function stopRuntimePolling() {
  if (state.runtimePollTimer !== null) {
    window.clearInterval(state.runtimePollTimer);
    state.runtimePollTimer = null;
  }
}

async function refreshRuntimeStatus() {
  try {
    const response = await fetch("/api/runtime/status");
    if (!response.ok) throw new Error("渲染环境检测失败");
    state.runtime = await response.json();
    renderAll();
    const jobStatus = state.runtime.download?.status;
    if (!["queued", "downloading", "installing"].includes(jobStatus)) stopRuntimePolling();
  } catch (error) {
    $("#renderRuntimeBadge").textContent = "不可用";
    $("#renderRuntimeBadge").className = "badge warning";
    $("#renderRuntimeStatus").textContent = error.message;
  }
}

function startRuntimePolling() {
  stopRuntimePolling();
  state.runtimePollTimer = window.setInterval(() => { void refreshRuntimeStatus(); }, 1200);
  void refreshRuntimeStatus();
}

async function loadRuntimeStatus() {
  await refreshRuntimeStatus();
  const jobStatus = state.runtime?.download?.status;
  if (["queued", "downloading", "installing"].includes(jobStatus)) startRuntimePolling();
}

async function downloadRuntime() {
  const button = $("#aiDownloadRuntimeButton");
  button.disabled = true;
  setStatus("正在后台准备 LibreOffice 运行包");
  try {
    const response = await fetch("/api/runtime/download", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: "{}",
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "运行包下载启动失败");
    state.runtime = { ...(state.runtime || {}), download: result };
    renderAll();
    startRuntimePolling();
  } catch (error) {
    button.disabled = false;
    setStatus(error.message, true);
  }
}

function visualReviewLabel(status) {
  return {
    passed: "通过",
    failed: "未通过",
    skipped: "已跳过",
    pending: "等待验收",
  }[status] || "未知状态";
}

function visualReviewLogClass(visual) {
  return {
    passed: "log-success",
    failed: "log-error",
    skipped: "log-neutral",
    pending: "log-pending",
  }[visual?.status] || "log-neutral";
}

function visualReviewMessage(visual) {
  const reason = visual?.reason ? ` · ${visual.reason}` : "";
  return `视觉验收：${visualReviewLabel(visual?.status)}${reason}`;
}

function renderSessionVisualLog() {
  const handoff = state.sessionHandoff;
  const visual = state.aiSession?.visual_verification || handoff?.visual_verification;
  if (!visual) return;
  const entries = [];
  if (handoff) {
    const methodCount = handoff.application_methods?.length || 0;
    const citationStatus = handoff.citation_context?.status === "ready"
      ? `引用上下文已就绪（${handoff.citation_context.paragraph_count} 个正文锚点，${handoff.citation_context.reference_count} 条文献）`
      : `引用上下文：${handoff.citation_context?.reason || "不可用"}`;
    entries.push(`<p class="log-success">已交接 ${methodCount} 个实际 Word/OOXML 写入方法</p>`);
    entries.push(`<p class="log-success">${escapeHtml(citationStatus)}</p>`);
    entries.push(`<p class="log-success">执行模式：${escapeHtml(handoff.execution_mode === "official-template" ? "官方模板" : "参数化写入")}</p>`);
    const renderer = handoff.tools?.verify_visual?.renderer_selection;
    if (renderer) {
      const rendererText = renderer.requested_method === "auto"
        ? `自动选择（${renderer.selected_method ? `当前为 ${runtimeLabel(renderer.selected_method)}` : "当前无可用渲染器"}）`
        : `${runtimeLabel(renderer.requested_method)}（不允许降级）`;
      entries.push(`<p class="log-success">视觉渲染：${escapeHtml(rendererText)}</p>`);
    }
  }
  entries.push(`<p class="${visualReviewLogClass(visual)}">${escapeHtml(visualReviewMessage(visual))}</p>`);
  if (visual.status === "pending") {
    entries.push('<p class="log-pending">AI 完成写入后会重新分析结构、渲染每一页，并回传最终视觉验收结果。</p>');
  }
}

function stopVisualStatusPolling() {
  if (state.visualPollTimer !== null) {
    window.clearInterval(state.visualPollTimer);
    state.visualPollTimer = null;
  }
}

async function refreshVisualVerification() {
  if (!state.aiSession || !state.sessionSubmitted || !state.verification.visual_enabled) return;
  try {
    const response = await fetch("/api/session/status");
    if (!response.ok) return;
    const session = await response.json();
    if (session.id !== state.aiSession.id) return;
    const previous = JSON.stringify(state.aiSession.visual_verification || null);
    const next = session.visual_verification || null;
    state.aiSession.visual_verification = next;
    if (session.verification) state.verification = { ...state.verification, ...session.verification };
    if (previous === JSON.stringify(next)) return;
    renderAll();
    renderSessionVisualLog();
    if (next?.status && next.status !== "pending") {
      stopVisualStatusPolling();
      setStatus(`AI 已回传${visualReviewMessage(next)}`);
    }
  } catch {
    // The session can finish after the dashboard closes; retain the last known state.
  }
}

function startVisualStatusPolling() {
  stopVisualStatusPolling();
  if (!state.aiSession || !state.sessionSubmitted || !state.verification.visual_enabled) return;
  if (state.aiSession.visual_verification?.status && state.aiSession.visual_verification.status !== "pending") return;
  state.visualPollTimer = window.setInterval(() => { void refreshVisualVerification(); }, 2500);
  void refreshVisualVerification();
}

function readFileBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1]);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function adoptAnalysisResult(file, result, citationContext = null) {
  state.file = file;
  state.citationContext = citationContext || result.citation_context || null;
  state.citationTask = state.citationContext?.task || null;
  state.citationRequested = false;
  $("#fileName").textContent = file.name;
  const detectedWorkflowId = result.template_identity?.workflow_id;
  if (detectedWorkflowId && state.templateWorkflows.some((item) => item.id === detectedWorkflowId)) {
    state.selectedPresetId = detectedWorkflowId;
    $("#presetSelect").value = detectedWorkflowId;
  }
  state.analysis = result;
  if (isOfficialTemplateWorkflow()) {
    state.spec = mergeIntoSelected(result.inferred_spec || {});
    const workflow = selectedPreset();
    state.spec.authority = {
      level: "uploaded-publisher-template",
      sources: [...(workflow.authority?.sources || []), file.name],
      override_note: workflow.authority?.override_note || "使用上传的当前官方模板。",
    };
  } else {
    state.spec = adoptExtractedTemplateSpec(result.inferred_spec);
    state.spec.authority = {
      level: "uploaded-template",
      sources: [file.name],
      override_note: "所有设置均直接提取自当前导入模板；模板未设置的项目保持为空或关闭。",
    };
  }
  renderAll();
}

function stageFile(file) {
  if (state.sessionSubmitted) {
    setStatus("当前任务已提交给 AI，无法再修改模板文件", true);
    return;
  }
  if (!file || !/\.(docx|dotx)$/i.test(file.name)) {
    setStatus("请选择 DOCX 或 DOTX 格式文件", true);
    return;
  }
  if (file.size > 30 * 1024 * 1024) {
    setStatus("文件超过 30 MB 限制", true);
    return;
  }
  state.stagedFile = file;
  const shortName = file.name.length > 24 ? file.name.slice(0, 22) + "..." : file.name;
  $("#fileName").textContent = `已选择：${shortName} (${(file.size / 1024).toFixed(1)} KB)`;
  const btn = $("#analyzeFileButton");
  if (btn) {
    btn.disabled = false;
    btn.textContent = "🤖 开始 AI 格式分析";
  }
  setStatus(`已选定【${file.name}】，点击下方【开始 AI 格式分析】按钮开始深度分析。`);
}

async function waitForConversationalAnalysis(taskId) {
  const deadline = Date.now() + 60 * 60 * 1000;
  while (Date.now() < deadline) {
    await new Promise((resolve) => window.setTimeout(resolve, 1000));
    const response = await fetch(`/api/analysis/tasks/${encodeURIComponent(taskId)}`);
    const task = await response.json();
    if (!response.ok) throw new Error(task.error || "无法读取当前对话 AI 的分析任务");
    if (task.status === "completed") return task.result;
    if (task.status === "failed") throw new Error(task.error || "当前对话 AI 分析失败");
  }
  throw new Error("等待当前对话 AI 分析超时，请确认 Dashboard 由当前 AI 任务启动");
}

function renderCrossValidation(cvReport) {
  const card = $("#crossValidationCard");
  const list = $("#cvCorroborationsList");
  const title = $("#cvEngineTitle");
  const summary = $("#cvAiSummary");
  if (!card || !list || !cvReport || !cvReport.corroborations) return;
  card.hidden = false;
  if (title && cvReport.ai_engine) {
    title.textContent = `AI (${cvReport.ai_engine}) 与程序双轨印证报告`;
  }
  if (summary) {
    if (cvReport.ai_summary) {
      summary.textContent = `💡 AI 大模型深度分析说明：${cvReport.ai_summary}`;
      summary.hidden = false;
    } else {
      summary.hidden = true;
    }
  }
  list.innerHTML = cvReport.corroborations.map((item) => {
    let desc = "";
    if (item.field === "body.line_spacing") {
      const v = item.value?.kind === "multiple" ? `${item.value.value} 倍行距` : `${item.value?.value_pt || ""} 磅固定值`;
      desc = `正文行距：<strong>${v}</strong>`;
    } else if (item.field === "body.first_line_indent_mm") {
      desc = `首行缩进：<strong>${item.value} mm (2 字符)</strong>`;
    } else if (item.field === "headings") {
      desc = `标题层级：<strong>${item.levels} 级标题体系完整识别</strong>`;
    } else if (item.field === "captions") {
      desc = `图表标注：<strong>图在下、表在上 (五号居中)</strong>`;
    } else if (item.field === "references") {
      desc = `参考文献：<strong>${item.standard || "GB/T 7714"} (两端对齐/悬挂缩进)</strong>`;
    } else if (item.field === "lists") {
      desc = `列表编号：<strong>${item.style || "编号列表"}</strong>`;
    } else {
      desc = `${item.field}：${JSON.stringify(item.value || "")}`;
    }
    return `
      <div class="cv-item">
        <span class="cv-item-icon">✓</span>
        <div class="cv-item-content">
          <div>${desc}</div>
          <div class="cv-item-track">${escapeHtml(item.track || "双轨印证一致")}</div>
        </div>
      </div>
    `;
  }).join("");
}

async function handleFile(file) {
  if (state.sessionSubmitted) {
    setStatus("当前任务已提交给 AI，无法再修改模板文件", true);
    return;
  }
  if (!file || !/\.(docx|dotx)$/i.test(file.name)) {
    setStatus("请选择 DOCX 或 DOTX 格式文件", true);
    return;
  }
  if (file.size > 30 * 1024 * 1024) {
    setStatus("文件超过 30 MB 限制", true);
    return;
  }
  const btn = $("#analyzeFileButton");
  if (btn) {
    btn.disabled = true;
    btn.textContent = "⏳ AI 正在深度分析中...";
  }
  setStatus(`AI 正在深度分析文件【${file.name}】中的格式要求与样式定义...`);
  try {
    const fileBase64 = await readFileBase64(file);
    const response = await fetch("/api/analyze/docx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filename: file.name, data: fileBase64 }),
    });
    const queued = await response.json();
    if (!response.ok) throw new Error(queued.error || "文件分析失败");
    if (!queued.task_id) throw new Error("本地服务没有返回对话 AI 分析任务");
    setStatus(`已将【${file.name}】交给当前对话 AI，等待分析结果自动回填...`);
    const result = await waitForConversationalAnalysis(queued.task_id);

    if (result.cross_validation) {
      renderCrossValidation(result.cross_validation);
    }

    // If we are in an AI session, keep the target file as session input, and adopt the template spec
    if (state.aiSession) {
      $("#fileName").textContent = `已提取规范：${file.name}`;
      state.customTemplate = {
        name: file.name,
        result: result,
      };
      state.selectedPresetId = "custom-template";
      if (result.inferred_spec) {
        state.spec = adoptExtractedTemplateSpec(result.inferred_spec);
      }
      state.spec.name = `模板提取规范：${file.name}`;
      state.spec.mode = "parameterized";
      state.spec.template_required = false;
      state.spec.authority = {
        level: "uploaded-template",
        sources: [file.name, state.aiSession.input.name],
        override_note: `已通过 AI 双轨印证从【${file.name}】提取格式规范，并将应用于目标文档【${state.aiSession.input.name}】。`,
      };
      renderPresetSelect();
      renderAll();
      setStatus(`✓ AI 全面分析与双轨印证完成！已成功从【${file.name}】提取规则并填充至模板配置！`);
    } else {
      adoptAnalysisResult(file, result, result.citation_context);
      setStatus(isOfficialTemplateWorkflow()
        ? `官方模板分析完成：${result.document.paragraphs} 个段落，${result.document.sections.length} 个分节；请在该模板中完成排版`
        : `✓ AI 双轨印证完成：${result.document.paragraphs} 个段落，${result.document.sections.length} 个分节，已自动配置参数`);
    }
    if (btn) {
      btn.disabled = false;
      btn.textContent = "✓ AI 分析完成 (已填充模板)";
    }
  } catch (error) {
    if (btn) {
      btn.disabled = false;
      btn.textContent = "🤖 重新进行 AI 分析";
    }
    setStatus(error.message, true);
  }
}

async function loadSession() {
  const response = await fetch("/api/session");
  if (!response.ok) throw new Error("AI 会话载入失败");
  const session = await response.json();
  if (session.mode !== "ai-session") return;
  if (!session.input?.name || !session.analysis) {
    throw new Error("AI 会话缺少预处理后的源文件或分析数据");
  }
  state.aiSession = session;
  state.sessionSubmitted = session.status === "submitted";
  state.sessionHandoff = null;
  if (session.verification) state.verification = { ...state.verification, ...session.verification };

  state.file = { name: session.input.name, size: session.input.size || 0 };
  state.citationContext = session.citation_context || null;
  state.citationTask = state.citationContext?.task || null;
  state.analysis = session.analysis;
  if (!state.selectedPresetId && state.presets.length) {
    state.selectedPresetId = state.presets[0].id;
  }
  applyPreset(state.selectedPresetId || state.presets[0]?.id);
  const select = $("#presetSelect");
  if (select && state.selectedPresetId) select.value = state.selectedPresetId;
  
  // Reset dropZone placeholder so it is clear this slot is for uploading specification/template documents
  $("#fileName").textContent = "点击或拖拽上传【格式规范文件 / 官方模板】提取要求";
  const dropZoneSmall = $("#dropZone small");
  if (dropZoneSmall) dropZoneSmall.textContent = "支持上传《撰写规范》或学校模板 .docx / .dotx · 自动解析提取排版规则";
  $("#fileInput").disabled = state.sessionSubmitted;
  $("#dropZone").classList.remove("is-locked");
  $("#dropZone").removeAttribute("aria-disabled");

  state.citationRequested = Boolean(session.citation_request);
  const completedVisual = session.visual_verification?.status && session.visual_verification.status !== "pending";
  setStatus(state.sessionSubmitted
    ? (completedVisual
      ? `AI 已回传${visualReviewMessage(session.visual_verification)}`
      : `本次会话已交给 AI：${session.input.name}，等待 AI 写入并验收。`)
    : `待排版文件已载入：${session.input.name}。您可以在上方选择预设、输入文字要求或上传规范文件定制排版规则。`);
  if (state.sessionSubmitted) {
    if (session.visual_verification?.status && session.visual_verification.status !== "pending") renderSessionVisualLog();
    startVisualStatusPolling();
  }
}

async function analyzeText() {
  const text = $("#requirementsText").value.trim();
  if (!text) {
    setStatus("格式要求为空", true);
    return;
  }
  setStatus("正在识别格式要求");
  try {
    const response = await fetch("/api/analyze/text", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "识别失败");
    state.analysis = result;
    state.customHeadingLevels.clear();
    state.spec = mergeIntoSelected(result.spec_patch || {});
    if (isOfficialTemplateWorkflow()) {
      const workflow = selectedPreset();
      state.spec.authority = {
        level: "publisher-template-with-explicit-requirements",
        sources: [...(workflow.authority?.sources || []), "pasted requirements"],
        override_note: "文字要求只补充会务或出版方明确规则，仍必须使用当前官方模板。",
      };
    } else {
      state.spec.authority = {
        level: "explicit-requirements",
        sources: ["pasted requirements"],
        override_note: "仅应用有明确证据的字段；未识别条款保留在右侧。",
      };
    }
    renderAll();
    setStatus(`识别完成：${result.evidence.length} 个字段，${result.unresolved.length} 条待确认`);
  } catch (error) {
    setStatus(error.message, true);
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function exportSpec() {
  const blob = new Blob([JSON.stringify(state.spec, null, 2) + "\n"], { type: "application/json" });
  downloadBlob(blob, `${state.spec.id || "word-format"}-spec.json`);
  setStatus("已导出格式规范");
}

async function applyDocument() {
  if (!state.aiSession) {
    setStatus("请使用 --input 启动目标文档 Dashboard 会话后再提交格式设置", true);
    return;
  }
  await submitToAI();
}

async function submitToAI() {
  if (!state.aiSession || state.sessionSubmitted) return;
  setStatus("正在把确认的参数和写入方法交给 AI");
  $("#applyButton").disabled = true;
  try {
    const response = await fetch("/api/session/submit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        session_id: state.aiSession.id,
        spec: state.spec,
        clear_direct_font_formatting: $("#clearDirectInput").checked,
        verification: state.verification,
        citation_request: state.citationRequested,
      }),
    });
    const handoff = await response.json();
    if (!response.ok) throw new Error(handoff.error || "AI 交接失败");
    state.sessionSubmitted = true;
    state.sessionHandoff = handoff;
    state.aiSession.visual_verification = handoff.visual_verification || null;
    renderAll();
    renderSessionVisualLog();
    startVisualStatusPolling();
    setStatus("已交给 AI 继续修改，等待 AI 写入并验收 DOCX");
  } catch (error) {
    setStatus(error.message, true);
    renderAll();
  }
}

function activateSettingsPanel(panelName) {
  $$('[data-setting-panel]').forEach((panel) => {
    panel.hidden = panel.dataset.settingPanel !== panelName;
  });
  $$('[data-settings-tab]').forEach((tab) => {
    const active = tab.dataset.settingsTab === panelName;
    tab.classList.toggle("active", active);
    tab.setAttribute("aria-selected", String(active));
  });
}

function bindEvents() {
  $("#presetSelect").addEventListener("change", (event) => applyPreset(event.target.value));
  $$('[data-settings-tab]').forEach((tab) => {
    tab.addEventListener("click", () => activateSettingsPanel(tab.dataset.settingsTab));
  });
  activateSettingsPanel("source");
  $$("[data-spec-path]").forEach((input) => {
    input.addEventListener("input", () => {
      const value = input.type === "checkbox"
        ? input.checked
        : input.type === "number"
          ? (input.value === "" ? null : Number(input.value))
          : input.value;
      setPath(state.spec, input.dataset.specPath, value);
      if (input.dataset.specPath === "page.size" && PAGE_SIZE_MM[value]) {
        const dimensions = PAGE_SIZE_MM[value];
        Object.entries(dimensions).forEach(([key, dimension]) => {
          setPath(state.spec, `page.${key}`, dimension);
          const dimensionInput = $(`[data-spec-path="page.${key}"]`);
          if (dimensionInput) dimensionInput.value = dimension;
        });
      }
      if (input.dataset.specPath.startsWith("body.")) updateBodyRoleSummary();
      renderPreview();
    });
  });
  // Body font size select
  $("#bodyFontSizeSelect")?.addEventListener("change", (event) => {
    const val = event.target.value;
    if (val === "custom") {
      $("#bodyCustomFontSizeRow").hidden = false;
      $("#bodyFontSizeInput").focus();
    } else {
      $("#bodyCustomFontSizeRow").hidden = true;
      const numPt = Number(val) || 12;
      setPath(state.spec, "body.font_size_pt", numPt);
      $("#bodyFontSizeInput").value = numPt;
      const unit = $("#bodyFirstLineIndentUnitSelect")?.value || "chars";
      if (unit === "chars") {
        const chars = Number($("#bodyFirstLineIndentValueInput")?.value) || 2;
        const newMm = convertToMm(chars, "chars", numPt);
        setPath(state.spec, "body.first_line_indent_mm", newMm);
        $("#bodyIndentConversionHint").textContent = formatIndentSummary(newMm, numPt);
      }
      updateBodyRoleSummary();
      renderPreview();
    }
  });

  // Body custom font size input
  $("#bodyFontSizeInput")?.addEventListener("input", (event) => {
    const numPt = Number(event.target.value) || 12;
    setPath(state.spec, "body.font_size_pt", numPt);
    const unit = $("#bodyFirstLineIndentUnitSelect")?.value || "chars";
    if (unit === "chars") {
      const chars = Number($("#bodyFirstLineIndentValueInput")?.value) || 2;
      const newMm = convertToMm(chars, "chars", numPt);
      setPath(state.spec, "body.first_line_indent_mm", newMm);
      $("#bodyIndentConversionHint").textContent = formatIndentSummary(newMm, numPt);
    }
    updateBodyRoleSummary();
    renderPreview();
  });

  // Body first line indent unit select
  $("#bodyFirstLineIndentUnitSelect")?.addEventListener("change", (event) => {
    const unit = event.target.value;
    const pt = state.spec.body?.font_size_pt || 12;
    const mm = state.spec.body?.first_line_indent_mm ?? 8;
    const unitNames = { chars: "字符", pt: "磅 (pt)", mm: "毫米 (mm)", cm: "厘米 (cm)" };
    $("#bodyFirstLineIndentValueLabel").textContent = `首行缩进 (${unitNames[unit] || "字符"})`;
    $("#bodyFirstLineIndentValueInput").value = convertFromMm(mm, unit, pt);
    $("#bodyFirstLineIndentValueInput").step = "1";
  });

  // Body first line indent value input
  $("#bodyFirstLineIndentValueInput")?.addEventListener("input", (event) => {
    const val = Number(event.target.value) || 0;
    const unit = $("#bodyFirstLineIndentUnitSelect")?.value || "chars";
    const pt = state.spec.body?.font_size_pt || 12;
    const mm = convertToMm(val, unit, pt);
    setPath(state.spec, "body.first_line_indent_mm", mm);
    $("#bodyIndentConversionHint").textContent = formatIndentSummary(mm, pt);
    renderPreview();
  });

  // Body line spacing kind select
  $("#bodyLineSpacingKindSelect")?.addEventListener("change", (event) => {
    const kind = event.target.value;
    state.spec.body ??= {};
    if (kind === "exact") {
      state.spec.body.line_spacing = { kind: "exact", value_pt: 22 };
    } else if (kind === "at_least") {
      state.spec.body.line_spacing = { kind: "at_least", value_pt: 22 };
    } else if (kind === "single") {
      state.spec.body.line_spacing = { kind: "single" };
    } else if (kind === "1.5") {
      state.spec.body.line_spacing = { kind: "1.5" };
    } else if (kind === "double") {
      state.spec.body.line_spacing = { kind: "double" };
    } else {
      state.spec.body.line_spacing = { kind: "multiple", value: 1.5 };
    }
    updateLineSpacingUI("body", state.spec.body.line_spacing, false);
    renderPreview();
  });

  // Body line spacing value input
  $("#lineSpacingInput")?.addEventListener("input", (event) => {
    const kind = $("#bodyLineSpacingKindSelect")?.value || "multiple";
    state.spec.body ??= {};
    if (kind === "exact" || kind === "at_least") {
      state.spec.body.line_spacing = { kind, value_pt: Number(event.target.value) || 22 };
    } else {
      state.spec.body.line_spacing = { kind: "multiple", value: Number(event.target.value) || 1.5 };
    }
    renderPreview();
  });

  // Reference line spacing kind select
  $("#referenceLineSpacingKindSelect")?.addEventListener("change", (event) => {
    const kind = event.target.value;
    state.spec.references ??= {};
    if (kind === "exact") {
      state.spec.references.line_spacing = { kind: "exact", value_pt: 22 };
    } else if (kind === "at_least") {
      state.spec.references.line_spacing = { kind: "at_least", value_pt: 22 };
    } else if (kind === "single") {
      state.spec.references.line_spacing = { kind: "single" };
    } else if (kind === "1.5") {
      state.spec.references.line_spacing = { kind: "1.5" };
    } else if (kind === "double") {
      state.spec.references.line_spacing = { kind: "double" };
    } else {
      state.spec.references.line_spacing = { kind: "multiple", value: 1 };
    }
    updateLineSpacingUI("reference", state.spec.references.line_spacing, false);
    renderPreview();
  });

  // Reference line spacing value input
  $("#referenceSpacingInput")?.addEventListener("input", (event) => {
    const kind = $("#referenceLineSpacingKindSelect")?.value || "multiple";
    state.spec.references ??= {};
    if (kind === "exact" || kind === "at_least") {
      state.spec.references.line_spacing = { kind, value_pt: Number(event.target.value) || 22 };
    } else {
      state.spec.references.line_spacing = { kind: "multiple", value: Number(event.target.value) || 1 };
    }
    renderPreview();
  });

  // Heading font size select
  $("#headingFontSizeSelect")?.addEventListener("change", (event) => {
    const val = event.target.value;
    const heading = headingForLevel(state.selectedHeadingLevel, true);
    if (val === "custom") {
      $("#headingCustomFontSizeRow").hidden = false;
      $("#headingSizeInput").focus();
    } else {
      $("#headingCustomFontSizeRow").hidden = true;
      const numPt = Number(val) || 12;
      heading.font_size_pt = numPt;
      $("#headingSizeInput").value = numPt;
      state.customHeadingLevels.add(state.selectedHeadingLevel);
      updateHeadingEditor(false);
      renderPreview();
    }
  });

  // Caption figure font size select
  $("#captionFigureFontSizeSelect")?.addEventListener("change", (event) => {
    const val = event.target.value;
    state.spec.captions ??= {};
    state.spec.captions.figure ??= {};
    if (val === "custom") {
      $("#captionFigureCustomSizeRow").hidden = false;
      $("#captionFigureSizeInput").focus();
    } else {
      $("#captionFigureCustomSizeRow").hidden = true;
      const numPt = Number(val) || 10.5;
      state.spec.captions.figure.font_size_pt = numPt;
      $("#captionFigureSizeInput").value = numPt;
      renderPreview();
    }
  });

  // Caption table font size select
  $("#captionTableFontSizeSelect")?.addEventListener("change", (event) => {
    const val = event.target.value;
    state.spec.captions ??= {};
    state.spec.captions.table ??= {};
    if (val === "custom") {
      $("#captionTableCustomSizeRow").hidden = false;
      $("#captionTableSizeInput").focus();
    } else {
      $("#captionTableCustomSizeRow").hidden = true;
      const numPt = Number(val) || 10.5;
      state.spec.captions.table.font_size_pt = numPt;
      $("#captionTableSizeInput").value = numPt;
      renderPreview();
    }
  });

  // Reference font size select
  $("#referenceFontSizeSelect")?.addEventListener("change", (event) => {
    const val = event.target.value;
    state.spec.references ??= {};
    if (val === "custom") {
      $("#referenceCustomSizeRow").hidden = false;
      $("#referenceSizeInput").focus();
    } else {
      $("#referenceCustomSizeRow").hidden = true;
      const numPt = Number(val) || 10.5;
      state.spec.references.font_size_pt = numPt;
      $("#referenceSizeInput").value = numPt;
      const unit = $("#referenceHangingIndentUnitSelect")?.value || "chars";
      if (unit === "chars") {
        const chars = Number($("#referenceHangingIndentValueInput")?.value) || 2;
        const newMm = convertToMm(chars, "chars", numPt);
        setPath(state.spec, "references.hanging_indent_mm", newMm);
        $("#referenceIndentConversionHint").textContent = formatIndentSummary(newMm, numPt);
      }
      renderPreview();
    }
  });

  // Reference custom font size input
  $("#referenceSizeInput")?.addEventListener("input", (event) => {
    const numPt = Number(event.target.value) || 10.5;
    setPath(state.spec, "references.font_size_pt", numPt);
    const unit = $("#referenceHangingIndentUnitSelect")?.value || "chars";
    if (unit === "chars") {
      const chars = Number($("#referenceHangingIndentValueInput")?.value) || 2;
      const newMm = convertToMm(chars, "chars", numPt);
      setPath(state.spec, "references.hanging_indent_mm", newMm);
      $("#referenceIndentConversionHint").textContent = formatIndentSummary(newMm, numPt);
    }
    renderPreview();
  });

  // Reference hanging indent unit select
  $("#referenceHangingIndentUnitSelect")?.addEventListener("change", (event) => {
    const unit = event.target.value;
    const pt = state.spec.references?.font_size_pt || 10.5;
    const mm = state.spec.references?.hanging_indent_mm ?? 7;
    const unitNames = { chars: "字符", pt: "磅 (pt)", mm: "毫米 (mm)", cm: "厘米 (cm)" };
    $("#referenceHangingIndentValueLabel").textContent = `悬挂缩进 (${unitNames[unit] || "字符"})`;
    $("#referenceHangingIndentValueInput").value = convertFromMm(mm, unit, pt);
    $("#referenceHangingIndentValueInput").step = "1";
  });

  // Reference hanging indent value input
  $("#referenceHangingIndentValueInput")?.addEventListener("input", (event) => {
    const val = Number(event.target.value) || 0;
    const unit = $("#referenceHangingIndentUnitSelect")?.value || "chars";
    const pt = state.spec.references?.font_size_pt || 10.5;
    const mm = convertToMm(val, unit, pt);
    setPath(state.spec, "references.hanging_indent_mm", mm);
    $("#referenceIndentConversionHint").textContent = formatIndentSummary(mm, pt);
    renderPreview();
  });

  const bindHeadingField = (selector, property, eventName, convert = (input) => input.value) => {
    $(selector)?.addEventListener(eventName, (event) => {
      const heading = headingForLevel(state.selectedHeadingLevel, true);
      heading[property] = convert(event.target);
      state.customHeadingLevels.add(state.selectedHeadingLevel);
      updateHeadingEditor(false);
      renderPreview();
    });
  };
  bindHeadingField("#headingEastAsiaInput", "font_east_asia", "input");
  bindHeadingField("#headingLatinInput", "font_latin", "input");
  bindHeadingField("#headingSizeInput", "font_size_pt", "input", (input) => Number(input.value) || 12);
  bindHeadingField("#headingAlignInput", "alignment", "change");
  bindHeadingField("#headingSpaceBeforeInput", "space_before_pt", "input", (input) => Number(input.value) || 0);
  bindHeadingField("#headingSpaceAfterInput", "space_after_pt", "input", (input) => Number(input.value) || 0);
  bindHeadingField("#headingFirstLineIndentInput", "first_line_indent_chars", "input", (input) => input.value === "" ? 0 : Number(input.value));
  bindHeadingField("#headingLeftIndentInput", "left_indent_chars", "input", (input) => input.value === "" ? 0 : Number(input.value));
  bindHeadingField("#headingBoldInput", "bold", "change", (input) => input.checked);
  bindHeadingField("#headingItalicInput", "italic", "change", (input) => input.checked);
  $$("[data-heading-level]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedHeadingLevel = Number(button.dataset.headingLevel);
      updateHeadingEditor(isOfficialTemplateWorkflow() && !state.analysis?.source);
    });
  });
  $("#fileInput").addEventListener("change", (event) => stageFile(event.target.files[0]));
  $("#analyzeFileButton")?.addEventListener("click", () => {
    if (state.stagedFile) {
      handleFile(state.stagedFile);
    } else {
      setStatus("请先选择或拖拽上传一个规范/模板 DOCX 文件", true);
    }
  });
  $("#analyzeTextButton").addEventListener("click", analyzeText);
  $("#exportButton").addEventListener("click", exportSpec);
  $("#applyButton").addEventListener("click", applyDocument);
  $("#citationRequestButton").addEventListener("click", () => {
    if (state.citationContext?.status !== "ready" || state.sessionSubmitted) return;
    state.citationRequested = !state.citationRequested;
    if (!state.aiSession && state.citationRequested) setStatus("已开启在正文中引用；请使用 AI 交接流程提交，AI 将在正文中匹配并写入引用标注");
    renderAll();
  });
  $("#visualReviewToggle").addEventListener("change", (event) => {
    if (!visualReviewRuntimeReady() || state.sessionSubmitted) {
      event.target.checked = false;
      if (!state.sessionSubmitted) state.verification.visual_enabled = false;
      renderAll();
      return;
    }
    state.verification.visual_enabled = event.target.checked;
    renderAll();
  });
  $("#visualModelSelect")?.addEventListener("change", (event) => {
    if (state.sessionSubmitted) return;
    state.verification.visual_model = event.target.value;
    renderAll();
  });
  $("#customVisualModelInput")?.addEventListener("input", (event) => {
    if (state.sessionSubmitted) return;
    state.verification.custom_model_id = event.target.value.trim();
  });
  $("#renderMethodSelect").addEventListener("change", (event) => {
    if (state.sessionSubmitted) return;
    state.verification.render_method = event.target.value;
    renderAll();
  });
  $("#aiDownloadRuntimeButton").addEventListener("click", downloadRuntime);
  $$(".source-tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      $$(".source-tab").forEach((item) => {
        const active = item === tab;
        item.classList.toggle("active", active);
        item.setAttribute("aria-selected", String(active));
      });
      $$(".source-content").forEach((panel) => panel.classList.remove("active"));
      $(tab.dataset.sourceTab === "file" ? "#fileSource" : "#textSource").classList.add("active");
    });
  });
  const dropZone = $("#dropZone");
  ["dragenter", "dragover"].forEach((name) => dropZone.addEventListener(name, (event) => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  }));
  ["dragleave", "drop"].forEach((name) => dropZone.addEventListener(name, (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  }));
  dropZone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
    stageFile(event.dataTransfer.files[0]);
  });
  let resizeFrame = null;
  window.addEventListener("resize", () => {
    cancelAnimationFrame(resizeFrame);
    resizeFrame = requestAnimationFrame(renderPreview);
  });
}

bindEvents();
async function bootstrap() {
  await loadPresets();
  try {
    await loadRuntimeStatus();
  } catch (error) {
    $("#renderRuntimeStatus").textContent = error.message;
  }
  try {
    await loadCapabilities();
  } catch {
    // Retain defaults if capabilities endpoint fails
  }
  await loadSession();
}
bootstrap().catch((error) => setStatus(error.message, true));
