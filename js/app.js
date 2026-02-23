/* ══════════════════════════════════════════════════════════════
   XLSX 条件求和工具 · 纯客户端应用
   所有 Excel 处理由 ExcelJS 在浏览器完成，无需后端服务器
   ══════════════════════════════════════════════════════════════ */

// ── 常量 ──────────────────────────────────────────

const DEFAULT_THEME_COLORS = [
  '000000', 'FFFFFF', '44546A', 'E7E6E3',
  '4472C4', 'ED7D31', 'A5A5A5', 'FFC000',
  '5B9BD5', '70AD47', '0563C1', '954F72',
];

const INDEXED_COLORS = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

const MAX_PREVIEW_ROWS = 500;

// ── 状态 ──────────────────────────────────────────

const state = {
  fileBuffer: null,      // ArrayBuffer — 原始文件
  fileName: null,
  workbook: null,        // ExcelJS Workbook
  themeColors: null,     // 从主题解析出的色表
  sourcePreview: null,   // {sheets: [...]}
  outputBuffer: null,    // ArrayBuffer — 输出文件
  outputPreview: null,
  results: null,
  activeView: 'source',  // 'source' | 'output'
  activeSheetByView: {
    source: 0,
    output: 0,
    result: 0,
  },
  colors: [],            // [{hex, count, locations}]
  formulas: [],          // [{name, count}]
  rules: [],             // [{field, cmp, value}]
  manualExcludes: new Set(), // e.g. Sheet1!G12
  pickExcludeMode: false,
  excludeLogic: 'OR',
  modalMode: 'create',   // 'create' | 'edit'
  editingRuleIndex: -1,
  modalRuleDraft: {
    field: 'fill_rgb',
    cmp: '==',
    value: '',
  },
  previewZoom: 1,
  ruleHitCountCache: new Map(),
};

function getActiveSheetIndex(view) {
  return state.activeSheetByView?.[view] ?? 0;
}

function setActiveSheetIndex(view, index) {
  if (!state.activeSheetByView) state.activeSheetByView = { source: 0, output: 0, result: 0 };
  state.activeSheetByView[view] = Math.max(0, Number(index) || 0);
}

// ── DOM 引用 ──────────────────────────────────────

const $ = (s) => document.querySelector(s);
const $$ = (s) => document.querySelectorAll(s);

const fileInput = $('#file-input');
const emptyState = $('#empty-state');
const previewState = $('#preview-state');
const dropZone = $('#drop-zone');
const loadingEl = $('#loading');
const loadingText = $('#loading-text');
const tableLoadingEl = $('#table-loading');
const sidebarUploadBtn = $('#btn-sidebar-upload');
const paneToolsBtn = $('#btn-pane-tools');
const panePreviewBtn = $('#btn-pane-preview');
const mobilePaneTabs = $('#mobile-pane-tabs');
const zoomOutBtn = $('#btn-zoom-out');
const zoomInBtn = $('#btn-zoom-in');
const zoomResetBtn = $('#btn-zoom-reset');
const zoomValueEl = $('#zoom-value');
const toolbarDownloadBtn = $('#btn-download');
const mobileDownloadBtn = $('#btn-download-mobile');
const mobileDownloadPanel = $('#panel-download-mobile');

// ── 初始化 ────────────────────────────────────────

document.addEventListener('DOMContentLoaded', init);

function init() {
  setEmptyStateVisible(true);
  hideLoading();

  // 文件选择
  fileInput.addEventListener('change', () => {
    if (fileInput.files.length) handleFile(fileInput.files[0]);
  });

  // 仅保留左侧上传按钮（已存在文件时先确认是否覆盖）
  if (sidebarUploadBtn) {
    sidebarUploadBtn.addEventListener('click', onSidebarUploadClick);
  }

  // 拖拽
  const mainEl = $('#main');
  for (const evt of ['dragenter', 'dragover']) {
    mainEl.addEventListener(evt, (e) => {
      e.preventDefault();
      dropZone && dropZone.classList.add('drag-over');
    });
  }
  mainEl.addEventListener('dragleave', (e) => {
    if (!mainEl.contains(e.relatedTarget)) {
      dropZone && dropZone.classList.remove('drag-over');
    }
  });
  mainEl.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone && dropZone.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  });

  // View tabs
  for (const tab of $$('.view-tab')) {
    tab.addEventListener('click', () => {
      if (tab.dataset.disabled === 'true') {
        toast(tab.title || '需要先进行运算', true);
        return;
      }
      const view = tab.dataset.view;
      if (view === 'output' && !state.outputPreview) return;
      if (view === 'result' && !state.results) return;
      state.activeView = view;
      refreshViewTabs();
      renderPreview();
      syncDownloadButton();
    });
  }

  // Rules
  $('#btn-add-rule').addEventListener('click', () => {
    openAddRuleModal();
  });

  const colInput = $('#col-input');
  if (colInput) {
    colInput.addEventListener('input', () => {
      if (!previewState.hidden) renderPreview();
    });
  }

  const logicToggle = $('#logic-toggle');
  if (logicToggle) {
    logicToggle.querySelectorAll('.logic-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const logic = btn.dataset.logic === 'AND' ? 'AND' : 'OR';
        state.excludeLogic = logic;
        logicToggle.querySelectorAll('.logic-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        if (!previewState.hidden) renderPreview();
      });
    });
  }

  const btnToggleExcludePick = $('#btn-toggle-exclude-pick');
  if (btnToggleExcludePick) {
    btnToggleExcludePick.addEventListener('click', () => {
      state.pickExcludeMode = !state.pickExcludeMode;
      btnToggleExcludePick.classList.toggle('is-active', state.pickExcludeMode);
      btnToggleExcludePick.textContent = state.pickExcludeMode ? '关闭点选剔除' : '开启点选剔除';
      if (!previewState.hidden) renderPreview();
    });
  }

  const btnClearExcludes = $('#btn-clear-excludes');
  if (btnClearExcludes) {
    btnClearExcludes.addEventListener('click', () => {
      state.manualExcludes.clear();
      renderExcludeList();
      if (!previewState.hidden) renderPreview();
      toast('已清空手动剔除列表');
    });
  }

  // Execute
  $('#btn-run').addEventListener('click', handleRun);

  // Download
  if (toolbarDownloadBtn) toolbarDownloadBtn.addEventListener('click', handleDownload);
  if (mobileDownloadBtn) mobileDownloadBtn.addEventListener('click', handleDownload);

  initPreviewZoomControls();
  initMobilePaneTabs();

  // Modals
  initModals();
}

function isMobileLayout() {
  return window.matchMedia('(max-width: 860px)').matches;
}

function setMobilePreviewPane(preview) {
  document.body.classList.toggle('mobile-pane-preview', !!preview);
  if (paneToolsBtn && panePreviewBtn) {
    paneToolsBtn.classList.toggle('active', !preview);
    panePreviewBtn.classList.toggle('active', !!preview);
  }
  syncDownloadButton();
}

function initMobilePaneTabs() {
  if (paneToolsBtn) {
    paneToolsBtn.addEventListener('click', () => setMobilePreviewPane(false));
  }
  if (panePreviewBtn) {
    panePreviewBtn.addEventListener('click', () => setMobilePreviewPane(true));
  }

  if (mobilePaneTabs) mobilePaneTabs.hidden = !isMobileLayout();
  setMobilePreviewPane(false);

  window.addEventListener('resize', () => {
    if (mobilePaneTabs) mobilePaneTabs.hidden = !isMobileLayout();
    if (!isMobileLayout()) {
      setMobilePreviewPane(false);
    }
    syncDownloadButton();
  });
}

function clampZoom(v) {
  return Math.max(0.5, Math.min(2, v));
}

function applyPreviewZoom() {
  const table = $('#preview-table');
  if (!table) return;

  // 移动端使用系统手势缩放，不做程序缩放
  if (isMobileLayout()) {
    state.previewZoom = 1;
    table.style.zoom = '';
    table.style.transform = '';
    table.style.transformOrigin = '';
    table.style.width = '';
    if (zoomValueEl) zoomValueEl.textContent = '100%';
    return;
  }

  const z = clampZoom(state.previewZoom || 1);
  state.previewZoom = z;
  if (CSS?.supports?.('zoom', '1')) {
    table.style.zoom = String(z);
    table.style.transform = '';
    table.style.transformOrigin = '';
    table.style.width = '';
  } else {
    table.style.zoom = '';
    table.style.transform = `scale(${z})`;
    table.style.transformOrigin = 'top left';
    table.style.width = `${100 / z}%`;
  }
  if (zoomValueEl) zoomValueEl.textContent = `${Math.round(z * 100)}%`;
}

function initPreviewZoomControls() {
  if (zoomOutBtn) {
    zoomOutBtn.addEventListener('click', () => {
      if (isMobileLayout()) return;
      state.previewZoom = clampZoom((state.previewZoom || 1) - 0.1);
      applyPreviewZoom();
    });
  }
  if (zoomInBtn) {
    zoomInBtn.addEventListener('click', () => {
      if (isMobileLayout()) return;
      state.previewZoom = clampZoom((state.previewZoom || 1) + 0.1);
      applyPreviewZoom();
    });
  }
  if (zoomResetBtn) {
    zoomResetBtn.addEventListener('click', () => {
      if (isMobileLayout()) return;
      state.previewZoom = 1;
      applyPreviewZoom();
    });
  }
  applyPreviewZoom();
}

// ── 文件处理 ──────────────────────────────────────

async function handleFile(file) {
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    toast('仅支持 .xlsx 格式文件', true);
    return;
  }

  showLoading('正在读取文件…');

  try {
    // 新文件覆盖旧状态（单文件模式）
    state.outputBuffer = null;
    state.outputPreview = null;
    state.results = null;
    state.activeView = 'source';
    state.activeSheetByView = { source: 0, output: 0, result: 0 };
    state.manualExcludes = new Set();
    state.pickExcludeMode = false;
    state.excludeLogic = 'OR';
    state.rules = [];
    state.ruleHitCountCache = new Map();

    const buffer = await file.arrayBuffer();
    state.fileBuffer = buffer;
    state.fileName = file.name;

    // 读 workbook
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    state.workbook = wb;
    state.themeColors = parseThemeColors(wb);

    // 更新侧栏文件名
    const nameEl = $('#file-name');
    nameEl.textContent = file.name;
    nameEl.classList.add('loaded');
    if (sidebarUploadBtn) {
      sidebarUploadBtn.innerHTML = `
        <img class="btn-icon-img" src="icon/upload.svg" alt="" aria-hidden="true">
        重新上传文件
      `;
    }

    // 提取预览
    state.sourcePreview = extractPreview(wb, state.themeColors);
    state.activeView = 'source';
    setActiveSheetIndex('source', 0);

    // 检测颜色（用于条件下拉选择）
    state.colors = detectColors(wb, state.themeColors);
    state.formulas = detectFormulaFunctions(wb);

    // 显示侧栏面板
    showSidebarPanels();

    // 添加默认规则
    resetRules();
    const logicToggle = $('#logic-toggle');
    if (logicToggle) {
      logicToggle.querySelectorAll('.logic-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.logic === 'OR');
      });
    }
    const colInput = $('#col-input');
    if (colInput) {
      colInput.value = '';
    }

    // 隐藏输出 tab / 结果
    setResultTabsEnabled(false);
    syncDownloadButton();
    renderExcludeList();
    const btnToggleExcludePick = $('#btn-toggle-exclude-pick');
    if (btnToggleExcludePick) {
      btnToggleExcludePick.classList.remove('is-active');
      btnToggleExcludePick.textContent = '开启点选剔除';
    }

    // 切换到预览
    setEmptyStateVisible(false);
    renderPreview();

  } catch (err) {
    console.error(err);
    toast('文件读取失败: ' + err.message, true);
  } finally {
    hideLoading();
    if (dropZone) {
      dropZone.classList.remove('drag-over');
    }
    fileInput.value = '';
  }
}

function onSidebarUploadClick() {
  // 首次上传：直接选择
  if (!state.fileBuffer) {
    fileInput.click();
    return;
  }
  openReuploadModal();
}

// ── 主题色解析 ────────────────────────────────────

function parseThemeColors(wb) {
  // ExcelJS 内部存储主题 XML
  const themeXml = wb._themes?.theme1;
  if (!themeXml) return [...DEFAULT_THEME_COLORS];

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(themeXml, 'application/xml');
    const ns = 'http://schemas.openxmlformats.org/drawingml/2006/main';

    const clrScheme = doc.getElementsByTagNameNS(ns, 'clrScheme')[0];
    if (!clrScheme) return [...DEFAULT_THEME_COLORS];

    const names = [
      'dk1', 'lt1', 'dk2', 'lt2',
      'accent1', 'accent2', 'accent3', 'accent4',
      'accent5', 'accent6', 'hlink', 'folHlink',
    ];
    const colors = [];
    for (let i = 0; i < names.length; i++) {
      const elem = clrScheme.getElementsByTagNameNS(ns, names[i])[0];
      if (!elem) { colors.push(DEFAULT_THEME_COLORS[i]); continue; }
      const srgb = elem.getElementsByTagNameNS(ns, 'srgbClr')[0];
      const sys = elem.getElementsByTagNameNS(ns, 'sysClr')[0];
      if (srgb) colors.push((srgb.getAttribute('val') || '000000').toUpperCase());
      else if (sys) colors.push((sys.getAttribute('lastClr') || '000000').toUpperCase());
      else colors.push(DEFAULT_THEME_COLORS[i]);
    }
    return colors;
  } catch (e) {
    console.warn('Theme parse failed:', e);
    return [...DEFAULT_THEME_COLORS];
  }
}

// ── 颜色工具 ──────────────────────────────────────

function applyTint(hexColor, tint) {
  let r = parseInt(hexColor.substr(0, 2), 16) / 255;
  let g = parseInt(hexColor.substr(2, 2), 16) / 255;
  let b = parseInt(hexColor.substr(4, 2), 16) / 255;

  // RGB → HLS
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h, l = (max + min) / 2, s;
  if (max === min) { h = 0; s = 0; }
  else {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }

  // Apply tint
  if (tint < 0) l = l * (1 + tint);
  else l = l * (1 - tint) + tint;
  l = Math.max(0, Math.min(1, l));

  // HLS → RGB
  let r2, g2, b2;
  if (s === 0) { r2 = g2 = b2 = l; }
  else {
    const hue2rgb = (p, q, t) => {
      if (t < 0) t += 1; if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r2 = hue2rgb(p, q, h + 1 / 3);
    g2 = hue2rgb(p, q, h);
    b2 = hue2rgb(p, q, h - 1 / 3);
  }

  const toHex = (v) => Math.round(v * 255).toString(16).padStart(2, '0').toUpperCase();
  return toHex(r2) + toHex(g2) + toHex(b2);
}

function getCellFillHex(cell, themeColors) {
  const fill = cell.fill || cell.style?.fill;
  if (!fill) return null;
  if (fill.type !== 'pattern' || fill.pattern === 'none') return null;

  const fg = fill.fgColor;
  if (!fg) return null;

  // ARGB 直取
  if (fg.argb) {
    const argb = String(fg.argb).toUpperCase();
    if (argb === '00000000') return null; // 透明
    const hex6 = argb.length === 8 ? argb.substring(2) : argb;
    if (hex6 === '000000' && argb.startsWith('00')) return null;
    return hex6;
  }

  // 主题色
  if (fg.theme != null) {
    const idx = fg.theme;
    // theme=0(dk1) 在部分文件里会作为“自动色”落到填充，实际并非用户设置的黑底
    // 这里优先过滤这类常见误判，避免预览出现整块黑底异常。
    if (idx === 0 && !fg.tint) return null;
    const base = (idx >= 0 && idx < themeColors.length)
      ? themeColors[idx] : null;
    if (!base) return null;
    const tint = fg.tint || 0;
    return tint !== 0 ? applyTint(base, tint) : base.toUpperCase();
  }

  // 索引色
  if (fg.indexed != null && fg.indexed >= 0 && fg.indexed < INDEXED_COLORS.length) {
    return INDEXED_COLORS[fg.indexed];
  }

  return null;
}

function isDarkColor(hex) {
  if (!hex) return false;
  const r = parseInt(hex.substr(0, 2), 16);
  const g = parseInt(hex.substr(2, 2), 16);
  const b = parseInt(hex.substr(4, 2), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.55;
}

// ── 颜色检测 ──────────────────────────────────────

function detectColors(wb, themeColors) {
  const map = {};
  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const hex = getCellFillHex(cell, themeColors);
        if (!hex || hex === 'FFFFFF' || hex === '000000') return;
        if (!map[hex]) map[hex] = { count: 0, locations: [] };
        map[hex].count++;
        if (map[hex].locations.length < 4) {
          map[hex].locations.push(`${ws.name} ${cell.address}`);
        }
      });
    });
  }
  return Object.keys(map).sort().map(hex => ({
    hex,
    count: map[hex].count,
    locations: map[hex].locations,
  }));
}

function detectFormulaFunctions(wb) {
  const map = {};
  const fnPattern = /([A-Z][A-Z0-9._]*)\s*\(/gi;
  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        // cell.formula works for regular, shared-master, AND shared-slave cells
        const formula = (typeof cell.formula === 'string' ? cell.formula : '').trim();
        if (!formula) return;
        // Extract ALL function names in the formula (including nested)
        const cleaned = formula.replace(/^=/, '');
        fnPattern.lastIndex = 0;
        let m;
        while ((m = fnPattern.exec(cleaned)) !== null) {
          const fn = m[1].toUpperCase();
          // Skip cell references that look like functions (e.g., A1, BC12)
          if (/^[A-Z]{1,3}\d+$/.test(fn)) continue;
          if (!map[fn]) map[fn] = { count: 0 };
          map[fn].count += 1;
        }
      });
    });
  }

  return Object.entries(map)
    .map(([name, info]) => ({ name, count: info.count }))
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
}

// ── 预览提取 ──────────────────────────────────────

function extractPreview(wb, themeColors, maxRows = MAX_PREVIEW_ROWS) {
  const sheets = [];
  for (const ws of wb.worksheets) {
    const totalRows = ws.rowCount || 0;
    const totalCols = ws.columnCount || 1;
    const limit = Math.min(totalRows, maxRows);

    // headers: 列字母 + 第一行内容
    const headers = [];
    for (let c = 1; c <= totalCols; c++) {
      const letter = colLetter(c);
      const cell = ws.getCell(1, c);
      const val = getCellDisplayValue(cell);
      headers.push({ col: letter, label: val !== '' ? String(val) : letter });
    }

    const rows = [];
    const colors = [];
    for (let r = 1; r <= limit; r++) {
      const rowVals = [];
      const rowColors = [];
      for (let c = 1; c <= totalCols; c++) {
        const cell = ws.getCell(r, c);
        rowVals.push(getCellDisplayValue(cell));
        const hex = getCellFillHex(cell, themeColors);
        rowColors.push(hex ? `#${hex}` : null);
      }
      rows.push(rowVals);
      colors.push(rowColors);
    }

    sheets.push({ name: ws.name, headers, rows, colors, totalRows, totalCols });
  }
  return { sheets };
}

function getCellDisplayValue(cell) {
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    // Formula
    if (v.formula != null) {
      return v.result != null ? v.result : `=${v.formula}`;
    }
    // RichText
    if (v.richText) return v.richText.map(t => t.text).join('');
    // Date
    if (v instanceof Date) return v.toLocaleDateString('zh-CN');
    // SharedFormula
    if (v.sharedFormula) return v.result != null ? v.result : '';
    return String(v);
  }
  return v;
}

function colLetter(n) {
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

function colIndex(letter) {
  let n = 0;
  for (const ch of letter.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}

// ── 表达式引擎 ────────────────────────────────────

function buildCellContext(cell, themeColors) {
  const v = cell.value;
  const formulaText = resolveCellFormulaText(cell, v);
  const isFormula = !!formulaText
    || (v != null && typeof v === 'object' && (v.formula != null || v.sharedFormula != null))
    || cell?.model?.sharedFormula != null;

  let numericValue = null;
  if (isFormula) {
    // Result may be in cell.value.result or cell.model.result
    const result = (v && typeof v === 'object' ? v.result : undefined)
      ?? cell?.model?.result;
    if (typeof result === 'number') numericValue = result;
  } else {
    if (typeof v === 'number') numericValue = v;
    else if (typeof v === 'string') {
      const parsed = Number(v.replace(/,/g, ''));
      if (!Number.isNaN(parsed) && v.trim() !== '') numericValue = parsed;
    }
  }

  const fillHex = getCellFillHex(cell, themeColors);

  return {
    value: isFormula ? v.result : v,
    numeric_value: numericValue,
    formula_text: formulaText,
    fill_rgb: fillHex ? `#${fillHex}` : null,
    cell_ref: '',
  };
}

function evaluateExpr(expr, ctx) {
  if (!expr) return false;
  const op = expr.op;
  if (op === 'AND') return (expr.children || []).every(c => evaluateExpr(c, ctx));
  if (op === 'OR') return (expr.children || []).some(c => evaluateExpr(c, ctx));
  if (op === 'NOT') return !evaluateExpr(expr.child, ctx);
  if (expr.type === 'condition') return evalCondition(expr, ctx);
  return false;
}

function evalCondition(cond, ctx) {
  const { field, cmp, value: target } = cond;
  const actual = ctx[field];

  // 颜色
  if (field === 'fill_rgb') {
    const norm = (s) => s ? s.toUpperCase().replace('#', '') : '';
    if (actual == null) return cmp === '!=' && target != null;
    if (cmp === '==') return norm(actual) === norm(target);
    if (cmp === '!=') return norm(actual) !== norm(target);
    return false;
  }

  // 公式文本
  if (field === 'formula_text') {
    const a = normalizeFormulaSelector(actual);
    const t = normalizeFormulaSelector(target);
    if (!t) return false;
    if (cmp === '==') return a === t;
    if (cmp === '!=') return a !== t;
    return false;
  }

  // 单元格位置匹配（sheet+坐标）
  if (field === 'cell_ref') {
    const arr = Array.isArray(target) ? target : [];
    const set = new Set(arr.map(v => String(v).toUpperCase()));
    const ref = String(actual || '').toUpperCase();
    if (cmp === 'in') return set.has(ref);
    if (cmp === 'not_in') return !set.has(ref);
    return false;
  }

  // 数值
  if (field === 'value') {
    const num = ctx.numeric_value;
    if (num == null || target == null) return false;
    const t = parseFloat(target);
    if (isNaN(t)) return false;
    if (cmp === '>') return num > t;
    if (cmp === '>=') return num >= t;
    if (cmp === '<') return num < t;
    if (cmp === '<=') return num <= t;
    if (cmp === '==') return num === t;
    if (cmp === '!=') return num !== t;
  }

  return false;
}

// ── 核心求和 ──────────────────────────────────────

function processSheet(ws, colLetters, excludeExpr, themeColors) {
  const manualExcludeExpr = buildManualExcludeExpr();
  const results = {};
  for (const letter of colLetters) {
    const ci = colIndex(letter);
    let total = 0, excluded = 0, included = 0;
    const totalRows = ws.rowCount || 1;

    for (let r = 2; r <= totalRows; r++) {
      const cell = ws.getCell(r, ci);
      if (cell.value == null) continue;

      const ctx = buildCellContext(cell, themeColors);
      ctx.cell_ref = `${ws.name}!${cell.address}`;

      if (manualExcludeExpr && evaluateExpr(manualExcludeExpr, ctx)) {
        excluded++;
        continue;
      }

      if (excludeExpr && evaluateExpr(excludeExpr, ctx)) {
        excluded++;
        continue;
      }

      if (ctx.numeric_value != null) {
        total += ctx.numeric_value;
        included++;
      }
    }

    results[letter] = { sum: total, excluded, included };
  }
  return results;
}

// ── 执行求和 ──────────────────────────────────────

async function handleRun() {
  if (!state.fileBuffer) {
    toast('请先上传文件', true);
    return;
  }

  const colsRaw = $('#col-input').value.trim().toUpperCase();
  if (!colsRaw) { toast('请输入求和列', true); return; }

  const parsed = parseColumnsInput(colsRaw);
  if (!parsed.ok) {
    toast(`无效的列输入: ${parsed.invalid.join(', ')}`, true);
    return;
  }
  const colLetters = parsed.columns;

  const excludeExpr = buildExcludeExpr();

  showLoading('正在计算…');
  const btn = $('#btn-run');
  btn.disabled = true;

  try {
    // 我们在新 workbook 上操作以生成输出文件
    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(state.fileBuffer);
    const outTheme = parseThemeColors(outWb);

    const allResults = {};

    for (const ws of outWb.worksheets) {
      const results = processSheet(ws, colLetters, excludeExpr, outTheme);
      allResults[ws.name] = results;

      // 写入求和行
      const sumRowNum = (ws.rowCount || 1) + 1;
      const sumRow = ws.getRow(sumRowNum);
      ws.getCell(sumRowNum, 1).value = '条件合计';
      ws.getCell(sumRowNum, 1).font = { bold: true };

      for (const letter of colLetters) {
        const ci = colIndex(letter);
        const cell = ws.getCell(sumRowNum, ci);
        cell.value = results[letter].sum;
        cell.font = { bold: true };

        // 复制数字格式
        for (let r = 2; r < sumRowNum; r++) {
          const ref = ws.getCell(r, ci);
          if (ref.numFmt && ref.numFmt !== 'General') {
            cell.numFmt = ref.numFmt;
            break;
          }
        }
      }
    }

    // 保存到 buffer
    const outBuf = await outWb.xlsx.writeBuffer();
    state.outputBuffer = outBuf;
    state.results = allResults;

    // 提取输出预览
    const outWb2 = new ExcelJS.Workbook();
    await outWb2.xlsx.load(outBuf);
    state.outputPreview = extractPreview(outWb2, parseThemeColors(outWb2));

    // 启用右侧结果相关 tab
    setResultTabsEnabled(true);
    syncDownloadButton();

    // 默认切到“计算结果”tab
    state.activeView = 'result';
    setActiveSheetIndex('result', 0);
    refreshViewTabs();
    renderPreview();
    syncDownloadButton();

    toast('计算完成');

  } catch (err) {
    console.error(err);
    toast('计算失败: ' + err.message, true);
  } finally {
    hideLoading();
    btn.disabled = false;
  }
}

// ── 下载 ──────────────────────────────────────────

function handleDownload() {
  if (!state.outputBuffer) { toast('请先执行求和', true); return; }

  const blob = new Blob([state.outputBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const baseName = state.fileName.replace(/\.xlsx$/i, '');
  a.href = url;
  a.download = `${baseName}_条件求和.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ── 渲染: 预览 ───────────────────────────────────

function renderPreview() {
  if (state.activeView === 'result') {
    renderResultsInMain();
    return;
  }

  const data = state.activeView === 'source' ? state.sourcePreview : state.outputPreview;
  if (!data || !data.sheets.length) return;

  const activeSheetIndex = Math.min(getActiveSheetIndex(state.activeView), Math.max(0, data.sheets.length - 1));
  const sheet = data.sheets[activeSheetIndex] || data.sheets[0];
  const selectedCols = getSelectedColumns();
  const ruleExpr = buildExcludeExpr();
  const manualExcludeExpr = buildManualExcludeExpr();
  const sourceWs = state.activeView === 'source' && state.workbook ? state.workbook.getWorksheet(sheet.name) : null;

  // Sheet tabs
  const tabsEl = $('#sheet-tabs');
  tabsEl.hidden = false;
  tabsEl.innerHTML = '';
  data.sheets.forEach((s, i) => {
    const btn = document.createElement('button');
    btn.className = 'sheet-tab' + (i === activeSheetIndex ? ' active' : '');
    btn.textContent = s.name;
    btn.onclick = () => { setActiveSheetIndex(state.activeView, i); renderPreview(); };
    tabsEl.appendChild(btn);
  });

  // Row info
  const shown = sheet.rows.length;
  const total = sheet.totalRows;
  $('#row-info').textContent = shown < total
    ? `显示 ${shown} / ${total} 行`
    : `共 ${total} 行`;

  // Table
  const table = $('#preview-table');
  let html = '<thead><tr><th class="col-rownum">#</th>';
  sheet.headers.forEach(h => {
    html += `<th title="列 ${h.col}">${escHtml(h.label)}<div style="font-weight:400;color:#94a3b8;font-size:11px">${h.col}</div></th>`;
  });
  html += '</tr></thead><tbody>';

  sheet.rows.forEach((row, ri) => {
    // 检查是否是条件合计行
    const isSumRow = (typeof row[0] === 'string' && row[0].includes('条件合计'));
    html += `<tr${isSumRow ? ' class="sum-row"' : ''}>`;
    html += `<td class="col-rownum">${ri + 1}</td>`;

    row.forEach((val, ci) => {
      const bg = (sheet.colors[ri] && sheet.colors[ri][ci]) || '';
      const dark = bg ? isDarkColor(bg.replace('#', '')) : false;
      let style = '';
      if (bg) style += `background:${bg};`;
      if (dark) style += 'color:#fff;';

      const excelCol = colLetter(ci + 1);
      const excelRow = ri + 1;
      const cellAddress = `${excelCol}${excelRow}`;
      const cellKey = `${sheet.name}!${cellAddress}`;
      const isSelectedCol = selectedCols.has(excelCol);
      let isManualExcluded = false;
      let isIncludedCell = false;

      if (sourceWs && isSelectedCol && excelRow >= 2) {
        const sourceCell = sourceWs.getCell(excelRow, ci + 1);
        if (sourceCell.value != null) {
          const ctx = buildCellContext(sourceCell, state.themeColors || DEFAULT_THEME_COLORS);
          ctx.cell_ref = `${sheet.name}!${cellAddress}`;
          isManualExcluded = manualExcludeExpr ? evaluateExpr(manualExcludeExpr, ctx) : false;

          if (!isManualExcluded && ctx.numeric_value != null) {
            if (!ruleExpr) {
              isIncludedCell = true;
            } else {
              const matched = evaluateExpr(ruleExpr, ctx);
              isIncludedCell = !matched;
            }
          }
        }
      } else {
        isManualExcluded = state.manualExcludes.has(cellKey);
      }

      const isNum = typeof val === 'number';
      let cls = isNum ? '' : ' cell-text';
      if (isManualExcluded) cls += ' manual-excluded';
      if (isIncludedCell) cls += ' calc-highlight';
      if (state.pickExcludeMode && state.activeView === 'source') cls += ' pickable-cell';
      const display = isNum ? fmtNum(val) : escHtml(String(val));
      html += `<td class="${cls.trim()}" style="${style}" data-cell-address="${cellAddress}" data-sheet-name="${escAttr(sheet.name)}" title="${escAttr(sheet.name)}!${cellAddress}">${display}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody>';
  table.innerHTML = html;
  applyPreviewZoom();

  bindPreviewCellInteractions();
}

function refreshViewTabs() {
  $$('.view-tab').forEach(t => {
    t.classList.toggle('active', t.dataset.view === state.activeView);
  });
}

function syncDownloadButton() {
  const hasOutput = !!state.outputBuffer;
  const downloadableView = state.activeView === 'output' || state.activeView === 'result';
  const visible = hasOutput && downloadableView;

  // 桌面端：保留 toolbar 下载按钮；移动端：隐藏 toolbar 下载按钮
  if (toolbarDownloadBtn) {
    const toolbarVisible = visible && !isMobileLayout();
    updateDownloadActionButton(toolbarDownloadBtn, toolbarVisible, state.activeView);
  }

  // 移动端：在底部显示绿色下载按钮
  if (mobileDownloadBtn) {
    const mobileVisible = visible && isMobileLayout() && document.body.classList.contains('mobile-pane-preview');
    updateDownloadActionButton(mobileDownloadBtn, mobileVisible, state.activeView);
    if (mobileDownloadPanel) mobileDownloadPanel.hidden = !mobileVisible;
    document.body.classList.toggle('mobile-download-visible', mobileVisible);
  } else {
    if (mobileDownloadPanel) mobileDownloadPanel.hidden = true;
    document.body.classList.remove('mobile-download-visible');
  }
}

function updateDownloadActionButton(btn, visible, activeView) {
  if (!btn) return;
  btn.hidden = !visible;
  const isResult = activeView === 'result';
  btn.title = isResult ? '下载计算结果文件' : '下载最终文件';
  btn.innerHTML = `
    <img class="btn-icon-img" src="icon/download.svg" alt="" aria-hidden="true">
    ${isResult ? '下载计算结果文件' : '下载最终文件'}
  `;
}

// ── 渲染: 颜色 ───────────────────────────────────

function renderResultsInMain() {
  const tabsEl = $('#sheet-tabs');
  tabsEl.hidden = true;
  $('#row-info').textContent = '计算结果总览';

  const table = $('#preview-table');
  if (!state.results) {
    table.innerHTML = '';
    return;
  }

  let html = '<thead><tr><th class="col-rownum">#</th><th>Sheet</th><th>列</th><th>合计</th><th>纳入</th><th>排除</th></tr></thead><tbody>';
  let idx = 1;
  for (const [sheet, cols] of Object.entries(state.results)) {
    for (const [col, r] of Object.entries(cols)) {
      html += `<tr>
        <td class="col-rownum">${idx++}</td>
        <td class="cell-text">${escHtml(sheet)}</td>
        <td class="cell-text">${col}</td>
        <td>${fmtNum(r.sum)}</td>
        <td>${r.included}</td>
        <td>${r.excluded}</td>
      </tr>`;
    }
  }
  html += '</tbody>';
  table.innerHTML = html;
  applyPreviewZoom();
}

function bindPreviewCellInteractions() {
  if (state.activeView !== 'source') return;
  const table = $('#preview-table');
  if (!table) return;
  table.querySelectorAll('td[data-cell-address]').forEach((td) => {
    td.addEventListener('click', () => {
      if (!state.pickExcludeMode) return;
      const sheet = td.dataset.sheetName;
      const address = td.dataset.cellAddress;
      if (!sheet || !address) return;
      toggleManualExclude(sheet, address);
      renderExcludeList();
      renderPreview();
    });
  });
}

function toggleManualExclude(sheetName, cellAddress) {
  const key = `${sheetName}!${cellAddress}`;
  if (state.manualExcludes.has(key)) {
    state.manualExcludes.delete(key);
  } else {
    state.manualExcludes.add(key);
  }
}

function renderExcludeList() {
  const listEl = $('#exclude-list');
  if (!listEl) return;
  if (!state.manualExcludes.size) {
    listEl.innerHTML = '<span class="hint">当前没有手动剔除项</span>';
    return;
  }
  listEl.innerHTML = '';
  Array.from(state.manualExcludes).sort().forEach((key) => {
    const tag = document.createElement('span');
    tag.className = 'exclude-tag';
    tag.innerHTML = `${escHtml(key)} <button type="button" class="btn-icon" title="移除" data-key="${escAttr(key)}">×</button>`;
    listEl.appendChild(tag);
  });
  listEl.querySelectorAll('button[data-key]').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      const key = e.currentTarget.dataset.key;
      if (!key) return;
      state.manualExcludes.delete(key);
      renderExcludeList();
      renderPreview();
    });
  });
}

// ── 排除条件 UI ──────────────────────────────────

const FIELD_OPTIONS = [
  { value: 'fill_rgb', label: '背景颜色' },
  { value: 'value', label: '数值' },
  { value: 'formula_text', label: '公式' },
];

const CMP_OPTIONS_BY_FIELD = {
  fill_rgb: [
    { value: '==', label: '等于' },
    { value: '!=', label: '不等于' },
  ],
  value: [
    { value: '==', label: '等于' },
    { value: '!=', label: '不等于' },
  ],
  formula_text: [
    { value: '==', label: '等于' },
    { value: '!=', label: '不等于' },
  ],
};

function resetRules() {
  state.rules = [];
  state.editingRuleIndex = -1;
  renderRulesEmptyState();
}

function renderRulesEmptyState() {
  const list = $('#rules-list');
  if (!list) return;
  list.innerHTML = '';

  if (!state.rules.length) {
    const div = document.createElement('div');
    div.className = 'rules-empty';
    div.textContent = '还没有条件：点击“添加条件”，在弹窗中配置规则。';
    list.appendChild(div);
    return;
  }

  state.rules.forEach((rule, idx) => {
    const hitCount = getRuleHitCount(rule);
    const row = document.createElement('div');
    row.className = 'rule-row';
    row.innerHTML = `
      <div class="rule-content">
        <div class="rule-title">条件 ${idx + 1}</div>
        <div class="rule-summary">${formatRuleSummaryHtml(rule)}</div>
        <div class="rule-hit">当前条件选中了<span class="hit-number">${hitCount}</span>个单元格</div>
      </div>
      <div class="rule-actions">
        <button type="button" class="btn-icon rule-edit-btn" data-edit-index="${idx}" title="编辑条件">
          <img class="btn-icon-img" src="icon/setting.svg" alt="" aria-hidden="true">
        </button>
        <button type="button" class="btn-icon rule-delete-btn" data-delete-index="${idx}" title="删除条件">
          <img class="btn-icon-img" src="icon/times.svg" alt="" aria-hidden="true">
        </button>
      </div>
    `;
    list.appendChild(row);
  });

  list.querySelectorAll('button[data-edit-index]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const index = Number(btn.dataset.editIndex);
      if (Number.isNaN(index)) return;
      const rule = state.rules[index];
      const hitCount = getRuleHitCount(rule);
      toast(`当前条件选中了 ${hitCount} 个单元格`);
      openAddRuleModal('edit', index);
    });
  });

  list.querySelectorAll('button[data-delete-index]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const index = Number(btn.dataset.deleteIndex);
      if (Number.isNaN(index)) return;
      const rule = state.rules[index];
      const hitCount = getRuleHitCount(rule);
      toast(`删除前提示：当前条件选中了 ${hitCount} 个单元格`);
      state.rules.splice(index, 1);
      renderRulesEmptyState();
      if (!previewState.hidden) renderPreview();
    });
  });
}

function getRuleHitCount(rule) {
  if (!state.workbook || !rule) return 0;
  const normalized = normalizeRule(rule);
  if (!normalized) return 0;
  const key = JSON.stringify(normalized);
  if (state.ruleHitCountCache?.has(key)) return state.ruleHitCountCache.get(key);

  let count = 0;
  const expr = { type: 'condition', ...normalized };
  for (const ws of state.workbook.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const ctx = buildCellContext(cell, state.themeColors || DEFAULT_THEME_COLORS);
        ctx.cell_ref = `${ws.name}!${cell.address}`;
        if (evaluateExpr(expr, ctx)) count += 1;
      });
    });
  }

  if (!state.ruleHitCountCache) state.ruleHitCountCache = new Map();
  state.ruleHitCountCache.set(key, count);
  return count;
}

function initModals() {
  const reuploadModal = $('#modal-reupload');
  const addRuleModal = $('#modal-add-rule');

  const closeIfOverlay = (e) => {
    if (e.target === reuploadModal) closeModal(reuploadModal);
    if (e.target === addRuleModal) closeModal(addRuleModal);
  };
  if (reuploadModal) reuploadModal.addEventListener('click', closeIfOverlay);
  if (addRuleModal) addRuleModal.addEventListener('click', closeIfOverlay);

  const btnReuploadCancel = $('#btn-reupload-cancel');
  const btnReuploadConfirm = $('#btn-reupload-confirm');
  if (btnReuploadCancel) {
    btnReuploadCancel.addEventListener('click', () => closeModal(reuploadModal));
  }
  if (btnReuploadConfirm) {
    btnReuploadConfirm.addEventListener('click', () => {
      closeModal(reuploadModal);
      fileInput.click();
    });
  }

  const btnAddRuleCancel = $('#btn-add-rule-cancel');
  const btnAddRuleConfirm = $('#btn-add-rule-confirm');
  if (btnAddRuleCancel) {
    btnAddRuleCancel.addEventListener('click', () => closeModal(addRuleModal));
  }
  if (btnAddRuleConfirm) {
    btnAddRuleConfirm.addEventListener('click', () => {
      const nextRule = normalizeRule(state.modalRuleDraft);
      if (!nextRule) {
        toast('请填写一条有效条件', true);
        return;
      }

      if (state.modalMode === 'edit' && state.editingRuleIndex >= 0 && state.editingRuleIndex < state.rules.length) {
        state.rules.splice(state.editingRuleIndex, 1, nextRule);
      } else {
        state.rules.push(nextRule);
      }

      renderRulesEmptyState();
      if (!previewState.hidden) renderPreview();
      closeModal(addRuleModal);
    });
  }

  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    if (addRuleModal && !addRuleModal.hidden) closeModal(addRuleModal);
    else if (reuploadModal && !reuploadModal.hidden) closeModal(reuploadModal);
  });
}

function openModal(el) {
  if (!el) return;
  el.hidden = false;
}

function closeModal(el) {
  if (!el) return;
  el.hidden = true;
}

function openReuploadModal() {
  openModal($('#modal-reupload'));
}

function openAddRuleModal(mode = 'create', index = -1) {
  state.modalMode = mode;
  state.editingRuleIndex = index;
  state.modalRuleDraft = createDraftRule(mode, index);

  const titleEl = $('#add-rule-title');
  const btnEl = $('#btn-add-rule-confirm');
  if (titleEl) titleEl.textContent = mode === 'edit' ? '编辑剔除条件' : '添加剔除条件';
  if (btnEl) btnEl.textContent = mode === 'edit' ? '保存条件' : '添加条件';

  renderModalRuleBuilder();
  refreshModalRuleHitPreview();
  openModal($('#modal-add-rule'));
}

function renderModalRuleBuilder() {
  const root = $('#modal-rule-builder');
  if (!root) return;

  const draft = state.modalRuleDraft || createDraftRule('create', -1);
  const field = draft.field || 'fill_rgb';
  const cmp = draft.cmp || ((CMP_OPTIONS_BY_FIELD[field] || [])[0]?.value || '==');
  const value = draft.value || '';

  root.innerHTML = `
    <div class="modal-section">
      <div class="modal-row-title">条件类型</div>
      <div class="modal-rule-line">
        <div class="rule-field-group" data-value="${escAttr(field)}">${renderFieldOptions(field)}</div>
      </div>
    </div>
    <div class="modal-section">
      <div class="modal-row-title">比较关系</div>
      <div class="modal-rule-line">
        <div class="rule-cmp-group" data-value="${escAttr(cmp)}">${renderCmpOptions(field, cmp)}</div>
      </div>
    </div>
    <div class="modal-section">
      <div class="modal-row-title">匹配值</div>
      <div class="modal-rule-line">
        <span class="rule-value-wrap">${renderRuleValueControl(field, value)}</span>
      </div>
    </div>
  `;

  bindModalRuleInteractions(root);
  refreshModalRuleHitPreview();
}

function bindModalRuleInteractions(root) {
  const fieldGroup = root.querySelector('.rule-field-group');
  const cmpGroup = root.querySelector('.rule-cmp-group');
  const valueWrap = root.querySelector('.rule-value-wrap');
  if (!fieldGroup || !cmpGroup || !valueWrap) return;

  const bindCmpButtons = () => {
    const cmpButtons = cmpGroup.querySelectorAll('.opt-btn');
    if (!cmpButtons.length) return;
    cmpButtons.forEach((btn) => {
      btn.addEventListener('click', () => {
        cmpButtons.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const next = btn.dataset.value || '==';
        cmpGroup.dataset.value = next;
        if (!state.modalRuleDraft) state.modalRuleDraft = createDraftRule('create', -1);
        state.modalRuleDraft.cmp = next;
        refreshModalRuleHitPreview();
      });
    });
  };

  fieldGroup.querySelectorAll('.opt-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      fieldGroup.querySelectorAll('.opt-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const field = btn.dataset.value || 'value';
      fieldGroup.dataset.value = field;

      const cmpCandidates = CMP_OPTIONS_BY_FIELD[field] || CMP_OPTIONS_BY_FIELD.value;
      const nextCmp = cmpCandidates[0]?.value || '==';
      cmpGroup.dataset.value = nextCmp;
      cmpGroup.innerHTML = renderCmpOptions(field, nextCmp);

      let defaultValue = '';
      if (field === 'fill_rgb' && state.colors && state.colors.length) {
        defaultValue = `#${state.colors[0].hex}`;
      } else if (field === 'formula_text' && state.formulas && state.formulas.length) {
        defaultValue = state.formulas[0].name;
      }
      valueWrap.innerHTML = renderRuleValueControl(field, defaultValue);

      if (!state.modalRuleDraft) state.modalRuleDraft = createDraftRule('create', -1);
      state.modalRuleDraft.field = field;
      state.modalRuleDraft.cmp = nextCmp;
      state.modalRuleDraft.value = defaultValue;

      bindCmpButtons();
      bindModalValueInteractions(root);
      refreshModalRuleHitPreview();
    });
  });

  bindCmpButtons();
  bindModalValueInteractions(root);
}

function bindModalValueInteractions(root) {
  const picker = root.querySelector('.color-swatch-picker');
  if (picker) {
    const hiddenInput = picker.querySelector('.rule-value');
    const customBtn = picker.querySelector('.swatch-custom-btn');
    const customInput = picker.querySelector('.swatch-custom-input');
    picker.querySelectorAll('.swatch-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const color = btn.dataset.color || '';
        if (hiddenInput) hiddenInput.value = color;
        picker.querySelectorAll('.swatch-btn').forEach(b => b.classList.remove('active'));
        if (customBtn) customBtn.classList.remove('active');
        btn.classList.add('active');
        if (!state.modalRuleDraft) state.modalRuleDraft = createDraftRule('create', -1);
        state.modalRuleDraft.value = color;
        if (customInput && /^#[0-9A-F]{6}$/i.test(color)) customInput.value = color;
        refreshModalRuleHitPreview();
      });
    });
    if (customBtn && customInput) {
      customBtn.addEventListener('click', () => {
        customInput.click();
      });
      customInput.addEventListener('input', () => {
        const color = String(customInput.value || '').toUpperCase();
        if (hiddenInput) hiddenInput.value = color;
        picker.querySelectorAll('.swatch-btn').forEach(b => b.classList.remove('active'));
        customBtn.classList.add('active');
        // Show the selected custom color on the button
        customBtn.style.background = color;
        customBtn.style.color = isDarkColor(color.replace('#', '')) ? '#fff' : '#1d4ed8';
        customBtn.style.borderColor = color;
        if (!state.modalRuleDraft) state.modalRuleDraft = createDraftRule('create', -1);
        state.modalRuleDraft.value = color;
        refreshModalRuleHitPreview();
      });
    }
  }

  const formulaPicker = root.querySelector('.formula-chip-picker');
  if (formulaPicker) {
    const textInput = root.querySelector('.rule-value-text');
    formulaPicker.querySelectorAll('.formula-chip').forEach((btn) => {
      btn.addEventListener('click', () => {
        const fn = btn.dataset.formula || '';
        formulaPicker.querySelectorAll('.formula-chip').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        if (textInput) {
          textInput.value = fn;
          textInput.classList.remove('formula-input-active');
        }
        if (!state.modalRuleDraft) state.modalRuleDraft = createDraftRule('create', -1);
        state.modalRuleDraft.value = fn;
        refreshModalRuleHitPreview();
      });
    });
  }

  const textInput = root.querySelector('.rule-value-text, .rule-value:not([type="hidden"])');
  if (textInput) {
    textInput.addEventListener('input', () => {
      if (!state.modalRuleDraft) state.modalRuleDraft = createDraftRule('create', -1);
      state.modalRuleDraft.value = textInput.value;
      if ((state.modalRuleDraft?.field || '') === 'value') {
        textInput.classList.toggle('formula-input-active', String(textInput.value || '').trim().length > 0);
      }
      const chips = root.querySelectorAll('.formula-chip');
      if (chips.length) {
        const v = String(textInput.value || '').trim().toUpperCase();
        let anyChipActive = false;
        chips.forEach((chip) => {
          const match = (chip.dataset.formula || '').toUpperCase() === v;
          chip.classList.toggle('active', match);
          if (match) anyChipActive = true;
        });
        // If text doesn't match any chip, highlight input as custom
        textInput.classList.toggle('formula-input-active', v && !anyChipActive);
      }
      refreshModalRuleHitPreview();
    });
  }
}

function refreshModalRuleHitPreview() {
  const el = $('#modal-rule-hit-preview');
  if (!el) return;
  if (!state.workbook) {
    el.innerHTML = '当前条件选中了<span class="hit-number">0</span>个单元格';
    return;
  }
  const normalized = normalizeRule(state.modalRuleDraft);
  if (!normalized) {
    el.innerHTML = '当前条件选中了<span class="hit-number">0</span>个单元格';
    return;
  }
  const hitCount = getRuleHitCount(normalized);
  el.innerHTML = `当前条件选中了<span class="hit-number">${hitCount}</span>个单元格`;
}

function buildExcludeExpr() {
  const conditions = (state.rules || [])
    .map(normalizeRule)
    .filter(Boolean)
    .map(rule => ({ type: 'condition', ...rule }));

  if (!conditions.length) return null;
  if (conditions.length === 1) return conditions[0];
  return { op: state.excludeLogic || 'OR', children: conditions };
}

function normalizeRule(rule) {
  if (!rule) return null;
  const field = rule.field;
  const cmp = rule.cmp;
  let value = rule.value;

  if (!FIELD_OPTIONS.some(opt => opt.value === field)) return null;
  const validCmp = (CMP_OPTIONS_BY_FIELD[field] || []).some(opt => opt.value === cmp);
  if (!validCmp) return null;

  if (field === 'value') {
    const n = parseFloat(String(value ?? '').trim());
    if (Number.isNaN(n)) return null;
    value = n;
  } else {
    value = String(value ?? '').trim();
    if (!value) return null;
    if (field === 'fill_rgb' && !value.startsWith('#')) value = `#${value}`;
    if (field === 'formula_text') value = normalizeFormulaSelector(value);
  }

  return { field, cmp, value };
}

function createDraftRule(mode, index) {
  const preferredColor = state.colors && state.colors.length ? `#${state.colors[0].hex}` : '#3B82F6';
  if (mode === 'edit' && index >= 0 && index < state.rules.length) {
    const r = state.rules[index];
    return {
      field: r.field || 'fill_rgb',
      cmp: r.cmp || ((CMP_OPTIONS_BY_FIELD[r.field] || [])[0]?.value || '=='),
      value: r.value == null ? '' : String(r.value),
    };
  }
  return {
    field: 'fill_rgb',
    cmp: '==',
    value: preferredColor,
  };
}

function getFieldLabel(field) {
  return FIELD_OPTIONS.find(opt => opt.value === field)?.label || field;
}

function getCmpLabel(field, cmp) {
  return (CMP_OPTIONS_BY_FIELD[field] || []).find(opt => opt.value === cmp)?.label || cmp;
}

function formatRuleSummary(rule) {
  const valueText = rule.field === 'fill_rgb' ? String(rule.value).toUpperCase() : String(rule.value);
  return `当 ${getFieldLabel(rule.field)} ${getCmpLabel(rule.field, rule.cmp)} ${valueText} 时剔除`;
}

function formatRuleSummaryHtml(rule) {
  if (rule.field === 'fill_rgb') {
    const color = String(rule.value || '').toUpperCase();
    const safeColor = /^#?[0-9A-F]{6}$/.test(color) ? (color.startsWith('#') ? color : `#${color}`) : null;
    const cmpLabel = escHtml(getCmpLabel(rule.field, rule.cmp));
    if (safeColor) {
      return `当 背景颜色 ${cmpLabel} <span class="inline-color-chip"><span class="inline-swatch" style="background:${escAttr(safeColor)}"></span><span class="inline-color-label">${escHtml(safeColor)}</span></span> 时剔除`;
    }
  }
  return escHtml(formatRuleSummary(rule));
}

// ── 侧栏面板显示 ─────────────────────────────────

function showSidebarPanels() {
  for (const id of ['panel-columns', 'panel-rules', 'panel-execute']) {
    $(`#${id}`).hidden = false;
  }
}

function getSelectedColumns() {
  const raw = ($('#col-input') && $('#col-input').value ? $('#col-input').value : '').toUpperCase();
  const parsed = parseColumnsInput(raw);
  return new Set(parsed.ok ? parsed.columns : []);
}

function parseColumnsInput(raw) {
  const tokens = String(raw || '').toUpperCase().split(/[,，\s]+/).filter(Boolean);
  const columns = [];
  const invalid = [];

  const pushUnique = (letter) => {
    if (!columns.includes(letter)) columns.push(letter);
  };

  for (const tokenRaw of tokens) {
    const token = tokenRaw.replace(/列/g, '');
    const single = token.match(/^([A-Z]{1,3})$/);
    if (single) {
      pushUnique(single[1]);
      continue;
    }

    const range = token.match(/^([A-Z]{1,3})\s*:\s*([A-Z]{1,3})$/);
    if (range) {
      const start = colIndex(range[1]);
      const end = colIndex(range[2]);
      if (!start || !end) {
        invalid.push(tokenRaw);
        continue;
      }
      const min = Math.min(start, end);
      const max = Math.max(start, end);
      for (let i = min; i <= max; i++) {
        pushUnique(colLetter(i));
      }
      continue;
    }

    invalid.push(tokenRaw);
  }

  if (!tokens.length) {
    return { ok: false, columns: [], invalid: ['(空)'] };
  }

  return {
    ok: invalid.length === 0 && columns.length > 0,
    columns,
    invalid,
  };
}

function buildManualExcludeExpr() {
  if (!state.manualExcludes || state.manualExcludes.size === 0) return null;
  return {
    type: 'condition',
    field: 'cell_ref',
    cmp: 'in',
    value: Array.from(state.manualExcludes),
  };
}

function renderRuleValueControl(field, value) {
  if (field === 'fill_rgb') {
    const colors = (state.colors || []).map(c => c.hex.toUpperCase());
    const current = value ? String(value).toUpperCase() : '';
    const currentForPicker = /^#[0-9A-F]{6}$/.test(current) ? current : '#3B82F6';
    const swatches = colors.map(hex => {
      const val = `#${hex}`;
      const active = current === val ? ' active' : '';
      return `<button type="button" class="swatch-btn${active}" data-color="${val}" title="${val}" style="background:${val}"></button>`;
    }).join('');
    const isCustom = current && !colors.some(hex => `#${hex}` === current);
    const customActive = isCustom ? ' active' : '';
    // If editing a custom color, show it on the button
    const customStyle = isCustom ? ` style="background:${escAttr(current)};color:${isDarkColor(current.replace('#', '')) ? '#fff' : '#1d4ed8'};border-color:${escAttr(current)}"` : '';
    return `<div class="color-swatch-picker"><input type="hidden" class="rule-value" value="${escAttr(current)}">${swatches}<button type="button" class="swatch-custom-btn${customActive}" title="自定义颜色"${customStyle}>自定义</button><input type="color" class="swatch-custom-input" value="${escAttr(currentForPicker)}"></div>`;
  }

  if (field === 'formula_text') {
    const formulas = (state.formulas || []).map(f => f.name.toUpperCase());
    const current = String(value || '').trim().toUpperCase();
    const matchesChip = formulas.some(name => name === current);
    const chips = formulas.slice(0, 12).map((name) => {
      const active = name === current ? ' active' : '';
      return `<button type="button" class="formula-chip${active}" data-formula="${escAttr(name)}" title="${escAttr(name)}">${escHtml(name)}</button>`;
    }).join('');
    // If value is custom (doesn't match any chip), add active class to input
    const inputActive = current && !matchesChip ? ' formula-input-active' : '';
    return `
      <div>
        ${chips ? `<div class="formula-chip-picker">${chips}</div>` : ''}
        <input class="rule-value rule-value-text${inputActive}" type="text" value="${escAttr(String(value || ''))}" placeholder="例如 SUM（支持自定义）">
      </div>
    `;
  }

  const placeholder = '值';
  const hasValue = String(value ?? '').trim().length > 0;
  const inputActive = hasValue ? ' formula-input-active' : '';
  return `<input class="rule-value${inputActive}" type="text" value="${escAttr(String(value || ''))}" placeholder="${placeholder}">`;
}

function renderFieldOptions(activeField) {
  return FIELD_OPTIONS.map((opt) => {
    const active = opt.value === activeField ? ' active' : '';
    return `<button type="button" class="opt-btn${active}" data-value="${opt.value}">${opt.label}</button>`;
  }).join('');
}

function renderCmpOptions(field, activeCmp) {
  const options = CMP_OPTIONS_BY_FIELD[field] || CMP_OPTIONS_BY_FIELD.value;
  const fallback = options[0]?.value || '==';
  const current = options.some(o => o.value === activeCmp) ? activeCmp : fallback;
  return options.map((opt) => {
    const active = opt.value === current ? ' active' : '';
    return `<button type="button" class="opt-btn${active}" data-value="${opt.value}">${opt.label}</button>`;
  }).join('');
}

function resolveCellFormulaText(cell, cellValue) {
  // 1. cell.formula getter — works for regular, shared-master AND shared-slave cells
  //    (ExcelJS automatically resolves the adjusted formula for shared slaves)
  const direct = typeof cell?.formula === 'string' ? cell.formula : '';
  if (direct) return direct;

  // 2. Fallback: model formula
  const modelFormula = typeof cell?.model?.formula === 'string' ? cell.model.formula : '';
  if (modelFormula) return modelFormula;

  // 3. Fallback: cell.value object
  if (cellValue && typeof cellValue === 'object') {
    const valueFormula = typeof cellValue.formula === 'string' ? cellValue.formula : '';
    if (valueFormula) return valueFormula;

    // Shared formula reference (string → master cell address)
    const sharedRef = typeof cellValue.sharedFormula === 'string' ? cellValue.sharedFormula : '';
    if (sharedRef && cell?.worksheet) {
      const master = cell.worksheet.getCell(sharedRef);
      const mf = typeof master?.formula === 'string' ? master.formula : '';
      if (mf) return mf;
    }
  }

  // 4. Fallback: model sharedFormula (may be string or object {master, formula})
  const modelSf = cell?.model?.sharedFormula;
  if (modelSf && cell?.worksheet) {
    if (typeof modelSf === 'object' && typeof modelSf.formula === 'string' && modelSf.formula) {
      return modelSf.formula;
    }
    const ref = typeof modelSf === 'string' ? modelSf : (modelSf?.master || '');
    if (ref) {
      const master = cell.worksheet.getCell(ref);
      const mf = typeof master?.formula === 'string' ? master.formula : '';
      if (mf) return mf;
    }
  }

  return '';
}

function normalizeFormulaSelector(input) {
  const raw = String(input ?? '').trim().toUpperCase();
  if (!raw) return '';
  const noEq = raw.startsWith('=') ? raw.slice(1).trim() : raw;
  const fnMatch = noEq.match(/^([A-Z][A-Z0-9._]*)\s*\(/);
  if (fnMatch) return fnMatch[1];
  const token = noEq.match(/^([A-Z][A-Z0-9._]*)$/);
  if (token) return token[1];
  return noEq;
}

function setResultTabsEnabled(enabled) {
  const resultTab = $('#tab-result');
  const outputTab = $('#tab-output');
  if (resultTab) {
    if (enabled) {
      resultTab.dataset.disabled = 'false';
      resultTab.classList.remove('is-disabled');
      resultTab.title = '计算结果汇总视图';
    } else {
      resultTab.dataset.disabled = 'true';
      resultTab.classList.add('is-disabled');
      resultTab.title = '需要先进行运算（统计汇总）';
    }
  }
  if (outputTab) {
    if (enabled) {
      outputTab.dataset.disabled = 'false';
      outputTab.classList.remove('is-disabled');
      outputTab.title = '输出结果文件预览（支持下载）';
    } else {
      outputTab.dataset.disabled = 'true';
      outputTab.classList.add('is-disabled');
      outputTab.title = '需要先进行运算（可下载文件）';
    }
  }
}

// ── 工具函数 ──────────────────────────────────────

function fmtNum(n) {
  if (Number.isInteger(n)) return n.toLocaleString('zh-CN');
  return n.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 4 });
}

function escHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function escAttr(s) {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function showLoading(text = '处理中…') {
  loadingText.textContent = text;
  loadingEl.hidden = false;
  loadingEl.style.display = 'flex';
  if (tableLoadingEl) {
    tableLoadingEl.hidden = false;
  }
}

function hideLoading() {
  loadingEl.hidden = true;
  loadingEl.style.display = 'none';
  if (tableLoadingEl) {
    tableLoadingEl.hidden = true;
  }
}

function setEmptyStateVisible(visible) {
  if (visible) {
    emptyState.hidden = false;
    previewState.hidden = true;
  } else {
    emptyState.hidden = true;
    previewState.hidden = false;
    if (isMobileLayout()) {
      setMobilePreviewPane(false);
    }
  }
}

// Toast 通知
let toastTimer = null;
function toast(msg, isError = false) {
  let el = document.querySelector('.toast');
  if (!el) {
    el = document.createElement('div');
    el.className = 'toast';
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.classList.toggle('error', isError);
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 3000);
}
