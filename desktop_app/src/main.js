/**
 * 石家庄心理医院 - 财务凭证与进销存自动化工作台
 * 核心交互逻辑 (Tauri 2.0 API + 原生 DOM)
 */

// 1. 安全适配 Tauri 2.0 全局与降级调用
const getInvoke = () => {
  if (window.__TAURI__ && window.__TAURI__.core && window.__TAURI__.core.invoke) {
    return window.__TAURI__.core.invoke;
  }
  return async (cmd, args) => {
    console.warn(`[Mock Invoke] ${cmd}`, args);
    return { success: false, error: '当前运行在纯浏览器环境，未连接到 Tauri 桌面内核' };
  };
};

const getDialog = () => {
  if (window.__TAURI__ && window.__TAURI__.dialog) {
    return window.__TAURI__.dialog;
  }
  if (window.__TAURI_PLUGIN_DIALOG__) {
    return window.__TAURI_PLUGIN_DIALOG__;
  }
  return null;
};

const invoke = getInvoke();

// 全局应用状态
const state = {
  activeTab: 'tab-sales',
  currentSystem: 'internal',
  activeInternalTab: 'tab-sales',
  activeExternalTab: 'tab-ext-inbound',
  extInboundData: null,
  extInboundFilter: 'ALL',
  extOutboundData: null,
  extOutboundFilter: 'ALL',
  extCompareData: null,
  extCompareFilter: 'ALL',
  scannedFiles: null,
  configData: null,
  lastOutputs: {
    sales: null,
    outbound: null,
    inbound: null,
    compare: null,
    extInbound: null,
    extOutbound: null,
    extCompare: null
  },
  auditData: null,
  currentAuditCat: 'ALL',
  auditDiffOnly: false,
  auditPreview: null,
  previewCat: 'ALL',
  previewUnmatchedOnly: false,
  previewSearch: '',
  outboundManualMappings: {},
  inboundManualMappings: {},
  manualSearch: {
    outbound: {
      ledgerPath: '', queries: {}, results: {}, defaults: {}, requestIds: {},
      selectedLabels: {}, selectedCandidates: {}
    },
    inbound: {
      ledgerPath: '', queries: {}, results: {}, defaults: {}, requestIds: {},
      selectedLabels: {}, selectedCandidates: {}
    },
    audit: {
      ledgerPath: '', queries: {}, results: {}, defaults: {}, requestIds: {},
      selectedLabels: {}, selectedCandidates: {}
    }
  }
};

// HTML 转义防注入
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// UI 辅助工具函数：Toast 消息通知
function showToast(message, type = 'info', duration = 3500) {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  const messageEl = document.createElement('span');
  messageEl.textContent = message;
  toast.appendChild(messageEl);
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// UI 辅助工具函数：Loading 遮罩
function showLoading(msg = '正在处理中，请稍候...') {
  const overlay = document.getElementById('loading-overlay');
  const textEl = document.getElementById('loading-msg');
  if (textEl) textEl.textContent = msg;
  if (overlay) overlay.classList.remove('hidden');
}

function hideLoading() {
  const overlay = document.getElementById('loading-overlay');
  if (overlay) overlay.classList.add('hidden');
}

// 数字金额格式化 (千分位 + 两位小数)
function formatMoney(amount) {
  if (amount === undefined || amount === null || isNaN(amount)) return '¥ 0.00';
  return '¥ ' + Number(amount).toLocaleString('zh-CN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

// 通用原生文件选择器调用
async function pickExcelFile(title = '选择 Excel 文件') {
  const dialog = getDialog();
  if (dialog && dialog.open) {
    try {
      const selected = await dialog.open({
        title,
        multiple: false,
        filters: [{ name: 'Excel 工作簿', extensions: ['xlsx', 'xls'] }]
      });
      return selected;
    } catch (err) {
      console.error('打开文件对话框失败:', err);
    }
  }

  // 降级使用 HTML5 input
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.onchange = (e) => {
      const file = e.target.files[0];
      if (file) {
        // 在桌面 WebView 中，部分平台可获取 fullPath
        resolve(file.name);
      } else {
        resolve(null);
      }
    };
    input.click();
  });
}

// 统一拖拽高亮与文件路径填充绑定
function setupDropzone(inputEl, wrapperEl) {
  if (!wrapperEl || !inputEl) return;

  ['dragenter', 'dragover'].forEach(eventName => {
    wrapperEl.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      wrapperEl.style.borderColor = 'var(--primary)';
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    wrapperEl.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      wrapperEl.style.borderColor = '';
    });
  });

  wrapperEl.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      const f = files[0];
      // 检查后缀
      if (f.name.endsWith('.xlsx') || f.name.endsWith('.xls')) {
        // 如果在 WebView 里有 path 属性则用 path，否则用 name
        inputEl.value = f.path || f.name;
        delete inputEl.dataset.autoFilled;
        showToast(`已载入文件: ${f.name}`, 'info');
      } else {
        showToast('请拖入有效的 Excel 文件 (.xlsx 或 .xls)', 'warning');
      }
    }
  });
}

function bindFilePicker(buttonId, inputEl, title) {
  const button = document.getElementById(buttonId);
  if (!button || !inputEl) return;

  button.onclick = async () => {
    const file = await pickExcelFile(title);
    if (file) {
      inputEl.value = file;
      delete inputEl.dataset.autoFilled;
    }
  };
}

function setAutoFilledFile(inputEl, path) {
  if (!inputEl || !path) return;
  inputEl.value = path;
  inputEl.dataset.autoFilled = 'true';
}

function setManuallySelectedFile(inputEl, path) {
  if (!inputEl || !path) return;
  inputEl.value = path;
  delete inputEl.dataset.autoFilled;
}

function findUniqueScannedFile(files, predicate = () => true) {
  const matches = (Array.isArray(files) ? files : []).filter(predicate);
  return matches.length === 1 ? matches[0] : null;
}

async function invokeWithLoading(command, args, options = {}) {
  const {
    loadingMessage = '正在处理中，请稍候...',
    errorPrefix = '执行失败',
    errorDuration = 5000
  } = options;

  showLoading(loadingMessage);
  try {
    const result = await invoke(command, args);
    hideLoading();
    if (!result || !result.success) {
      showToast(`${errorPrefix}: ${result?.error || '未知错误'}`, 'error', errorDuration);
      return null;
    }
    return result;
  } catch (err) {
    hideLoading();
    showToast(`执行异常: ${err}`, 'error');
    return null;
  }
}

function resetManualSearchContext(kind, ledgerPath = '') {
  if (scheduleManualCandidateSearch.timers) {
    for (const key of scheduleManualCandidateSearch.timers.keys()) {
      if (key.startsWith(`${kind}:`)) {
        clearTimeout(scheduleManualCandidateSearch.timers.get(key));
        scheduleManualCandidateSearch.timers.delete(key);
      }
    }
  }
  state.manualSearch[kind] = {
    ledgerPath,
    queries: {},
    results: {},
    defaults: {},
    requestIds: {},
    selectedLabels: {},
    selectedCandidates: {}
  };
}

function candidateLabel(candidate) {
  const name = candidate.name || '未命名科目';
  const spec = candidate.spec ? ` ${candidate.spec}` : '';
  const qty = candidate.qty ?? 0;
  return `${name}${spec} (${candidate.code || '-'}) [期末:${qty}]`;
}

function getManualCandidates(kind, itemId) {
  const context = state.manualSearch[kind];
  if (!context) return [];
  return Object.prototype.hasOwnProperty.call(context.results, itemId)
    ? context.results[itemId]
    : context.defaults[itemId] || [];
}

function renderCandidateOptions(picker, candidates, selectedCode = '', options = {}) {
  if (!picker) return;
  const list = picker.querySelector('.candidate-options');
  if (!list) return;

  const {
    includeNone = false,
    emptyLabel = '-- 选择财务存货科目 --',
    noneLabel = '【置空/不关联】',
    selectedCandidate = null
  } = options;
  const visibleCandidates = Array.isArray(candidates) ? [...candidates] : [];
  const selected = selectedCode || '';
  if (
    selected &&
    selected !== '__NONE__' &&
    selectedCandidate?.code === selected &&
    !visibleCandidates.some(candidate => candidate.code === selected)
  ) {
    visibleCandidates.unshift(selectedCandidate);
  }

  list.innerHTML = '';
  if (visibleCandidates.length === 0) {
    const emptyOption = document.createElement('div');
    emptyOption.className = 'candidate-option-empty';
    emptyOption.textContent = emptyLabel;
    list.appendChild(emptyOption);
  }

  visibleCandidates.forEach(candidate => {
    const code = candidate.code || '';
    const option = document.createElement('button');
    option.type = 'button';
    option.className = `candidate-option${code === selected ? ' is-selected' : ''}`;
    option.dataset.code = code;
    option.setAttribute('role', 'option');
    option.setAttribute('aria-selected', String(code === selected));
    option.textContent = candidateLabel(candidate);
    list.appendChild(option);
  });
  if (includeNone) {
    const noneOption = document.createElement('button');
    noneOption.type = 'button';
    noneOption.className = `candidate-option candidate-option-none${selected === '__NONE__' ? ' is-selected' : ''}`;
    noneOption.dataset.code = '__NONE__';
    noneOption.setAttribute('role', 'option');
    noneOption.setAttribute('aria-selected', String(selected === '__NONE__'));
    noneOption.textContent = noneLabel;
    list.appendChild(noneOption);
  }
  picker.dataset.selectedCode = selected === '__NONE__' ? '' : selected;
}

function openCandidateOptions(picker) {
  const list = picker?.querySelector('.candidate-options');
  const input = picker?.querySelector('.candidate-search-input');
  if (!list) return;
  list.classList.remove('hidden');
  input?.setAttribute('aria-expanded', 'true');
}

function closeCandidateOptions(picker) {
  const list = picker?.querySelector('.candidate-options');
  const input = picker?.querySelector('.candidate-search-input');
  if (!list) return;
  list.classList.add('hidden');
  input?.setAttribute('aria-expanded', 'false');
}

function findManualCandidate(kind, itemId, code) {
  return [
    ...getManualCandidates(kind, itemId),
    ...(state.manualSearch[kind]?.defaults[itemId] || []),
    state.manualSearch[kind]?.selectedCandidates[itemId]
  ].find(candidate => candidate?.code === code) || null;
}

function scheduleManualCandidateSearch(input, picker, kind, selectedCode = '', category = '') {
  const itemId = String(input.dataset.id);
  const context = state.manualSearch[kind];
  if (!context) return;
  const query = input.value.trim();
  context.queries[itemId] = query;
  context.requestIds[itemId] = (context.requestIds[itemId] || 0) + 1;
  const requestId = context.requestIds[itemId];
  const timerKey = `${kind}:${itemId}`;

  if (!scheduleManualCandidateSearch.timers) {
    scheduleManualCandidateSearch.timers = new Map();
  }
  if (scheduleManualCandidateSearch.timers?.has(timerKey)) {
    clearTimeout(scheduleManualCandidateSearch.timers.get(timerKey));
  }

  const selected = selectedCode || picker?.dataset.selectedCode || '';
  const pickerOptions = {
    includeNone: kind === 'audit',
    emptyLabel: kind === 'audit' ? '-- 手动选择科目 --' : '-- 选择财务存货科目 --',
    selectedCandidate: context.selectedCandidates[itemId]
  };
  if (!query) {
    delete context.results[itemId];
    input.title = '输入品名、规格或科目编码进行模糊搜索';
    input.removeAttribute('aria-busy');
    renderCandidateOptions(picker, context.defaults[itemId] || [], selected, pickerOptions);
    openCandidateOptions(picker);
    return;
  }
  if (!context.ledgerPath) {
    renderCandidateOptions(picker, [], selected, {
      ...pickerOptions,
      emptyLabel: '尚未加载总账，无法搜索'
    });
    openCandidateOptions(picker);
    return;
  }

  input.title = '正在搜索总账科目…';
  input.setAttribute('aria-busy', 'true');
  renderCandidateOptions(picker, [], selected, {
    ...pickerOptions,
    emptyLabel: '正在搜索科目…'
  });
  openCandidateOptions(picker);

  const timer = setTimeout(async () => {
    try {
      const result = await invoke('search_ledger_candidates', {
        ledger: context.ledgerPath,
        query,
        limit: 80,
        category: category || null
      });
      if (
        state.manualSearch[kind] !== context ||
        context.requestIds[itemId] !== requestId ||
        context.queries[itemId] !== query
      ) {
        return;
      }
      if (!result?.success) {
        showToast(`科目搜索失败: ${result?.error || '未知错误'}`, 'error', 5000);
        return;
      }

      const candidates = Array.isArray(result.candidates) ? result.candidates : [];
      context.results[itemId] = candidates;
      const currentSelected = picker?.dataset.selectedCode || selected;
      input.removeAttribute('aria-busy');
      renderCandidateOptions(picker, candidates, currentSelected, {
        ...pickerOptions,
        selectedCandidate: context.selectedCandidates[itemId]
      });
      const emptyText = candidates.length === 0 ? '未找到匹配科目' : `找到 ${candidates.length} 个候选科目`;
      input.title = `${emptyText}，可继续输入缩小范围`;
      if (document.activeElement === input) {
        openCandidateOptions(picker);
      }
    } catch (error) {
      if (state.manualSearch[kind] === context && context.requestIds[itemId] === requestId) {
        input.removeAttribute('aria-busy');
        showToast(`科目搜索失败: ${error}`, 'error', 5000);
      }
    }
  }, 250);
  scheduleManualCandidateSearch.timers.set(timerKey, timer);
}

function bindManualCandidatePickers(table, kind, getSelectedCode, getCategory, onSelect) {
  table?.querySelectorAll('.manual-candidate-picker').forEach(picker => {
    const input = picker.querySelector('.candidate-search-input');
    const list = picker.querySelector('.candidate-options');
    if (!input || !list) return;

    const itemId = String(picker.dataset.id);
    const context = state.manualSearch[kind];
    picker.dataset.selectedCode = getSelectedCode(itemId) || '';

    input.onfocus = () => {
      input.select();
      const selected = getSelectedCode(itemId) || picker.dataset.selectedCode || '';
      renderCandidateOptions(picker, getManualCandidates(kind, itemId), selected, {
        includeNone: kind === 'audit',
        emptyLabel: kind === 'audit' ? '-- 手动选择科目 --' : '-- 选择财务存货科目 --',
        selectedCandidate: context.selectedCandidates[itemId]
      });
      openCandidateOptions(picker);
    };

    input.oninput = () => {
      scheduleManualCandidateSearch(
        input,
        picker,
        kind,
        getSelectedCode(itemId) || picker.dataset.selectedCode || '',
        getCategory(itemId) || ''
      );
    };

    input.onkeydown = (event) => {
      if (event.key === 'Escape') {
        closeCandidateOptions(picker);
        input.blur();
      }
    };

    input.onblur = () => {
      setTimeout(() => {
        if (!picker.contains(document.activeElement)) {
          closeCandidateOptions(picker);
        }
      }, 120);
    };

    list.onmousedown = (event) => {
      const option = event.target.closest('.candidate-option');
      if (!option) return;
      event.preventDefault();

      const code = option.dataset.code || '';
      const candidate = code === '__NONE__' ? null : findManualCandidate(kind, itemId, code);
      const normalizedCode = code === '__NONE__' ? '' : code;
      context.queries[itemId] = '';
      delete context.results[itemId];
      if (normalizedCode) {
        context.selectedLabels[itemId] = candidate ? candidateLabel(candidate) : normalizedCode;
        if (candidate) context.selectedCandidates[itemId] = candidate;
      } else {
        delete context.selectedLabels[itemId];
        delete context.selectedCandidates[itemId];
      }
      picker.dataset.selectedCode = normalizedCode;

      onSelect(itemId, normalizedCode, candidate);
      if (picker.isConnected) {
        input.value = normalizedCode
          ? (candidate ? candidateLabel(candidate) : context.selectedLabels[itemId] || normalizedCode)
          : '';
        renderCandidateOptions(picker, context.defaults[itemId] || [], normalizedCode, {
          includeNone: kind === 'audit',
          emptyLabel: kind === 'audit' ? '-- 手动选择科目 --' : '-- 选择财务存货科目 --',
          selectedCandidate: context.selectedCandidates[itemId]
        });
        closeCandidateOptions(picker);
      }
    };
  });
}

// 打开系统文件与所在文件夹
async function openSystemPath(path) {
  if (!path) return;
  try {
    const res = await invoke('open_in_system', { path });
    if (!res) showToast('无法直接打开该文件', 'warning');
  } catch (err) {
    showToast(`打开文件失败: ${err}`, 'error');
  }
}

async function showInSystemFolder(path) {
  if (!path) return;
  try {
    await invoke('show_in_folder', { path });
  } catch (err) {
    showToast(`在文件夹中显示失败: ${err}`, 'error');
  }
}

// ----------------------------------------------------
// 2. 自动化文件扫描与智能预填
// ----------------------------------------------------
async function triggerFileScan() {
  try {
    const res = await invoke('scan_files', { dir: null });
    if (res && res.success && res.data) {
      state.scannedFiles = res.data;
      renderScanBanner(res.data);
      autoFillDetectedFiles(res.data);
      showToast('工作区候选文件扫描成功', 'success');
    }
  } catch (err) {
    console.error('扫描文件失败:', err);
  }
}

function renderScanBanner(data) {
  const banner = document.getElementById('scan-banner');
  const linksContainer = document.getElementById('scan-quick-links');
  if (!banner || !linksContainer) return;

  linksContainer.innerHTML = '';
  const all = data.all_excel || [];
  if (all.length === 0) {
    banner.classList.add('hidden');
    return;
  }

  banner.classList.remove('hidden');
  all.forEach(item => {
    const chip = document.createElement('span');
    chip.className = 'scan-chip';
    chip.textContent = item.name;
    chip.title = `点击填入当前页面对应输入框\n完整路径: ${item.path}`;
    chip.onclick = () => {
      applyScannedFileToActiveTab(item);
    };
    linksContainer.appendChild(chip);
  });
}

function autoFillDetectedFiles(data) {
  const ledgerFiles = Array.isArray(data.ledger_files) ? data.ledger_files : [];
  const templateFiles = Array.isArray(data.template_files) ? data.template_files : [];

  // 销售明细输入框
  const salesInput = document.getElementById('sales-input-file');
  const rawSales = findUniqueScannedFile(data.sales_files, file => !file.name.includes('已汇总'));
  if (salesInput && !salesInput.value && rawSales) {
    setAutoFilledFile(salesInput, rawSales.path);
  }

  // 出库凭证输入框
  const obSales = document.getElementById('outbound-sales-file');
  const obLedger = document.getElementById('outbound-ledger-file');
  const obTmpl = document.getElementById('outbound-template-file');
  const rolledSales = findUniqueScannedFile(data.sales_files, file => file.name.includes('已汇总'));
  if (obSales && !obSales.value && rolledSales) {
    // 优先选择带有“已汇总”字样的文件
    setAutoFilledFile(obSales, rolledSales.path);
  }
  if (obLedger && !obLedger.value && ledgerFiles.length === 1) {
    setAutoFilledFile(obLedger, ledgerFiles[0].path);
  }
  const outboundTemplate = findUniqueScannedFile(
    templateFiles,
    file => !file.name.includes('入库') && !file.name.includes('中药')
  );
  if (obTmpl && !obTmpl.value && outboundTemplate) {
    setAutoFilledFile(obTmpl, outboundTemplate.path);
  }

  // 入库凭证输入框
  const inInbound = document.getElementById('inbound-file');
  const inLedger = document.getElementById('inbound-ledger-file');
  const inTmpl = document.getElementById('inbound-template-file');
  const inboundFile = findUniqueScannedFile(data.inbound_files);
  if (inInbound && !inInbound.value && inboundFile) {
    setAutoFilledFile(inInbound, inboundFile.path);
  }
  if (inLedger && !inLedger.value && ledgerFiles.length === 1) {
    setAutoFilledFile(inLedger, ledgerFiles[0].path);
  }
  const inboundTemplate = findUniqueScannedFile(
    templateFiles,
    file => file.name.includes('入库') || file.name.includes('中药')
  );
  if (inTmpl && !inTmpl.value && inboundTemplate) {
    setAutoFilledFile(inTmpl, inboundTemplate.path);
  }

  // 账实库存核对输入框
  const cmpLedger = document.getElementById('compare-ledger-file');
  const cmpWest = document.getElementById('compare-west-file');
  const cmpTcm = document.getElementById('compare-tcm-file');
  const cmpHc = document.getElementById('compare-hc-file');
  if (cmpLedger && !cmpLedger.value && ledgerFiles.length === 1) {
    setAutoFilledFile(cmpLedger, ledgerFiles[0].path);
  }
  const westFile = findUniqueScannedFile(data.west_wh_files);
  if (cmpWest && !cmpWest.value && westFile) {
    setAutoFilledFile(cmpWest, westFile.path);
  }
  const tcmFile = findUniqueScannedFile(data.tcm_wh_files);
  if (cmpTcm && !cmpTcm.value && tcmFile) {
    setAutoFilledFile(cmpTcm, tcmFile.path);
  }
  const hcFile = findUniqueScannedFile(data.hc_wh_files);
  if (cmpHc && !cmpHc.value && hcFile) {
    setAutoFilledFile(cmpHc, hcFile.path);
  }

  // 外账入库凭证输入框
  const extInbound = document.getElementById('ext-inbound-file');
  const extTmpl = document.getElementById('ext-template-file');
  if (extInbound && !extInbound.value && inboundFile) {
    setAutoFilledFile(extInbound, inboundFile.path);
  }
  const extTemplateCandidate = findUniqueScannedFile(
    data.all_excel || [],
    file => file.name.includes('迁账') || file.name.includes('外账')
  );
  if (extTmpl && !extTmpl.value && extTemplateCandidate) {
    setAutoFilledFile(extTmpl, extTemplateCandidate.path);
  }

  // 外账出库凭证输入框
  const extOutSales = document.getElementById('ext-outbound-sales-file');
  const extOutTmpl = document.getElementById('ext-outbound-template-file');
  if (extOutSales && !extOutSales.value && salesFile) {
    setAutoFilledFile(extOutSales, salesFile.path);
  }
  if (extOutTmpl && !extOutTmpl.value && extTemplateCandidate) {
    setAutoFilledFile(extOutTmpl, extTemplateCandidate.path);
  }

  // 外账结存数比对输入框
  const extCmpTmpl = document.getElementById('ext-compare-template-file');
  const extCmpWh = document.getElementById('ext-compare-wh-file');
  if (extCmpTmpl && !extCmpTmpl.value && extTemplateCandidate) {
    setAutoFilledFile(extCmpTmpl, extTemplateCandidate.path);
  }
  if (extCmpWh && !extCmpWh.value && westFile) {
    setAutoFilledFile(extCmpWh, westFile.path);
  }

  if (ledgerFiles.length > 1) {
    [obLedger, inLedger, cmpLedger].forEach(input => {
      if (input?.dataset.autoFilled === 'true') {
        input.value = '';
        delete input.dataset.autoFilled;
      }
    });
    showToast('检测到多个总账文件：请按业务阶段手工选择对应文件，避免把入库前总账误用于出库。', 'warning', 7000);
  }
}

function applyScannedFileToActiveTab(item) {
  const name = item.name;
  if (state.activeTab === 'tab-sales') {
    setManuallySelectedFile(document.getElementById('sales-input-file'), item.path);
    showToast(`已选定销售表: ${name}`, 'info');
  } else if (state.activeTab === 'tab-outbound') {
    if (name.includes('总账') || name.includes('数量金额')) {
      setManuallySelectedFile(document.getElementById('outbound-ledger-file'), item.path);
      showToast(`已填入出库前总账（入库后快照）: ${name}`, 'info');
    } else if (name.includes('模板')) {
      setManuallySelectedFile(document.getElementById('outbound-template-file'), item.path);
      showToast(`已填入凭证模板: ${name}`, 'info');
    } else {
      setManuallySelectedFile(document.getElementById('outbound-sales-file'), item.path);
      showToast(`已填入销售汇总表: ${name}`, 'info');
    }
  } else if (state.activeTab === 'tab-inbound') {
    if (name.includes('总账') || name.includes('数量金额')) {
      setManuallySelectedFile(document.getElementById('inbound-ledger-file'), item.path);
      showToast(`已填入入库前总账: ${name}`, 'info');
    } else if (name.includes('模板')) {
      setManuallySelectedFile(document.getElementById('inbound-template-file'), item.path);
      showToast(`已填入凭证模板: ${name}`, 'info');
    } else {
      setManuallySelectedFile(document.getElementById('inbound-file'), item.path);
      showToast(`已填入入库单: ${name}`, 'info');
    }
  } else if (state.activeTab === 'tab-compare') {
    if (name.includes('总账') || name.includes('数量金额')) {
      setManuallySelectedFile(document.getElementById('compare-ledger-file'), item.path);
      showToast(`已填入财务总账: ${name}`, 'info');
    } else if (name.includes('西药')) {
      setManuallySelectedFile(document.getElementById('compare-west-file'), item.path);
      showToast(`已填入西药房库存表: ${name}`, 'info');
    } else if (name.includes('中药')) {
      setManuallySelectedFile(document.getElementById('compare-tcm-file'), item.path);
      showToast(`已填入中药房库存表: ${name}`, 'info');
    } else if (name.includes('耗材') || name.includes('材料')) {
      setManuallySelectedFile(document.getElementById('compare-hc-file'), item.path);
      showToast(`已填入耗材库库存表: ${name}`, 'info');
    } else {
      setManuallySelectedFile(document.getElementById('compare-west-file'), item.path);
      showToast(`已填入库管库存表: ${name}`, 'info');
    }
  } else if (state.activeTab === 'tab-ext-inbound') {
    if (name.includes('迁账') || name.includes('外账') || name.includes('模板')) {
      setManuallySelectedFile(document.getElementById('ext-template-file'), item.path);
      showToast(`已填入外账参考模板: ${name}`, 'info');
    } else {
      setManuallySelectedFile(document.getElementById('ext-inbound-file'), item.path);
      showToast(`已填入入库单: ${name}`, 'info');
    }
  } else if (state.activeTab === 'tab-ext-outbound') {
    if (name.includes('迁账') || name.includes('外账') || name.includes('模板')) {
      setManuallySelectedFile(document.getElementById('ext-outbound-template-file'), item.path);
      showToast(`已填入外账参考模板: ${name}`, 'info');
    } else {
      setManuallySelectedFile(document.getElementById('ext-outbound-sales-file'), item.path);
      showToast(`已填入销售汇总表: ${name}`, 'info');
    }
  } else if (state.activeTab === 'tab-ext-compare') {
    if (name.includes('迁账') || name.includes('外账') || name.includes('模板')) {
      setManuallySelectedFile(document.getElementById('ext-compare-template-file'), item.path);
      showToast(`已填入外账参考模板: ${name}`, 'info');
    } else {
      setManuallySelectedFile(document.getElementById('ext-compare-wh-file'), item.path);
      showToast(`已填入库管在库报表: ${name}`, 'info');
    }
  }
}

// ----------------------------------------------------
// 3. TAB 1: 销售汇总处理
// ----------------------------------------------------
function initTabSales() {
  const inputSales = document.getElementById('sales-input-file');
  const btnRun = document.getElementById('btn-run-sales');
  const inputSheet = document.getElementById('sales-sheet-name');
  const inputOutput = document.getElementById('sales-output-name');

  bindFilePicker('btn-browse-sales', inputSales, '选择原始销售明细表');

  setupDropzone(inputSales, inputSales.closest('.file-input-wrapper'));

  btnRun.onclick = async () => {
    const file = inputSales.value.trim();
    if (!file) {
      showToast('请先选择或拖入原始销售明细表', 'warning');
      return;
    }

    const sheetName = inputSheet.value.trim() || null;
    const output = inputOutput.value.trim() || null;

    const res = await invokeWithLoading('execute_sales_process', {
      file,
      output,
      sheetName
    }, {
      loadingMessage: '正在对销售明细进行去重、合并及多维度汇总计算...',
      errorPrefix: '销售汇总失败'
    });
    if (!res) return;

    state.lastOutputs.sales = res.output_file;
    renderSalesResult(res);
    showToast('销售明细汇总处理完成！', 'success');

    // 智能联动：如果出库面板中的销售表仍是扫描自动填入的旧文件，则替换为本次新汇总结果；
    // 用户手工选定的文件不覆盖。
    const obSales = document.getElementById('outbound-sales-file');
    if (obSales && (!obSales.value || obSales.dataset.autoFilled === 'true')) {
      setAutoFilledFile(obSales, res.output_file);
    }
  };

  document.getElementById('btn-open-sales-file').onclick = () => {
    openSystemPath(state.lastOutputs.sales);
  };
  document.getElementById('btn-show-sales-folder').onclick = () => {
    showInSystemFolder(state.lastOutputs.sales);
  };
}

function renderSalesResult(data) {
  const resultBox = document.getElementById('sales-result');
  resultBox.classList.remove('hidden');

  const totals = data.totals || {
    original_count: data.total_raw_rows,
    unique_count: data.unique_drugs_count,
    total_qty: data.total_qty,
    total_in_amt: data.total_cost_amt,
    total_retail_amt: data.total_retail_amt
  };
  document.getElementById('stat-sales-orig').textContent = totals.original_count?.toLocaleString() || '-';
  document.getElementById('stat-sales-unique').textContent = totals.unique_count?.toLocaleString() || '-';
  document.getElementById('stat-sales-qty').textContent = totals.total_qty?.toLocaleString() || '-';
  document.getElementById('stat-sales-in-amt').textContent = formatMoney(totals.total_in_amt);
  document.getElementById('stat-sales-retail-amt').textContent = formatMoney(totals.total_retail_amt);
}

// ----------------------------------------------------
// 3. TAB 3: 销售出库凭证生成
// ----------------------------------------------------
function initTabOutbound() {
  const inSales = document.getElementById('outbound-sales-file');
  const inLedger = document.getElementById('outbound-ledger-file');
  const inTemplate = document.getElementById('outbound-template-file');
  const inDate = document.getElementById('outbound-date');
  const chkFallback = document.getElementById('outbound-fallback-price');
  const btnRun = document.getElementById('btn-run-outbound');
  const btnManualRun = document.getElementById('btn-rerun-outbound-manual');

  bindFilePicker('btn-browse-outbound-sales', inSales, '选择销售汇总表');
  bindFilePicker('btn-browse-outbound-ledger', inLedger, '选择入库凭证导入后重新导出的数量金额总账表');
  bindFilePicker('btn-browse-outbound-template', inTemplate, '选择凭证导入模板');

  setupDropzone(inSales, inSales.closest('.file-input-wrapper'));
  setupDropzone(inLedger, inLedger.closest('.file-input-wrapper'));
  setupDropzone(inTemplate, inTemplate.closest('.file-input-wrapper'));

  const generateOutbound = async (confirmedItems, successMessage) => {
    const sales = inSales.value.trim();
    const ledger = inLedger.value.trim();
    const template = inTemplate.value.trim();

    if (!sales) return showToast('请指定销售汇总表', 'warning');
    if (!ledger) return showToast('请指定数量金额总账表', 'warning');
    if (!template) return showToast('请指定凭证导入模板', 'warning');

    resetManualSearchContext('outbound', ledger);

    const dateVal = inDate.value.trim() || null;
    const fallback = chkFallback.checked;

    const res = await invokeWithLoading('execute_outbound_voucher', {
      sales,
      ledger,
      template,
      output: null,
      date: dateVal,
      fallbackPrice: fallback,
      confirmedItems: confirmedItems || null,
      config: null
    }, {
      loadingMessage: '正在匹配总账存货编码与单价，生成销售出库凭证...',
      errorPrefix: '生成出库凭证失败',
      errorDuration: 6000
    });
    if (!res) return;

    state.lastOutputs.outbound = res.output_file;
    renderOutboundResult(res);
    showToast(successMessage, 'success');
  };

  btnRun.onclick = async () => {
    state.outboundManualMappings = {};
    await generateOutbound(null, '销售出库凭证生成成功！');
  };

  if (btnManualRun) {
    btnManualRun.onclick = async () => {
      const confirmedItems = collectConfirmedLedgerMappings(state.outboundManualMappings);
      if (confirmedItems.length === 0) {
        showToast('请先在未匹配列表中选择至少一个财务存货科目', 'warning');
        return;
      }
      await generateOutbound(confirmedItems, `已按 ${confirmedItems.length} 项手工匹配重新生成出库凭证！`);
    };
  }

  document.getElementById('btn-open-outbound-file').onclick = () => {
    openSystemPath(state.lastOutputs.outbound);
  };
  document.getElementById('btn-show-outbound-folder').onclick = () => {
    showInSystemFolder(state.lastOutputs.outbound);
  };
}

// 诊断未匹配药品原因
function diagnoseUnmatchedReason(item) {
  const name = item.name || item.target_name || '';
  const factory = item.factory || item.supplier || '';
  
  if (name.includes('海螵蛸')) {
    return {
      text: '严格厂家防串户：总账未建蕴德新批次',
      type: 'strict-vendor'
    };
  }
  
  const strictList = state.configData?.strict_vendor_suffix_drugs || [];
  if (strictList.some(s => name.includes(s))) {
    return {
      text: '严格厂家锁定药：总账无此厂家科目',
      type: 'strict-vendor'
    };
  }

  return {
    text: '总账中尚未建档此品规存货科目',
    type: 'new-drug'
  };
}

// 导出未建档药品为 CSV (带 UTF-8 BOM，Excel 双击无乱码)
function exportUnmatchedToCsv(items, filename = '未匹配建档药品清单.csv') {
  if (!items || items.length === 0) {
    showToast('当前没有未匹配药品可导出', 'info');
    return;
  }

  let csvContent = '\uFEFF药品名称,规格,生产厂家/供应商,出入库数量,进价,未匹配诊断\n';
  items.forEach(it => {
    const diag = diagnoseUnmatchedReason(it).text;
    const name = `"${(it.name || it.target_name || '').replace(/"/g, '""')}"`;
    const spec = `"${(it.spec || '').replace(/"/g, '""')}"`;
    const fac = `"${(it.factory || it.supplier || '').replace(/"/g, '""')}"`;
    const qty = it.qty ?? 0;
    const price = it.in_price ?? 0;
    csvContent += `${name},${spec},${fac},${qty},${price},"${diag}"\n`;
  });

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  showToast(`已成功导出 ${items.length} 条未匹配药品清单`, 'success');
}

// 复制未匹配清单至剪贴板
function copyUnmatchedToClipboard(items) {
  if (!items || items.length === 0) {
    showToast('当前没有未匹配药品可复制', 'info');
    return;
  }

  let text = `【未匹配新药品清单】共 ${items.length} 种：\n`;
  items.forEach((it, idx) => {
    const diag = diagnoseUnmatchedReason(it).text;
    const fac = it.factory || it.supplier || '无厂家';
    text += `${idx + 1}. ${it.name || it.target_name} | 规格: ${it.spec || '-'} | 厂家: ${fac} | 数量: ${it.qty} | 诊断: ${diag}\n`;
  });

  navigator.clipboard.writeText(text).then(() => {
    showToast('已复制未匹配药品清单到剪贴板，可直接粘贴分享', 'success');
  }).catch(err => {
    showToast(`复制失败: ${err}`, 'error');
  });
}

// 点击未匹配药品，一键跳转到配置中心
function jumpToConfigWithDrug(drugName) {
  const tabBtn = document.querySelector('.nav-tab[data-tab="tab-config"]');
  if (tabBtn) tabBtn.click();

  const keyInput = document.getElementById('override-key-input');
  const valInput = document.getElementById('override-val-input');
  if (keyInput) {
    keyInput.value = drugName;
    if (valInput) valInput.focus();
    showToast(`已将【${drugName}】载入字典配置，请在光标处输入科目编码后点击添加`, 'info', 5000);
  }
}

// 凭证未匹配项的人工科目选择：与账实核对预览使用同一份候选数据结构。
function renderVoucherCandidateSelect(item, mappings, kind) {
  const itemId = String(item.id);
  const context = state.manualSearch[kind];
  const defaultCandidates = Array.isArray(item.candidates) ? item.candidates : [];
  if (!Object.prototype.hasOwnProperty.call(context.defaults, itemId)) {
    context.defaults[itemId] = defaultCandidates;
  }

  const selectedCode = mappings[itemId] || '';
  const selectedCandidate = context.selectedCandidates[itemId]
    || defaultCandidates.find(candidate => candidate.code === selectedCode);
  const selectedLabel = context.selectedLabels[itemId]
    || (selectedCandidate ? candidateLabel(selectedCandidate) : selectedCode);
  const searchValue = context.queries[itemId] || '';
  const inputValue = searchValue || (selectedCode ? selectedLabel : '');
  const optionsId = `${kind}-candidate-options-${itemId}`;

  return `
    <div class="manual-candidate-picker" data-id="${escapeHtml(item.id)}">
      <input
        class="candidate-search-input"
        data-id="${escapeHtml(item.id)}"
        type="search"
        value="${escapeHtml(inputValue)}"
        placeholder="模糊搜索品名/规格/编码"
        title="输入品名、规格或科目编码进行模糊搜索"
        role="combobox"
        aria-autocomplete="list"
        aria-controls="${escapeHtml(optionsId)}"
        aria-expanded="false"
        autocomplete="off"
      />
      <div class="candidate-options hidden" id="${escapeHtml(optionsId)}" role="listbox"></div>
    </div>
  `;
}

function collectConfirmedLedgerMappings(mappings) {
  return Object.entries(mappings || {})
    .map(([id, ledgerCode]) => ({
      id: Number(id),
      ledger_code: String(ledgerCode || '').trim()
    }))
    .filter(item => Number.isInteger(item.id) && item.id >= 0 && item.ledger_code);
}

function updateVoucherManualButton(buttonId, mappings, visible) {
  const button = document.getElementById(buttonId);
  if (!button) return;
  button.classList.toggle('hidden', !visible);
  button.disabled = collectConfirmedLedgerMappings(mappings).length === 0;
}

function bindVoucherManualSelectors(tableId, stateKey, buttonId, kind) {
  const table = document.getElementById(tableId);
  if (!table) return;

  bindManualCandidatePickers(
    table,
    kind,
    itemId => state[stateKey][itemId] || '',
    () => '',
    (itemId, selectedCode) => {
      if (selectedCode) {
        state[stateKey][itemId] = selectedCode;
      } else {
        delete state[stateKey][itemId];
      }
      updateVoucherManualButton(buttonId, state[stateKey], true);
    }
  );
}

function renderOutboundResult(data) {
  const box = document.getElementById('outbound-result');
  box.classList.remove('hidden');

  document.getElementById('stat-outbound-total').textContent = data.total_items?.toLocaleString() || '-';
  document.getElementById('stat-outbound-rate').textContent = `${data.match_rate}%`;
  document.getElementById('stat-outbound-matched').textContent = data.matched_count?.toLocaleString() || '-';
  document.getElementById('stat-outbound-unmatched').textContent = data.unmatched_count?.toLocaleString() || '0';
  document.getElementById('stat-outbound-amt').textContent = formatMoney(data.total_credit_amt);
  const depletionSummary = document.getElementById('outbound-depletion-summary');
  if (depletionSummary) {
    const full = data.fully_depleted_count ?? 0;
    const partial = data.partially_depleted_count ?? 0;
    const tail = formatMoney(data.tail_adjustment_amt ?? 0);
    depletionSummary.textContent = `库存状态：${full} 个品规本次出库后数量清零，${partial} 个品规部分出库；已将清零品规的总账尾差摊入其最后一笔出库，尾差调整合计 ${tail}。`;
  }

  // 未匹配新药表格
  const unmatchedBox = document.getElementById('outbound-unmatched-box');
  const tbody = document.getElementById('outbound-unmatched-table').querySelector('tbody');
  tbody.innerHTML = '';

  const unmatched = data.unmatched_items || [];
  const countEl = document.getElementById('outbound-unmatched-count-text');
  if (countEl) countEl.textContent = unmatched.length;

  if (unmatched.length > 0) {
    unmatchedBox.classList.remove('hidden');
    unmatched.forEach(item => {
      const diag = diagnoseUnmatchedReason(item);
      const drugDisplayName = item.name || item.target_name || '';
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-weight: 600; color: #fff;">${escapeHtml(drugDisplayName)}</td>
        <td>${escapeHtml(item.spec || '-')}</td>
        <td>${escapeHtml(item.factory || '-')}</td>
        <td style="font-family: monospace;">${escapeHtml(item.qty)}</td>
        <td style="font-family: monospace;">${item.in_price !== null ? '¥ ' + escapeHtml(item.in_price) : '-'}</td>
        <td><span class="diag-tag ${escapeHtml(diag.type)}">${escapeHtml(diag.text)}</span></td>
        <td>
          <div class="manual-match-cell">
            ${renderVoucherCandidateSelect(item, state.outboundManualMappings, 'outbound')}
            <button class="btn btn-outline btn-xs btn-jump-config" title="在字典配置中心为该药品设置编码">
              去配置中心
            </button>
          </div>
        </td>
      `;
      tr.querySelector('.btn-jump-config').onclick = () => jumpToConfigWithDrug(drugDisplayName);
      tbody.appendChild(tr);
    });

    bindVoucherManualSelectors(
      'outbound-unmatched-table',
      'outboundManualMappings',
      'btn-rerun-outbound-manual',
      'outbound'
    );
    updateVoucherManualButton(
      'btn-rerun-outbound-manual',
      state.outboundManualMappings,
      true
    );

    // 绑定导出与复制按钮
    document.getElementById('btn-copy-outbound-unmatched').onclick = () => copyUnmatchedToClipboard(unmatched);
    document.getElementById('btn-export-outbound-unmatched').onclick = () => exportUnmatchedToCsv(unmatched, '销售出库未建档药品清单.csv');
  } else {
    closeManualWorkspace(unmatchedBox);
    unmatchedBox.classList.add('hidden');
    updateVoucherManualButton('btn-rerun-outbound-manual', state.outboundManualMappings, false);
  }
}

// ----------------------------------------------------
// 2. TAB 2: 药房入库凭证生成 (西药 / 中药)
// ----------------------------------------------------
function initTabInbound() {
  const inInbound = document.getElementById('inbound-file');
  const inLedger = document.getElementById('inbound-ledger-file');
  const inTemplate = document.getElementById('inbound-template-file');
  const inVoucherNo = document.getElementById('inbound-voucher-no');
  const inDate = document.getElementById('inbound-date');
  const btnRun = document.getElementById('btn-run-inbound');
  const btnManualRun = document.getElementById('btn-rerun-inbound-manual');

  bindFilePicker('btn-browse-inbound', inInbound, '选择药品入库单 (西药或中药)');
  bindFilePicker('btn-browse-inbound-ledger', inLedger, '选择入库凭证生成前的数量金额总账表');
  bindFilePicker('btn-browse-inbound-template', inTemplate, '选择凭证导入模板');

  setupDropzone(inInbound, inInbound.closest('.file-input-wrapper'));
  setupDropzone(inLedger, inLedger.closest('.file-input-wrapper'));
  setupDropzone(inTemplate, inTemplate.closest('.file-input-wrapper'));

  const generateInbound = async (confirmedItems, successMessage) => {
    const inbound = inInbound.value.trim();
    const ledger = inLedger.value.trim();
    const template = inTemplate.value.trim();

    if (!inbound) return showToast('请指定药品入库单', 'warning');
    if (!ledger) return showToast('请指定数量金额总账表', 'warning');
    if (!template) return showToast('请指定凭证导入模板', 'warning');

    resetManualSearchContext('inbound', ledger);

    const voucherNo = inVoucherNo.value.trim() || null;
    const dateVal = inDate.value.trim() || null;

    const res = await invokeWithLoading('execute_inbound_voucher', {
      inbound,
      ledger,
      template,
      output: null,
      date: dateVal,
      voucherNo,
      confirmedItems: confirmedItems || null,
      config: null
    }, {
      loadingMessage: '正在解析入库单据、匹配供应商与存货编码并进行借贷平衡审计...',
      errorPrefix: '生成入库凭证失败',
      errorDuration: 6000
    });
    if (!res) return;

    state.lastOutputs.inbound = res.output_file;
    renderInboundResult(res);
    showToast(successMessage, 'success');
  };

  btnRun.onclick = async () => {
    state.inboundManualMappings = {};
    await generateInbound(null, '药房入库凭证生成成功！');
  };

  if (btnManualRun) {
    btnManualRun.onclick = async () => {
      const confirmedItems = collectConfirmedLedgerMappings(state.inboundManualMappings);
      if (confirmedItems.length === 0) {
        showToast('请先在未匹配列表中选择至少一个财务存货科目', 'warning');
        return;
      }
      await generateInbound(confirmedItems, `已按 ${confirmedItems.length} 项手工匹配重新生成入库凭证！`);
    };
  }

  document.getElementById('btn-open-inbound-file').onclick = () => {
    openSystemPath(state.lastOutputs.inbound);
  };
  document.getElementById('btn-show-inbound-folder').onclick = () => {
    showInSystemFolder(state.lastOutputs.inbound);
  };
}

function renderInboundResult(data) {
  const box = document.getElementById('inbound-result');
  box.classList.remove('hidden');

  // 借贷审计仪表盘
  const debitAmt = data.total_debit_amt || 0;
  const creditAmt = data.total_credit_amt || 0;
  const diff = Math.abs(debitAmt - creditAmt);

  document.getElementById('audit-debit-amt').textContent = formatMoney(debitAmt);
  document.getElementById('audit-debit-desc').textContent = `共 ${data.total_items} 笔入库明细 (借方)`;
  document.getElementById('audit-credit-amt').textContent = formatMoney(creditAmt);
  document.getElementById('audit-credit-desc').textContent = `共 ${data.suppliers_summary?.length || 0} 家供应商往来 (贷方)`;

  const indicatorBox = document.getElementById('balance-indicator-box');
  const statusText = document.getElementById('balance-status-text');
  const diffText = document.getElementById('balance-diff-text');

  if (data.is_balanced) {
    indicatorBox.className = 'balance-indicator balanced';
    statusText.textContent = '借贷完美平衡';
    diffText.textContent = '差额: ¥ 0.00';
  } else {
    indicatorBox.className = 'balance-indicator unbalanced';
    statusText.textContent = '借贷存在差额！';
    diffText.textContent = `差额: ${formatMoney(diff)}`;
  }

  // 指标卡片
  document.getElementById('stat-inbound-items').textContent = data.total_items?.toLocaleString() || '-';
  document.getElementById('stat-inbound-rate').textContent = `${data.match_rate}%`;
  document.getElementById('stat-inbound-matched').textContent = data.matched_count?.toLocaleString() || '-';
  document.getElementById('stat-inbound-unmatched').textContent = data.unmatched_count?.toLocaleString() || '0';
  document.getElementById('stat-inbound-entries').textContent = data.total_entries?.toLocaleString() || '-';

  // 供应商汇总表格
  const supTbody = document.getElementById('inbound-supplier-table').querySelector('tbody');
  supTbody.innerHTML = '';
  (data.suppliers_summary || []).forEach(s => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-weight: 600; color: #fff;">${escapeHtml(s.supplier)}</td>
      <td style="font-family: monospace; color: #38bdf8;">${s.code ? escapeHtml(s.code) : '<span style="color:#f87171">未匹配编码</span>'}</td>
      <td>${escapeHtml(s.brief || '-')}</td>
      <td style="font-family: monospace;">${escapeHtml(s.count)} 条</td>
      <td style="font-family: monospace; font-weight: 700; color: #34d399;">${formatMoney(s.amount)}</td>
    `;
    supTbody.appendChild(tr);
  });

  // 未匹配新药表格
  const unmatchedBox = document.getElementById('inbound-unmatched-box');
  const tbody = document.getElementById('inbound-unmatched-table').querySelector('tbody');
  tbody.innerHTML = '';
  const unmatched = data.unmatched_items || [];
  const countEl = document.getElementById('inbound-unmatched-count-text');
  if (countEl) countEl.textContent = unmatched.length;

  if (unmatched.length > 0) {
    unmatchedBox.classList.remove('hidden');
    unmatched.forEach(item => {
      const diag = diagnoseUnmatchedReason(item);
      const drugDisplayName = item.target_name || item.name || '';
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(item.supplier || '-')}</td>
        <td style="font-weight: 600; color: #fbbf24;">${escapeHtml(drugDisplayName)}</td>
        <td>${escapeHtml(item.spec || '-')}</td>
        <td style="font-family: monospace;">${escapeHtml(item.qty)}</td>
        <td style="font-family: monospace;">¥ ${escapeHtml(item.in_price)}</td>
        <td style="font-family: monospace; font-weight: 600;">${formatMoney(item.in_amt)}</td>
        <td><span class="diag-tag ${escapeHtml(diag.type)}">${escapeHtml(diag.text)}</span></td>
        <td>
          <div class="manual-match-cell">
            ${renderVoucherCandidateSelect(item, state.inboundManualMappings, 'inbound')}
            <button class="btn btn-outline btn-xs btn-jump-config" title="在字典配置中心为该药品设置编码">
              去配置中心
            </button>
          </div>
        </td>
      `;
      tr.querySelector('.btn-jump-config').onclick = () => jumpToConfigWithDrug(drugDisplayName);
      tbody.appendChild(tr);
    });

    bindVoucherManualSelectors(
      'inbound-unmatched-table',
      'inboundManualMappings',
      'btn-rerun-inbound-manual',
      'inbound'
    );
    updateVoucherManualButton(
      'btn-rerun-inbound-manual',
      state.inboundManualMappings,
      true
    );

    // 绑定导出与复制按钮
    document.getElementById('btn-copy-inbound-unmatched').onclick = () => copyUnmatchedToClipboard(unmatched);
    document.getElementById('btn-export-inbound-unmatched').onclick = () => exportUnmatchedToCsv(unmatched, '药品入库未建档药品清单.csv');
  } else {
    closeManualWorkspace(unmatchedBox);
    unmatchedBox.classList.add('hidden');
    updateVoucherManualButton('btn-rerun-inbound-manual', state.inboundManualMappings, false);
  }
}

// ----------------------------------------------------
// 2.5 TAB: 外账入库凭证生成 (迁账导入模板)
// ----------------------------------------------------
function initTabExtInbound() {
  const inInbound = document.getElementById('ext-inbound-file');
  const inTemplate = document.getElementById('ext-template-file');
  const inOutput = document.getElementById('ext-output-file');
  const inDate = document.getElementById('ext-inbound-date');
  const inNo = document.getElementById('ext-inbound-no');
  const btnRun = document.getElementById('btn-run-ext-inbound');

  bindFilePicker('btn-browse-ext-inbound', inInbound, '选择药品入库单 (西药或中药)');
  bindFilePicker('btn-browse-ext-template', inTemplate, '选择外账表格迁账参考模板');

  setupDropzone(inInbound, inInbound.closest('.file-input-wrapper'));
  setupDropzone(inTemplate, inTemplate.closest('.file-input-wrapper'));

  btnRun.onclick = async () => {
    const inbound = inInbound.value.trim();
    const template = inTemplate.value.trim();
    const output = inOutput.value.trim() || '表格迁账参考模板_西药入库_已生成.xlsx';
    const dateVal = inDate.value.trim() || null;
    const voucherNo = inNo.value.trim() || '1';

    if (!inbound) return showToast('请指定药品入库单文件', 'warning');
    if (!template) return showToast('请指定外账迁账参考模板', 'warning');

    const res = await invokeWithLoading('execute_external_inbound_voucher', {
      inbound,
      template,
      output,
      date: dateVal,
      voucherNo,
      config: null
    }, {
      loadingMessage: '正在解析入库单、匹配外账5位存货与供应商编码并生成凭证...',
      errorPrefix: '生成外账凭证失败',
      errorDuration: 6000
    });
    if (!res) return;

    state.lastOutputs.extInbound = res.output_file;
    state.extInboundData = res;
    state.extInboundFilter = 'ALL';
    renderExtInboundResult(res);
    showToast('外账入库凭证生成成功！', 'success');
  };

  const btnOpen = document.getElementById('btn-open-ext-file');
  if (btnOpen) {
    btnOpen.onclick = () => {
      openSystemPath(state.lastOutputs.extInbound);
    };
  }

  const btnShow = document.getElementById('btn-show-ext-folder');
  if (btnShow) {
    btnShow.onclick = () => {
      showInSystemFolder(state.lastOutputs.extInbound);
    };
  }

  // 预览过滤按钮
  const filterPills = document.querySelectorAll('#ext-preview-filters .filter-pill');
  filterPills.forEach(pill => {
    pill.onclick = () => {
      filterPills.forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      state.extInboundFilter = pill.dataset.filter;
      renderExtPreviewTable();
    };
  });
}

function renderExtInboundResult(data) {
  const box = document.getElementById('ext-inbound-result');
  if (!box) return;
  box.classList.remove('hidden');

  // 指标卡片
  document.getElementById('stat-ext-items').textContent = `${data.total_items} 药 / ${data.supplier_count} 供`;
  document.getElementById('stat-ext-entries').textContent = `${data.total_entries} 行`;
  document.getElementById('stat-ext-amt').textContent = `¥ ${formatMoney(data.total_debit)}`;

  const balanceCard = document.getElementById('card-ext-balance');
  const balanceVal = document.getElementById('stat-ext-balance');
  if (data.is_balanced) {
    balanceCard.className = 'stat-card highlight';
    balanceVal.innerHTML = '<span style="color:#10b981;">✅ 借贷平衡 (¥0.00)</span>';
  } else {
    balanceCard.className = 'stat-card warning';
    balanceVal.innerHTML = `<span style="color:#ef4444;">❌ 差额: ¥${formatMoney(data.diff)}</span>`;
  }

  const unmatchedCard = document.getElementById('card-ext-unmatched');
  const unmatchedVal = document.getElementById('stat-ext-unmatched');
  const unmatchedSupplierList = data.unmatched_suppliers || [];
  const unmatchedSupplierCount = data.unmatched_supplier_count ?? unmatchedSupplierList.length;
  if (data.unmatched_count === 0 && unmatchedSupplierCount === 0) {
    unmatchedCard.className = 'stat-card';
    unmatchedVal.innerHTML = '<span style="color:#10b981;">0 (全部命中)</span>';
  } else {
    unmatchedCard.className = 'stat-card warning';
    unmatchedVal.innerHTML = `<span style="color:#f59e0b;">存货 ${data.unmatched_count} / 供应商 ${unmatchedSupplierCount}</span>`;
  }

  // 未匹配清单
  const unmatchedBox = document.getElementById('ext-unmatched-box');
  const tbodyUnmatched = document.getElementById('ext-unmatched-table').querySelector('tbody');
  tbodyUnmatched.innerHTML = '';
  const unmatchedList = data.unmatched_drugs || [];
  if (unmatchedList.length > 0) {
    unmatchedBox.classList.remove('hidden');
    unmatchedList.forEach(d => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-family: monospace;">${d.row_index}</td>
        <td style="font-weight: 600; color: #fff;">${escapeHtml(d.name)}</td>
        <td>${escapeHtml(d.spec || '-')}</td>
        <td>${escapeHtml(d.factory || '-')}</td>
        <td style="font-family: monospace;">${d.qty}</td>
        <td style="font-family: monospace; color: #fbbf24;">¥ ${formatMoney(d.amount)}</td>
        <td>${escapeHtml(d.supplier || '-')}</td>
      `;
      tbodyUnmatched.appendChild(tr);
    });
  } else {
    unmatchedBox.classList.add('hidden');
  }

  // 未匹配供应商清单：供应商编码为空会影响 2202 贷方辅助核算，必须单独提示。
  const unmatchedSupplierBox = document.getElementById('ext-unmatched-supplier-box');
  const supplierTbody = document.getElementById('ext-unmatched-supplier-table')?.querySelector('tbody');
  if (supplierTbody) supplierTbody.innerHTML = '';
  if (unmatchedSupplierBox && supplierTbody && unmatchedSupplierList.length > 0) {
    unmatchedSupplierBox.classList.remove('hidden');
    unmatchedSupplierList.forEach(s => {
      const tr = document.createElement('tr');
      tr.innerHTML =
        '<td style="font-weight: 600; color: #fff;">' + escapeHtml(s.supplier || '-') + '</td>' +
        '<td style="font-family: monospace;">' + (s.item_count ?? 0) + '</td>' +
        '<td style="font-family: monospace; color: #fbbf24;">¥ ' + formatMoney(s.amount) + '</td>';
      supplierTbody.appendChild(tr);
    });
  } else if (unmatchedSupplierBox) {
    unmatchedSupplierBox.classList.add('hidden');
  }

  // 计数提示
  const rows = data.voucher_rows || [];
  const debitRows = rows.filter(r => !r.is_credit);
  const creditRows = rows.filter(r => r.is_credit);
  const unmatchedRows = rows.filter(r => r.is_unmatched);

  document.getElementById('cnt-ext-all').textContent = rows.length;
  document.getElementById('cnt-ext-debit').textContent = debitRows.length;
  document.getElementById('cnt-ext-credit').textContent = creditRows.length;
  document.getElementById('cnt-ext-unmatched').textContent = unmatchedRows.length;
  document.getElementById('ext-preview-count-tip').textContent = `共 ${rows.length} 行凭证分录 (已写入 ${escapeHtml(data.output_file)})`;

  renderExtPreviewTable();
}

function renderExtPreviewTable() {
  if (!state.extInboundData || !state.extInboundData.voucher_rows) return;
  const tbody = document.getElementById('ext-preview-table').querySelector('tbody');
  tbody.innerHTML = '';

  const filter = state.extInboundFilter || 'ALL';
  const rows = state.extInboundData.voucher_rows.filter(r => {
    if (filter === 'DEBIT') return !r.is_credit;
    if (filter === 'CREDIT') return r.is_credit;
    if (filter === 'UNMATCHED') return r.is_unmatched;
    return true;
  });

  rows.forEach(r => {
    const tr = document.createElement('tr');
    if (r.is_credit) tr.className = 'row-credit';
    else if (r.is_unmatched) tr.className = 'row-unmatched';

    const statusBadge = r.is_unmatched
      ? '<span class="badge badge-warning">待补录</span>'
      : (r.is_credit ? '<span class="badge badge-success">贷方汇总</span>' : '<span class="badge" style="background:rgba(255,255,255,0.06);color:#94a3b8;">借方明细</span>');

    tr.innerHTML = `
      <td style="font-family: monospace; color: #64748b;">${r.row_no}</td>
      <td style="font-family: monospace;">${escapeHtml(r.date)}</td>
      <td style="text-align: center;">${escapeHtml(r.voucher_type)}</td>
      <td style="text-align: center; font-family: monospace;">${escapeHtml(r.voucher_no)}</td>
      <td>${escapeHtml(r.summary)}</td>
      <td style="font-family: monospace; color: ${r.is_credit ? '#34d399' : '#38bdf8'}; font-weight: 600;">${escapeHtml(r.subject_code)}</td>
      <td>${escapeHtml(r.subject_name)}</td>
      <td style="text-align: right; font-family: monospace; color: ${r.debit_amount ? '#38bdf8' : ''}; font-weight: ${r.debit_amount ? '600' : 'normal'};">
        ${r.debit_amount !== null ? formatMoney(r.debit_amount) : ''}
      </td>
      <td style="text-align: right; font-family: monospace; color: ${r.credit_amount ? '#34d399' : ''}; font-weight: ${r.credit_amount ? '700' : 'normal'};">
        ${r.credit_amount !== null ? formatMoney(r.credit_amount) : ''}
      </td>
      <td style="text-align: right; font-family: monospace;">${r.qty !== null ? r.qty : ''}</td>
      <td style="font-family: monospace; font-weight: 600; color: ${r.aux_code ? '#38bdf8' : '#fbbf24'};">
        ${r.aux_code ? escapeHtml(r.aux_code) : '⚠️ 未匹配'}
      </td>
      <td style="color: ${r.is_credit ? '#a7f3d0' : '#f8fafc'}; font-weight: ${r.is_credit ? '600' : 'normal'};">
        ${escapeHtml(r.aux_name)}
      </td>
      <td style="text-align: center;">${statusBadge}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ----------------------------------------------------
// 5.2 TAB: 外账出库凭证生成 (销售成本结转)
// ----------------------------------------------------
function initTabExtOutbound() {
  const inSales = document.getElementById('ext-outbound-sales-file');
  const inTemplate = document.getElementById('ext-outbound-template-file');
  const inOutput = document.getElementById('ext-outbound-output-file');
  const inDate = document.getElementById('ext-outbound-date');
  const inNo = document.getElementById('ext-outbound-no');
  const btnRun = document.getElementById('btn-run-ext-outbound');

  bindFilePicker('btn-browse-ext-outbound-sales', inSales, '选择销售明细汇总表');
  bindFilePicker('btn-browse-ext-outbound-template', inTemplate, '选择外账迁账参考模板');

  if (inSales) setupDropzone(inSales, inSales.closest('.file-input-wrapper'));
  if (inTemplate) setupDropzone(inTemplate, inTemplate.closest('.file-input-wrapper'));

  // 默认日期设为当前月末最后一天
  if (inDate && !inDate.value) {
    const now = new Date();
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    inDate.value = lastDay.toISOString().split('T')[0];
  }

  btnRun.onclick = async () => {
    const sales = inSales.value.trim();
    const template = inTemplate.value.trim();
    const output = inOutput.value.trim() || '表格迁账参考模板_西药出库_已生成.xlsx';
    const dateVal = inDate.value.trim() || null;
    const voucherNo = inNo.value.trim() || '1';

    if (!sales) return showToast('请指定销售明细汇总表文件', 'warning');
    if (!template) return showToast('请指定外账迁账参考模板', 'warning');

    const res = await invokeWithLoading('execute_external_outbound_voucher', {
      sales,
      template,
      output,
      date: dateVal,
      voucherNo,
      config: null
    }, {
      loadingMessage: '正在解析销售汇总、匹配外账5位存货编码并生成成本结转凭证...',
      errorPrefix: '生成外账出库凭证失败',
      errorDuration: 6000
    });
    if (!res) return;

    state.lastOutputs.extOutbound = res.output_file;
    state.extOutboundData = res;
    state.extOutboundFilter = 'ALL';
    renderExtOutboundResult(res);
    showToast('外账出库结转凭证生成成功！', 'success');
  };

  const btnOpen = document.getElementById('btn-open-ext-outbound-file');
  if (btnOpen) {
    btnOpen.onclick = () => {
      openSystemPath(state.lastOutputs.extOutbound);
    };
  }

  const btnShow = document.getElementById('btn-show-ext-outbound-folder');
  if (btnShow) {
    btnShow.onclick = () => {
      showInSystemFolder(state.lastOutputs.extOutbound);
    };
  }

  // 预览过滤按钮
  const filterPills = document.querySelectorAll('#ext-outbound-preview-filters .filter-pill');
  filterPills.forEach(pill => {
    pill.onclick = () => {
      filterPills.forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      state.extOutboundFilter = pill.dataset.filter;
      renderExtOutboundPreviewTable();
    };
  });
}

function renderExtOutboundResult(data) {
  const box = document.getElementById('ext-outbound-result');
  if (!box) return;
  box.classList.remove('hidden');

  // 指标卡片
  document.getElementById('stat-ext-outbound-items').textContent = `${data.total_items} 种销售药品`;
  document.getElementById('stat-ext-outbound-entries').textContent = `${data.total_entries} 行`;
  document.getElementById('stat-ext-outbound-amt').textContent = `¥ ${formatMoney(data.total_debit)}`;

  const balanceCard = document.getElementById('card-ext-outbound-balance');
  const balanceVal = document.getElementById('stat-ext-outbound-balance');
  if (data.is_balanced) {
    balanceCard.className = 'stat-card highlight';
    balanceVal.innerHTML = '<span style="color:#10b981;">✅ 借贷平衡 (¥0.00)</span>';
  } else {
    balanceCard.className = 'stat-card warning';
    balanceVal.innerHTML = `<span style="color:#ef4444;">❌ 差额: ¥${formatMoney(data.diff)}</span>`;
  }

  const unmatchedCard = document.getElementById('card-ext-outbound-unmatched');
  const unmatchedVal = document.getElementById('stat-ext-outbound-unmatched');
  if (data.unmatched_count === 0) {
    unmatchedCard.className = 'stat-card';
    unmatchedVal.innerHTML = '<span style="color:#10b981;">0 (全部命中)</span>';
  } else {
    unmatchedCard.className = 'stat-card warning';
    unmatchedVal.innerHTML = `<span style="color:#f59e0b;">${data.unmatched_count} 笔 (标黄)</span>`;
  }

  // 未匹配清单
  const unmatchedBox = document.getElementById('ext-outbound-unmatched-box');
  const tbodyUnmatched = document.getElementById('ext-outbound-unmatched-table').querySelector('tbody');
  tbodyUnmatched.innerHTML = '';
  const unmatchedList = data.unmatched_drugs || [];
  if (unmatchedList.length > 0) {
    unmatchedBox.classList.remove('hidden');
    unmatchedList.forEach(d => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-family: monospace;">${d.row_index}</td>
        <td style="font-weight: 600; color: #fff;">${escapeHtml(d.name)}</td>
        <td>${escapeHtml(d.spec || '-')}</td>
        <td>${escapeHtml(d.factory || '-')}</td>
        <td style="font-family: monospace;">${d.qty}</td>
        <td style="font-family: monospace; color: #fbbf24;">¥ ${formatMoney(d.amount)}</td>
      `;
      tbodyUnmatched.appendChild(tr);
    });
  } else {
    unmatchedBox.classList.add('hidden');
  }

  // 计数提示
  const rows = data.voucher_rows || [];
  const debitRows = rows.filter(r => !r.is_credit);
  const creditRows = rows.filter(r => r.is_credit);
  const unmatchedRows = rows.filter(r => r.is_unmatched);

  document.getElementById('cnt-ext-outbound-all').textContent = rows.length;
  document.getElementById('cnt-ext-outbound-debit').textContent = debitRows.length;
  document.getElementById('cnt-ext-outbound-credit').textContent = creditRows.length;
  document.getElementById('cnt-ext-outbound-unmatched').textContent = unmatchedRows.length;
  document.getElementById('ext-outbound-preview-count-tip').textContent = `共 ${rows.length} 行凭证分录 (已写入 ${escapeHtml(data.output_file)})`;

  renderExtOutboundPreviewTable();
}

function renderExtOutboundPreviewTable() {
  if (!state.extOutboundData || !state.extOutboundData.voucher_rows) return;
  const tbody = document.getElementById('ext-outbound-preview-table').querySelector('tbody');
  tbody.innerHTML = '';

  const filter = state.extOutboundFilter || 'ALL';
  const rows = state.extOutboundData.voucher_rows.filter(r => {
    if (filter === 'DEBIT') return !r.is_credit;
    if (filter === 'CREDIT') return r.is_credit;
    if (filter === 'UNMATCHED') return r.is_unmatched;
    return true;
  });

  rows.forEach(r => {
    const tr = document.createElement('tr');
    if (!r.is_credit) tr.className = 'row-debit';
    else if (r.is_unmatched) tr.className = 'row-unmatched';

    const statusBadge = r.is_unmatched
      ? '<span class="badge badge-warning">待补录</span>'
      : (!r.is_credit ? '<span class="badge badge-success">借方汇总</span>' : '<span class="badge" style="background:rgba(255,255,255,0.06);color:#94a3b8;">贷方明细</span>');

    tr.innerHTML = `
      <td style="font-family: monospace; color: #64748b;">${r.row_no}</td>
      <td style="font-family: monospace;">${escapeHtml(r.date)}</td>
      <td style="text-align: center;">${escapeHtml(r.voucher_type)}</td>
      <td style="text-align: center; font-family: monospace;">${escapeHtml(r.voucher_no)}</td>
      <td>${escapeHtml(r.summary)}</td>
      <td style="font-family: monospace; color: ${r.is_credit ? '#38bdf8' : '#34d399'}; font-weight: 600;">${escapeHtml(r.subject_code)}</td>
      <td>${escapeHtml(r.subject_name)}</td>
      <td style="text-align: right; font-family: monospace; color: ${r.debit_amount ? '#34d399' : ''}; font-weight: ${r.debit_amount ? '700' : 'normal'};">
        ${r.debit_amount !== null ? formatMoney(r.debit_amount) : ''}
      </td>
      <td style="text-align: right; font-family: monospace; color: ${r.credit_amount ? '#38bdf8' : ''}; font-weight: ${r.credit_amount ? '600' : 'normal'};">
        ${r.credit_amount !== null ? formatMoney(r.credit_amount) : ''}
      </td>
      <td style="text-align: right; font-family: monospace;">${r.qty !== null ? r.qty : ''}</td>
      <td style="font-family: monospace; font-weight: 600; color: ${r.aux_code ? '#38bdf8' : '#fbbf24'};">
        ${r.aux_code ? escapeHtml(r.aux_code) : '⚠️ 未匹配'}
      </td>
      <td style="color: ${r.is_credit ? '#f8fafc' : '#a7f3d0'}; font-weight: ${r.is_credit ? 'normal' : '600'};">
        ${escapeHtml(r.aux_name)}
      </td>
      <td style="text-align: center;">${statusBadge}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ----------------------------------------------------
// 5.3 TAB: 外账结存数智能比对 (辅助余额表 vs 库管实盘)
// ----------------------------------------------------
function initTabExtCompare() {
  const inTemplate = document.getElementById('ext-compare-template-file');
  const inWarehouse = document.getElementById('ext-compare-wh-file');
  const inOutput = document.getElementById('ext-compare-output-file');
  const btnRun = document.getElementById('btn-run-ext-compare');

  bindFilePicker('btn-browse-ext-compare-template', inTemplate, '选择外账迁账参考模板');
  bindFilePicker('btn-browse-ext-compare-wh', inWarehouse, '选择药库在库实盘报表');

  if (inTemplate) setupDropzone(inTemplate, inTemplate.closest('.file-input-wrapper'));
  if (inWarehouse) setupDropzone(inWarehouse, inWarehouse.closest('.file-input-wrapper'));

  btnRun.onclick = async () => {
    const template = inTemplate.value.trim();
    const warehouse = inWarehouse.value.trim();
    const output = inOutput.value.trim() || '外账账实库存核对分析报告.xlsx';

    if (!template) return showToast('请指定外账迁账参考模板', 'warning');
    if (!warehouse) return showToast('请指定药库在库实盘报表', 'warning');

    const res = await invokeWithLoading('execute_external_inventory_audit', {
      template,
      warehouse,
      output,
      config: null
    }, {
      loadingMessage: '正在比对外账辅助余额表与库管在库数据并计算四维差异...',
      errorPrefix: '外账结存数比对失败',
      errorDuration: 6000
    });
    if (!res) return;

    state.lastOutputs.extCompare = res.output_file;
    state.extCompareData = res;
    state.extCompareFilter = 'ALL';
    renderExtCompareResult(res);
    showToast('外账账实库存核对完成！', 'success');
  };

  const btnOpen = document.getElementById('btn-open-ext-compare-file');
  if (btnOpen) {
    btnOpen.onclick = () => {
      openSystemPath(state.lastOutputs.extCompare);
    };
  }

  const btnShow = document.getElementById('btn-show-ext-compare-folder');
  if (btnShow) {
    btnShow.onclick = () => {
      showInSystemFolder(state.lastOutputs.extCompare);
    };
  }

  // 筛选标签
  const filterPills = document.querySelectorAll('#ext-compare-filters .filter-pill');
  filterPills.forEach(pill => {
    pill.onclick = () => {
      filterPills.forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      state.extCompareFilter = pill.dataset.filter;
      renderExtCompareTable();
    };
  });
}

function renderExtCompareResult(data) {
  const box = document.getElementById('ext-compare-result');
  if (!box) return;
  box.classList.remove('hidden');

  // 指标卡片
  document.getElementById('stat-ext-compare-total').textContent = `${data.total_items} 种品规`;
  document.getElementById('stat-ext-compare-equal').textContent = `${data.equal_count} 种`;
  document.getElementById('stat-ext-compare-diff').textContent = `${data.diff_count} 种`;
  document.getElementById('stat-ext-compare-ext-only').textContent = `${data.ext_only_count} 种`;
  document.getElementById('stat-ext-compare-wh-only').textContent = `${data.wh_only_count} 种`;
  document.getElementById('stat-ext-compare-rate').textContent = `${data.match_rate}%`;

  const diffCard = document.getElementById('card-ext-compare-diff');
  if (data.diff_count === 0) {
    diffCard.className = 'stat-card highlight';
  } else {
    diffCard.className = 'stat-card warning';
  }

  // 计数提示
  const records = data.records || [];
  const equalCount = records.filter(r => r.status === '完全吻合').length;
  const diffCount = records.filter(r => r.status === '数量差异').length;
  const extCount = records.filter(r => r.status === '仅外账有').length;
  const whCount = records.filter(r => r.status === '仅库管有').length;

  document.getElementById('cnt-ext-cmp-all').textContent = records.length;
  document.getElementById('cnt-ext-cmp-equal').textContent = equalCount;
  document.getElementById('cnt-ext-cmp-diff').textContent = diffCount;
  document.getElementById('cnt-ext-cmp-ext').textContent = extCount;
  document.getElementById('cnt-ext-cmp-wh').textContent = whCount;
  document.getElementById('ext-compare-count-tip').textContent = `共 ${records.length} 条核对记录 (报告: ${escapeHtml(data.output_file)})`;

  renderExtCompareTable();
}

function renderExtCompareTable() {
  if (!state.extCompareData || !state.extCompareData.records) return;
  const tbody = document.getElementById('ext-compare-table').querySelector('tbody');
  tbody.innerHTML = '';

  const filter = state.extCompareFilter || 'ALL';
  const records = state.extCompareData.records.filter(r => {
    if (filter === 'EQUAL') return r.status === '完全吻合';
    if (filter === 'DIFF') return r.status === '数量差异';
    if (filter === 'EXT_ONLY') return r.status === '仅外账有';
    if (filter === 'WH_ONLY') return r.status === '仅库管有';
    return true;
  });

  records.forEach((r, idx) => {
    const tr = document.createElement('tr');
    let badgeHtml = '';
    if (r.status === '完全吻合') {
      badgeHtml = '<span class="badge badge-success">完全吻合</span>';
    } else if (r.status === '数量差异') {
      tr.className = 'row-warning';
      badgeHtml = '<span class="badge badge-danger">数量差异</span>';
    } else if (r.status === '仅外账有') {
      badgeHtml = '<span class="badge badge-warning">仅外账有</span>';
    } else {
      badgeHtml = '<span class="badge badge-info">仅库管有</span>';
    }

    const diffDisplay = r.diff_qty !== 0 ? (r.diff_qty > 0 ? `+${r.diff_qty}` : `${r.diff_qty}`) : '0';
    const diffColor = r.diff_qty !== 0 ? (r.diff_qty > 0 ? '#fbbf24' : '#f87171') : '#94a3b8';

    tr.innerHTML = `
      <td style="font-family: monospace; color: #64748b;">${idx + 1}</td>
      <td style="font-family: monospace; font-weight: 600; color: #38bdf8;">${escapeHtml(r.aux_code || '-')}</td>
      <td style="font-weight: 600; color: #fff;">${escapeHtml(r.name)}</td>
      <td>${escapeHtml(r.spec || '-')}</td>
      <td>${escapeHtml(r.factory || '-')}</td>
      <td style="text-align: right; font-family: monospace;">${r.ext_qty}</td>
      <td style="text-align: right; font-family: monospace; color: #94a3b8;">¥ ${formatMoney(r.ext_price)}</td>
      <td style="text-align: right; font-family: monospace;">¥ ${formatMoney(r.ext_amount)}</td>
      <td style="text-align: right; font-family: monospace;">${r.wh_qty}</td>
      <td style="text-align: right; font-family: monospace; font-weight: 700; color: ${diffColor};">${diffDisplay}</td>
      <td style="text-align: center;">${badgeHtml}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ----------------------------------------------------
// 6. TAB 4: 厂家与字典配置 (可视化编辑)
// ----------------------------------------------------
async function loadConfigData() {
  try {
    const res = await invoke('get_config', { config: null });
    if (res && res.success && res.data) {
      state.configData = res.data;
      renderConfigPanel(res.data);
    }
  } catch (err) {
    console.error('加载字典配置失败:', err);
  }
}

function renderConfigPanel(config) {
  // 1. 厂家映射表
  const vendorTbody = document.getElementById('vendor-mapping-table').querySelector('tbody');
  vendorTbody.innerHTML = '';

  const vendorMap = config.factory_abbreviations || config.vendor_abbr_map || {};
  const entries = Object.entries(vendorMap);
  document.getElementById('vendor-count-pill').textContent = `${entries.length} 家厂家`;

  entries.forEach(([fullName, brief]) => {
    addVendorRowToTable(fullName, brief);
  });

  // 2. 严格厂家后缀药品标签
  renderStrictDrugsTags(config.strict_vendor_suffix_drugs || []);

  // 3. 特殊编码覆盖表
  const overrideTbody = document.getElementById('overrides-table').querySelector('tbody');
  overrideTbody.innerHTML = '';
  const overrides = config.drug_code_overrides || {};
  Object.entries(overrides).forEach(([key, val]) => {
    addOverrideRowToTable(key, val);
  });
}

function addVendorRowToTable(fullName = '', brief = '') {
  const vendorTbody = document.getElementById('vendor-mapping-table').querySelector('tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" class="form-input input-sm vendor-fullname" value="${escapeHtml(fullName)}" placeholder="单据中的厂家全称..." /></td>
    <td><input type="text" class="form-input input-sm vendor-brief" value="${escapeHtml(brief)}" placeholder="总账规范简称..." /></td>
    <td style="text-align: center;">
      <button class="btn btn-danger-outline btn-xs btn-delete-row">删除</button>
    </td>
  `;
  tr.querySelector('.btn-delete-row').onclick = () => tr.remove();
  vendorTbody.appendChild(tr);
}

function renderStrictDrugsTags(drugsList) {
  const container = document.getElementById('strict-drugs-tags');
  container.innerHTML = '';
  drugsList.forEach(drug => {
    const tag = document.createElement('span');
    tag.className = 'drug-tag';
    tag.innerHTML = `
      <span>${escapeHtml(drug)}</span>
      <span class="tag-remove" title="移除此项">×</span>
    `;
    tag.querySelector('.tag-remove').onclick = () => {
      tag.remove();
    };
    container.appendChild(tag);
  });
}

function addOverrideRowToTable(drugName = '', code = '') {
  const tbody = document.getElementById('overrides-table').querySelector('tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" class="form-input input-sm override-key" value="${escapeHtml(drugName)}" placeholder="药品目标名..." /></td>
    <td><input type="text" class="form-input input-sm override-val" value="${escapeHtml(code)}" placeholder="科目编码 (留空置空)" /></td>
    <td style="text-align: center;">
      <button class="btn btn-danger-outline btn-xs btn-delete-override">删除</button>
    </td>
  `;
  tr.querySelector('.btn-delete-override').onclick = () => tr.remove();
  tbody.appendChild(tr);
}

function initTabConfig() {
  document.getElementById('btn-add-vendor-row').onclick = () => {
    addVendorRowToTable('', '');
  };

  // 添加严格药品标签
  const drugInput = document.getElementById('strict-drug-input');
  document.getElementById('btn-add-strict-drug').onclick = () => {
    const val = drugInput.value.trim();
    if (!val) return;
    const container = document.getElementById('strict-drugs-tags');
    const existing = Array.from(container.querySelectorAll('.drug-tag span:first-child')).map(el => el.textContent);
    if (existing.includes(val)) {
      showToast('该药品已在严格列表中', 'warning');
      return;
    }
    const tag = document.createElement('span');
    tag.className = 'drug-tag';
    tag.innerHTML = `
      <span>${escapeHtml(val)}</span>
      <span class="tag-remove" title="移除">×</span>
    `;
    tag.querySelector('.tag-remove').onclick = () => tag.remove();
    container.appendChild(tag);
    drugInput.value = '';
  };

  // 添加编码覆盖项
  const overrideKey = document.getElementById('override-key-input');
  const overrideVal = document.getElementById('override-val-input');
  document.getElementById('btn-add-override').onclick = () => {
    const k = overrideKey.value.trim();
    const v = overrideVal.value.trim();
    if (!k) {
      showToast('请输入药品目标全名', 'warning');
      return;
    }
    addOverrideRowToTable(k, v);
    overrideKey.value = '';
    overrideVal.value = '';
  };

  // 保存当前配置
  document.getElementById('btn-save-config').onclick = async () => {
    // 收集厂家映射
    const newVendorMap = {};
    const rows = document.getElementById('vendor-mapping-table').querySelectorAll('tbody tr');
    rows.forEach(tr => {
      const full = tr.querySelector('.vendor-fullname')?.value.trim();
      const brief = tr.querySelector('.vendor-brief')?.value.trim();
      if (full && brief) {
        newVendorMap[full] = brief;
      }
    });

    // 收集严格药品名单
    const tags = document.querySelectorAll('#strict-drugs-tags .drug-tag span:first-child');
    const newStrictDrugs = Array.from(tags).map(t => t.textContent.trim()).filter(Boolean);

    // 收集特殊编码覆盖
    const newOverrides = {};
    const ovRows = document.getElementById('overrides-table').querySelectorAll('tbody tr');
    ovRows.forEach(tr => {
      const k = tr.querySelector('.override-key')?.value.trim();
      const v = tr.querySelector('.override-val')?.value.trim() ?? '';
      if (k) {
        newOverrides[k] = v;
      }
    });

    const newConfig = {
      ...(state.configData || {}),
      factory_abbreviations: newVendorMap,
      vendor_abbr_map: newVendorMap,
      strict_vendor_suffix_drugs: newStrictDrugs,
      drug_code_overrides: newOverrides
    };

    showLoading('正在保存映射规则配置文件...');
    try {
      const res = await invoke('save_config', {
        data: JSON.stringify(newConfig, null, 2),
        config: null
      });

      hideLoading();
      if (res && res.success) {
        state.configData = newConfig;
        showToast('厂家与字典配置已成功保存！即时生效。', 'success');
        document.getElementById('vendor-count-pill').textContent = `${Object.keys(newVendorMap).length} 家厂家`;
      } else {
        showToast(`保存失败: ${res?.error || '未知错误'}`, 'error');
      }
    } catch (err) {
      hideLoading();
      showToast(`保存异常: ${err}`, 'error');
    }
  };

  // 重新加载配置
  document.getElementById('btn-reload-config').onclick = () => {
    loadConfigData();
    showToast('已从磁盘重新加载最新字典配置', 'info');
  };
}

// ----------------------------------------------------
// 7. TAB 4: 账实库存智能核对 (compare_inventory)
// ----------------------------------------------------
function initTabCompare() {
  const inputLedger = document.getElementById('compare-ledger-file');

  const inputWest = document.getElementById('compare-west-file');

  const inputTcm = document.getElementById('compare-tcm-file');

  const inputHc = document.getElementById('compare-hc-file');

  const inputOut = document.getElementById('compare-output-file');

  const btnRun = document.getElementById('btn-run-compare');

  // 文件浏览
  bindFilePicker('btn-browse-compare-ledger', inputLedger, '选择财务系统数量金额总账 (.xlsx)');
  bindFilePicker('btn-browse-compare-west', inputWest, '选择库管系统西药房库存汇总报表 (.xls / .xlsx)');
  bindFilePicker('btn-browse-compare-tcm', inputTcm, '选择库管系统中药房库存汇总报表 (.xls / .xlsx)');
  bindFilePicker('btn-browse-compare-hc', inputHc, '选择库管系统耗材库库存汇总报表 (.xls / .xlsx)');
  bindFilePicker('btn-browse-compare-output', inputOut, '选择或指定对账分析报告保存路径');

  // 拖拽支持
  [inputLedger, inputWest, inputTcm, inputHc].forEach(input => {
    if (input) setupDropzone(input, input.closest('.file-input-row') || input.parentElement);
  });

  // 第一步：智能识别多库与总账映射关系
  if (btnRun) {
    btnRun.onclick = async () => {
      const ledger = inputLedger.value.trim();
      const west = inputWest.value.trim();
      const tcm = inputTcm.value.trim();
      const hc = inputHc.value.trim();

      if (!ledger) {
        showToast('请指定财务系统数量金额总账文件！', 'warning');
        inputLedger.focus();
        return;
      }

      if (!west && !tcm && !hc) {
        showToast('请至少提供一个库管系统的报表文件（西药房 / 中药房 / 耗材库）！', 'warning');
        return;
      }

      resetManualSearchContext('audit', ledger);

      const res = await invokeWithLoading('preview_inventory_audit_mapping', {
        ledger,
        west: west || null,
        tcm: tcm || null,
        hc: hc || null,
        config: null
      }, {
        loadingMessage: '正在运用增强智能算法匹配多库在库品规与财务存货科目...',
        errorPrefix: '智能识别失败'
      });
      if (!res) return;

      state.auditPreview = res;
      renderComparePreviewDashboard(res);
      renderComparePreviewTable();

      const previewCard = document.getElementById('compare-preview-card');
      if (previewCard) {
        previewCard.classList.remove('hidden');
        previewCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }

      showToast(`智能识别完成：共 ${res.total_items} 个品规，已自动匹配 ${res.matched_count} 个（${res.match_rate}%），识别项已默认自动勾选！`, 'success');
    };
  }

  // 第二步：确认匹配关系，生成核对报告
  const btnConfirmRun = document.getElementById('btn-confirm-and-run-compare');
  if (btnConfirmRun) {
    btnConfirmRun.onclick = async () => {
      if (!state.auditPreview || !state.auditPreview.items || state.auditPreview.items.length === 0) {
        showToast('请先执行【第一步：智能识别多库映射关系】！', 'warning');
        return;
      }

      const ledger = inputLedger.value.trim();
      const output = inputOut.value.trim();

      const confirmedItems = state.auditPreview.items.map(it => ({
        id: it.id,
        category: it.category,
        wh_name: it.wh_name,
        wh_spec: it.wh_spec,
        wh_factory: it.wh_factory,
        wh_unit: it.wh_unit,
        wh_qty: it.wh_qty,
        wh_price: it.wh_price,
        wh_amount: it.wh_amount,
        checked: !!it.checked,
        ledger_code: it.ledger_code || ''
      }));

      const checkedCount = confirmedItems.filter(it => it.checked && it.ledger_code).length;
      if (checkedCount === 0) {
        showToast('未勾选任何已匹配的品规！请至少勾选确认一项后再生成报告。', 'warning');
        return;
      }

      const res = await invokeWithLoading('execute_inventory_audit_with_mapping', {
        ledger,
        confirmedItems: confirmedItems,
        output: output || null,
        config: null
      }, {
        loadingMessage: `正在依据确认的 ${checkedCount} 项映射关系执行账实对账，并生成分析报告...`,
        errorPrefix: '核对生成失败'
      });
      if (!res) return;

      state.auditData = res;
      state.lastOutputs.compare = res.output_file;

      renderCompareDashboard(res);
      renderCompareTable();

      const resultCard = document.getElementById('compare-result-card');
      if (resultCard) {
        resultCard.classList.remove('hidden');
        resultCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }

      showToast('账实库存核对分析报告生成成功！', 'success');
    };
  }

  // 预览卡片：库房分类切换
  const previewCatTabs = document.getElementById('preview-cat-tabs');
  if (previewCatTabs) {
    previewCatTabs.addEventListener('click', (e) => {
      const btn = e.target.closest('.cat-pill');
      if (!btn) return;
      previewCatTabs.querySelectorAll('.cat-pill').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.previewCat = btn.dataset.cat || 'ALL';
      renderComparePreviewTable();
    });
  }

  // 预览卡片：仅看未自动匹配
  const chkPreviewUnmatched = document.getElementById('chk-preview-unmatched-only');
  if (chkPreviewUnmatched) {
    chkPreviewUnmatched.addEventListener('change', () => {
      state.previewUnmatchedOnly = chkPreviewUnmatched.checked;
      renderComparePreviewTable();
    });
  }

  // 预览卡片：搜索过滤
  const inputPreviewSearch = document.getElementById('input-preview-search');
  if (inputPreviewSearch) {
    inputPreviewSearch.addEventListener('input', () => {
      state.previewSearch = inputPreviewSearch.value.trim().toLowerCase();
      renderComparePreviewTable();
    });
  }

  // 预览卡片：全选与取消全选
  const btnPreviewSelectAll = document.getElementById('btn-preview-select-all');
  if (btnPreviewSelectAll) {
    btnPreviewSelectAll.onclick = () => {
      setPreviewSelection(true);
    };
  }

  const btnPreviewDeselectAll = document.getElementById('btn-preview-deselect-all');
  if (btnPreviewDeselectAll) {
    btnPreviewDeselectAll.onclick = () => {
      setPreviewSelection(false);
    };
  }

  // 表头全选 Checkbox
  const thChkAll = document.getElementById('th-chk-all');
  if (thChkAll) {
    thChkAll.addEventListener('change', () => {
      setPreviewSelection(thChkAll.checked);
    });
  }

  // 预览表格内的交互事件委托（勾选框；手工科目组合框由渲染函数绑定）
  const previewTable = document.getElementById('compare-preview-table');
  if (previewTable) {
    previewTable.addEventListener('change', (e) => {
      const target = e.target;
      if (target.classList.contains('chk-preview-row')) {
        const id = parseInt(target.dataset.id, 10);
        const item = state.auditPreview?.items?.find(it => it.id === id);
        if (item) {
          item.checked = target.checked;
          updatePreviewStats();
        }
      }
    });
  }

  // 库房分类 Tab 筛选切换 (结果卡片)
  const catTabs = document.getElementById('compare-cat-tabs');
  if (catTabs) {
    catTabs.addEventListener('click', (e) => {
      const btn = e.target.closest('.cat-pill');
      if (!btn) return;
      catTabs.querySelectorAll('.cat-pill').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.currentAuditCat = btn.dataset.cat || 'ALL';
      renderCompareTable();
    });
  }

  // 差异快速过滤复选框 (结果卡片)
  const chkDiff = document.getElementById('chk-compare-diff-only');
  if (chkDiff) {
    chkDiff.addEventListener('change', () => {
      state.auditDiffOnly = chkDiff.checked;
      renderCompareTable();
    });
  }

  // 打开报告与在文件夹中显示
  const btnOpenReport = document.getElementById('btn-open-compare-excel');
  if (btnOpenReport) {
    btnOpenReport.onclick = () => {
      if (state.lastOutputs.compare) {
        openSystemPath(state.lastOutputs.compare);
      } else {
        showToast('尚未生成核对分析报告', 'warning');
      }
    };
  }

  const btnShowFolder = document.getElementById('btn-show-compare-folder');
  if (btnShowFolder) {
    btnShowFolder.onclick = () => {
      if (state.lastOutputs.compare) {
        showInSystemFolder(state.lastOutputs.compare);
      } else {
        showToast('尚未生成核对分析报告', 'warning');
      }
    };
  }
}

// 批量修改当前筛选视图下的勾选状态
function setPreviewSelection(checked) {
  if (!state.auditPreview || !state.auditPreview.items) return;
  const filtered = getFilteredPreviewItems();
  filtered.forEach(it => {
    // 只有已匹配或者已选总账科目的项才允许勾选
    if (checked) {
      if (it.ledger_code) {
        it.checked = true;
      }
    } else {
      it.checked = false;
    }
  });
  renderComparePreviewTable();
  updatePreviewStats();
}

// 获取经过分类、未匹配、搜索过滤的预览项目
function getFilteredPreviewItems() {
  if (!state.auditPreview || !state.auditPreview.items) return [];
  let list = state.auditPreview.items;

  if (state.previewCat && state.previewCat !== 'ALL') {
    list = list.filter(it => it.category === state.previewCat);
  }

  if (state.previewUnmatchedOnly) {
    list = list.filter(it => !it.matched || !it.ledger_code);
  }

  if (state.previewSearch) {
    const kw = state.previewSearch;
    list = list.filter(it => {
      return (it.wh_name && it.wh_name.toLowerCase().includes(kw)) ||
             (it.wh_factory && it.wh_factory.toLowerCase().includes(kw)) ||
             (it.ledger_name && it.ledger_name.toLowerCase().includes(kw)) ||
             (it.ledger_code && it.ledger_code.toLowerCase().includes(kw));
    });
  }

  return list;
}

// 动态刷新预览界面的匹配统计数值
function updatePreviewStats() {
  if (!state.auditPreview || !state.auditPreview.items) return;
  const items = state.auditPreview.items;
  const total = items.length;
  const matched = items.filter(it => it.ledger_code && it.ledger_code.length > 0).length;
  const unmatched = total - matched;
  const rate = total > 0 ? ((matched / total) * 100).toFixed(1) : '0.0';

  const totalEl = document.getElementById('preview-total-count');
  const matchedEl = document.getElementById('preview-matched-count');
  const rateEl = document.getElementById('preview-rate-text');
  const unmatchedEl = document.getElementById('preview-unmatched-count');

  if (totalEl) totalEl.textContent = total.toLocaleString();
  if (matchedEl) matchedEl.textContent = matched.toLocaleString();
  if (rateEl) rateEl.textContent = `${rate}%`;
  if (unmatchedEl) unmatchedEl.textContent = unmatched.toLocaleString();
}

// 渲染预览仪表盘
function renderComparePreviewDashboard(res) {
  updatePreviewStats();
}

// 匹配方法样式徽标生成
function getMethodBadgeHtml(method) {
  if (!method) return '<span class="badge-method method-none">未匹配</span>';
  if (method.includes('精确品名+厂家')) {
    return `<span class="badge-method method-exact">⚡ ${escapeHtml(method)}</span>`;
  }
  if (method.includes('别名') || method.includes('厂商简称') || method.includes('反向提取')) {
    return `<span class="badge-method method-alias">🏷️ ${escapeHtml(method)}</span>`;
  }
  if (method.includes('纯品名')) {
    return `<span class="badge-method method-pure">🔍 ${escapeHtml(method)}</span>`;
  }
  if (method.includes('包含')) {
    return `<span class="badge-method method-sub">💡 ${escapeHtml(method)}</span>`;
  }
  if (method.includes('炮制')) {
    return `<span class="badge-method method-tcm">🌿 ${escapeHtml(method)}</span>`;
  }
  if (method.includes('人工')) {
    return `<span class="badge-method method-manual">✏️ ${escapeHtml(method)}</span>`;
  }
  return `<span class="badge-method method-none">${escapeHtml(method)}</span>`;
}

function renderAuditCandidatePicker(item) {
  const context = state.manualSearch.audit;
  const itemId = String(item.id);
  const defaultCandidates = Array.isArray(item.candidates) ? item.candidates : [];
  if (!Object.prototype.hasOwnProperty.call(context.defaults, itemId)) {
    context.defaults[itemId] = defaultCandidates;
  }

  const selectedCode = item.ledger_code || '';
  const selectedCandidate = context.selectedCandidates[itemId]
    || defaultCandidates.find(candidate => candidate.code === selectedCode);
  if (selectedCode && selectedCandidate && !context.selectedCandidates[itemId]) {
    context.selectedCandidates[itemId] = selectedCandidate;
  }
  const selectedLabel = context.selectedLabels[itemId]
    || (selectedCandidate
      ? candidateLabel(selectedCandidate)
      : (item.ledger_name ? `${item.ledger_name} (${selectedCode})` : selectedCode));
  const searchValue = context.queries[itemId] || (selectedCode ? selectedLabel : '');
  const optionsId = `audit-candidate-options-${itemId}`;

  return `
    <div class="manual-candidate-picker" data-id="${escapeHtml(item.id)}">
      <input
        class="candidate-search-input"
        data-id="${escapeHtml(item.id)}"
        type="search"
        value="${escapeHtml(searchValue)}"
        placeholder="模糊搜索品名/规格/编码"
        title="输入品名、规格或科目编码进行模糊搜索"
        role="combobox"
        aria-autocomplete="list"
        aria-controls="${escapeHtml(optionsId)}"
        aria-expanded="false"
        autocomplete="off"
      />
      <div class="candidate-options hidden" id="${escapeHtml(optionsId)}" role="listbox"></div>
    </div>
  `;
}

function handleAuditCandidateSelection(itemId, selectedCode, candidate) {
  const item = state.auditPreview?.items?.find(it => String(it.id) === String(itemId));
  if (!item) return;

  if (!selectedCode) {
    item.ledger_code = '';
    item.ledger_name = '';
    item.ledger_qty = 0;
    item.ledger_amount = 0;
    item.matched = false;
    item.checked = false;
    item.match_method = '未自动匹配';
  } else if (candidate) {
    item.ledger_code = candidate.code;
    item.ledger_name = candidate.name;
    item.ledger_qty = candidate.qty;
    item.ledger_amount = candidate.amount;
    item.matched = true;
    item.checked = true;
    item.match_method = '人工指定';
  }

  renderComparePreviewTable();
  updatePreviewStats();
}

// 渲染第一步映射对照预览表格
function renderComparePreviewTable() {
  const tbody = document.getElementById('compare-preview-table')?.querySelector('tbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const items = getFilteredPreviewItems();

  if (items.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="12" style="text-align: center; color: var(--text-muted); padding: 24px;">暂无可显示的品规（当前筛选条件下无记录）</td>`;
    tbody.appendChild(tr);
    return;
  }

  items.forEach(it => {
    const tr = document.createElement('tr');
    if (!it.matched || !it.ledger_code) {
      tr.className = 'row-unmatched';
    }

    // 手动调整：支持在当前候选之外按品名、规格或科目编码模糊检索总账。
    const candidatePickerHtml = renderAuditCandidatePicker(it);

    const ledgerCodeText = it.ledger_code
      ? `<span style="font-family: monospace; color: #38bdf8; font-weight: 600;">${escapeHtml(it.ledger_code)}</span>`
      : `<span style="color: var(--text-muted);">-</span>`;

    const ledgerNameText = it.ledger_name
      ? `<span style="color: #e2e8f0; font-size: 12px;">${escapeHtml(it.ledger_name)}</span>`
      : `<span style="color: #f59e0b; font-size: 12px;">（待确认/未匹配）</span>`;

    const ledgerQtyText = it.ledger_code
      ? (it.ledger_qty ?? 0).toLocaleString()
      : '-';

    const disabledAttr = !it.ledger_code ? 'disabled' : '';

    tr.innerHTML = `
      <td style="text-align: center;">
        <input type="checkbox" class="chk-preview-row" data-id="${it.id}" ${it.checked ? 'checked' : ''} ${disabledAttr} />
      </td>
      <td><span class="cat-pill" style="padding: 2px 6px; font-size: 11px;">${escapeHtml(it.category)}</span></td>
      <td style="font-weight: 600; color: #fff;">${escapeHtml(it.wh_name)}</td>
      <td style="font-size: 12px; color: var(--text-muted);">${escapeHtml(it.wh_spec || '-')}</td>
      <td style="font-size: 12px; color: var(--text-muted);">${escapeHtml(it.wh_factory || '-')}</td>
      <td style="text-align: right; font-family: monospace; color: #e2e8f0;">${(it.wh_qty ?? 0).toLocaleString()}</td>
      <td style="text-align: center; color: var(--text-muted);">➔</td>
      <td>${ledgerCodeText}</td>
      <td>${ledgerNameText}</td>
      <td style="text-align: right; font-family: monospace; color: #a5f3fc;">${ledgerQtyText}</td>
      <td>${getMethodBadgeHtml(it.match_method)}</td>
      <td style="text-align: center;">${candidatePickerHtml}</td>
    `;
    tbody.appendChild(tr);
  });

  bindManualCandidatePickers(
    document.getElementById('compare-preview-table'),
    'audit',
    itemId => state.auditPreview?.items?.find(item => String(item.id) === String(itemId))?.ledger_code || '',
    itemId => state.auditPreview?.items?.find(item => String(item.id) === String(itemId))?.category || '',
    handleAuditCandidateSelection
  );
}

// 渲染核对结果仪表盘卡片群
function renderCompareDashboard(data) {
  const card = document.getElementById('compare-result-card');
  if (!card) return;
  card.classList.remove('hidden');

  document.getElementById('compare-output-path-text').textContent = data.output_file || '-';

  const overall = data.overall || {};
  document.getElementById('stat-compare-rate').textContent = `${overall.match_rate ?? 0.0}%`;
  document.getElementById('stat-compare-total').textContent = (overall.total_items ?? 0).toLocaleString();
  document.getElementById('stat-compare-equal').textContent = (overall.equal_count ?? 0).toLocaleString();
  document.getElementById('stat-compare-diff').textContent = (overall.diff_count ?? 0).toLocaleString();
  document.getElementById('stat-compare-ledger-only').textContent = (overall.ledger_only_count ?? 0).toLocaleString();
  document.getElementById('stat-compare-wh-only').textContent = (overall.wh_only_count ?? 0).toLocaleString();
  document.getElementById('stat-compare-diff-amt').textContent = formatMoney(overall.total_diff_amt ?? 0);
}

// 渲染核对交互式明细表格
function renderCompareTable() {
  const data = state.auditData;
  if (!data || !data.categories) return;

  const tbody = document.getElementById('compare-details-table')?.querySelector('tbody');
  const counterEl = document.getElementById('compare-record-counter');
  if (!tbody) return;

  tbody.innerHTML = '';

  // 1. 过滤当前选中的库别
  let selectedCategories = data.categories;
  if (state.currentAuditCat && state.currentAuditCat !== 'ALL') {
    selectedCategories = data.categories.filter(c => c.category === state.currentAuditCat);
  }

  let allRecordsInCat = [];
  selectedCategories.forEach(cat => {
    allRecordsInCat = allRecordsInCat.concat(cat.records || []);
  });

  const totalInCat = allRecordsInCat.length;

  // 2. 差异过滤
  let displayRecords = allRecordsInCat;
  if (state.auditDiffOnly) {
    displayRecords = allRecordsInCat.filter(r => r.status !== 'EQUAL');
  }

  if (counterEl) {
    counterEl.textContent = `显示 ${displayRecords.length} / ${totalInCat} 条品规`;
  }

  if (displayRecords.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="13" style="text-align: center; color: var(--text-muted); padding: 24px;">暂无可显示的记录（${state.auditDiffOnly ? '该分类下无差异品规，全部账实吻合' : '无数据'}）</td>`;
    tbody.appendChild(tr);
    return;
  }

  displayRecords.forEach(r => {
    const tr = document.createElement('tr');

    let statusTag = '';
    let rowClass = '';

    if (r.status === 'EQUAL') {
      statusTag = `<span class="status-tag status-equal">${escapeHtml(r.status_desc || '数量和金额均吻合')}</span>`;
      rowClass = 'row-equal';
    } else if (r.status === 'DIFF' || r.status === 'DIFF_QTY') {
      statusTag = `<span class="status-tag status-diff">${escapeHtml(r.status_desc || '存在数量/金额差异')}</span>`;
      rowClass = 'row-diff';
    } else if (r.status === 'LEDGER_ONLY') {
      statusTag = '<span class="status-tag status-ledger-only">仅财务有结存</span>';
      rowClass = 'row-ledger-only';
    } else if (r.status === 'WH_ONLY') {
      statusTag = '<span class="status-tag status-wh-only">仅库管有在库</span>';
      rowClass = 'row-wh-only';
    }

    tr.className = rowClass;

    const diffQtyFormatted = r.diff_qty > 0 ? `+${r.diff_qty}` : `${r.diff_qty}`;
    const diffQtyColor = r.diff_qty !== 0 ? (r.diff_qty > 0 ? '#f87171' : '#38bdf8') : 'inherit';
    const diffAmtColor = r.diff_amt !== 0 ? (r.diff_amt > 0 ? '#f87171' : '#38bdf8') : 'inherit';

    tr.innerHTML = `
      <td><span class="cat-pill" style="padding: 2px 8px; font-size: 11px;">${escapeHtml(r.category)}</span></td>
      <td>${statusTag}</td>
      <td style="font-weight: 600; color: #fff;">${escapeHtml(r.name)}</td>
      <td>${escapeHtml(r.spec || '-')}</td>
      <td style="color: var(--text-muted); font-size: 12px;">${escapeHtml(r.factory || '-')}</td>
      <td style="text-align: center;">${escapeHtml(r.unit || '-')}</td>
      <td style="font-family: monospace; color: #38bdf8;">${escapeHtml(r.ledger_code || '-')}</td>
      <td style="text-align: right; font-family: monospace;">${r.ledger_qty?.toLocaleString() ?? 0}</td>
      <td style="text-align: right; font-family: monospace;">${r.wh_qty?.toLocaleString() ?? 0}</td>
      <td style="text-align: right; font-family: monospace; font-weight: 700; color: ${diffQtyColor};">${escapeHtml(diffQtyFormatted)}</td>
      <td style="text-align: right; font-family: monospace;">${formatMoney(r.ledger_amt)}</td>
      <td style="text-align: right; font-family: monospace;">${formatMoney(r.wh_amt)}</td>
      <td style="text-align: right; font-family: monospace; font-weight: 700; color: ${diffAmtColor};">${formatMoney(r.diff_amt)}</td>
    `;
    tbody.appendChild(tr);
  });
}

// ----------------------------------------------------
// 人工确认工作区：内嵌 / 独立操作区 / 全屏
// ----------------------------------------------------
const manualWorkspaceState = {
  activePanel: null
};

function updateManualWorkspaceControls(panel) {
  if (!panel) return;

  const expanded = panel.classList.contains('manual-workspace-expanded');
  const fullscreen = panel.classList.contains('manual-workspace-fullscreen');

  panel.querySelectorAll('[data-manual-workspace-action]').forEach(button => {
    const action = button.dataset.manualWorkspaceAction;
    if (action === 'expand') {
      button.classList.toggle('hidden', expanded);
    } else if (action === 'fullscreen') {
      button.classList.toggle('hidden', !expanded);
      button.textContent = fullscreen ? '⤓ 窗口化' : '⤢ 全屏';
      button.title = fullscreen ? '恢复为带边距的独立操作区' : '让独立操作区占满应用窗口';
    } else if (action === 'restore') {
      button.classList.toggle('hidden', !expanded);
    }
  });
}

function openManualWorkspace(panel) {
  if (!panel) return;

  const currentPanel = document.querySelector('.manual-workspace-expanded');
  if (currentPanel && currentPanel !== panel) {
    closeManualWorkspace(currentPanel);
  }

  panel.classList.add('manual-workspace-expanded');
  panel.classList.remove('manual-workspace-fullscreen');
  document.body.classList.add('manual-workspace-open');
  manualWorkspaceState.activePanel = panel;
  updateManualWorkspaceControls(panel);
}

function toggleManualWorkspaceFullscreen(panel) {
  if (!panel || !panel.classList.contains('manual-workspace-expanded')) return;

  panel.classList.toggle('manual-workspace-fullscreen');
  updateManualWorkspaceControls(panel);
}

function closeManualWorkspace(panel) {
  if (!panel) return;

  panel.classList.remove('manual-workspace-expanded', 'manual-workspace-fullscreen');
  if (manualWorkspaceState.activePanel === panel) {
    manualWorkspaceState.activePanel = null;
  }

  if (!document.querySelector('.manual-workspace-expanded')) {
    document.body.classList.remove('manual-workspace-open');
  }
  updateManualWorkspaceControls(panel);
}

function initManualWorkspaceControls() {
  document.querySelectorAll('[data-manual-workspace]').forEach(panel => {
    panel.querySelectorAll('[data-manual-workspace-action]').forEach(button => {
      button.addEventListener('click', () => {
        const action = button.dataset.manualWorkspaceAction;
        if (action === 'expand') {
          openManualWorkspace(panel);
        } else if (action === 'fullscreen') {
          toggleManualWorkspaceFullscreen(panel);
        } else if (action === 'restore') {
          closeManualWorkspace(panel);
        }
      });
    });
    updateManualWorkspaceControls(panel);
  });

  window.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape') return;
    if (document.querySelector('.modal-overlay:not(.hidden)')) return;

    const panel = document.querySelector('.manual-workspace-expanded');
    if (panel) {
      event.preventDefault();
      closeManualWorkspace(panel);
    }
  });
}

// ----------------------------------------------------
// 8. 全局 Tab 切换与应用初始化
// ----------------------------------------------------
function setupAccountSystemSwitcher() {
  const btnInternal = document.getElementById('btn-switch-internal');
  const btnExternal = document.getElementById('btn-switch-external');
  const groupInternal = document.getElementById('group-nav-internal');
  const groupExternal = document.getElementById('group-nav-external');

  const switchSystem = (sys) => {
    state.currentSystem = sys;
    if (sys === 'internal') {
      if (btnInternal) btnInternal.classList.add('active');
      if (btnExternal) btnExternal.classList.remove('active');
      if (groupInternal) groupInternal.classList.remove('hidden');
      if (groupExternal) groupExternal.classList.add('hidden');

      const targetTab = state.activeInternalTab || 'tab-sales';
      const targetTabEl = document.querySelector(`.nav-tab[data-tab="${targetTab}"]`);
      if (targetTabEl) targetTabEl.click();
    } else {
      if (btnExternal) btnExternal.classList.add('active');
      if (btnInternal) btnInternal.classList.remove('active');
      if (groupExternal) groupExternal.classList.remove('hidden');
      if (groupInternal) groupInternal.classList.add('hidden');

      const targetTab = state.activeExternalTab || 'tab-ext-inbound';
      const targetTabEl = document.querySelector(`.nav-tab[data-tab="${targetTab}"]`);
      if (targetTabEl) targetTabEl.click();
    }
  };

  if (btnInternal) btnInternal.onclick = () => switchSystem('internal');
  if (btnExternal) btnExternal.onclick = () => switchSystem('external');
}

function setupTabNavigation() {
  const tabs = document.querySelectorAll('.nav-tab');
  const panels = document.querySelectorAll('.tab-panel');

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const expandedWorkspace = document.querySelector('.manual-workspace-expanded');
      if (expandedWorkspace) closeManualWorkspace(expandedWorkspace);

      const targetId = tab.dataset.tab;
      state.activeTab = targetId;

      if (tab.closest('#group-nav-internal')) {
        state.activeInternalTab = targetId;
      } else if (tab.closest('#group-nav-external')) {
        state.activeExternalTab = targetId;
      }

      tabs.forEach(t => t.classList.remove('active'));
      panels.forEach(p => p.classList.remove('active'));

      tab.classList.add('active');
      const targetPanel = document.getElementById(targetId);
      if (targetPanel) targetPanel.classList.add('active');
    });
  });

  // 顶部一键扫描按钮
  const btnScan = document.getElementById('btn-quick-scan');
  if (btnScan) {
    btnScan.onclick = () => triggerFileScan();
  }
}

// 初始化“业务规则与使用帮助”模态弹窗
function initHelpModal() {
  const modal = document.getElementById('help-modal');
  const btnOpen = document.getElementById('btn-help-modal');
  const btnClose = document.getElementById('btn-close-help');
  const btnOk = document.getElementById('btn-modal-ok');

  if (!modal) return;

  const openModal = () => modal.classList.remove('hidden');
  const closeModal = () => modal.classList.add('hidden');

  if (btnOpen) btnOpen.onclick = openModal;
  if (btnClose) btnClose.onclick = closeModal;
  if (btnOk) btnOk.onclick = closeModal;

  // 点击背景遮罩关闭
  modal.onclick = (e) => {
    if (e.target === modal) closeModal();
  };

  // ESC 按键关闭
  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !modal.classList.contains('hidden')) {
      closeModal();
    }
  });
}

// 启动入口
window.addEventListener('DOMContentLoaded', async () => {
  setupTabNavigation();
  setupAccountSystemSwitcher();
  initHelpModal();
  initTabSales();
  initTabOutbound();
  initTabInbound();
  initTabExtInbound();
  initTabExtOutbound();
  initTabExtCompare();
  initTabCompare();
  initTabConfig();
  initManualWorkspaceControls();

  // 预载配置与自动扫描
  loadConfigData();
  triggerFileScan();
});
