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
  scannedFiles: null,
  configData: null,
  lastOutputs: {
    sales: null,
    outbound: null,
    inbound: null
  }
};

// UI 辅助工具函数：Toast 消息通知
function showToast(message, type = 'info', duration = 3500) {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `<span>${message}</span>`;
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
        showToast(`已载入文件: ${f.name}`, 'info');
      } else {
        showToast('请拖入有效的 Excel 文件 (.xlsx 或 .xls)', 'warning');
      }
    }
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
  // 销售明细输入框
  const salesInput = document.getElementById('sales-input-file');
  if (salesInput && !salesInput.value && data.sales_files.length > 0) {
    salesInput.value = data.sales_files[0].path;
  }

  // 出库凭证输入框
  const obSales = document.getElementById('outbound-sales-file');
  const obLedger = document.getElementById('outbound-ledger-file');
  const obTmpl = document.getElementById('outbound-template-file');
  if (obSales && !obSales.value && data.sales_files.length > 0) {
    // 优先选择带有“已汇总”字样的文件
    const rolled = data.sales_files.find(f => f.name.includes('已汇总')) || data.sales_files[0];
    obSales.value = rolled.path;
  }
  if (obLedger && !obLedger.value && data.ledger_files.length > 0) {
    obLedger.value = data.ledger_files[0].path;
  }
  if (obTmpl && !obTmpl.value && data.template_files.length > 0) {
    obTmpl.value = data.template_files[0].path;
  }

  // 入库凭证输入框
  const inInbound = document.getElementById('inbound-file');
  const inLedger = document.getElementById('inbound-ledger-file');
  const inTmpl = document.getElementById('inbound-template-file');
  if (inInbound && !inInbound.value && data.inbound_files.length > 0) {
    inInbound.value = data.inbound_files[0].path;
  }
  if (inLedger && !inLedger.value && data.ledger_files.length > 0) {
    inLedger.value = data.ledger_files[0].path;
  }
  if (inTmpl && !inTmpl.value && data.template_files.length > 0) {
    inTmpl.value = data.template_files[0].path;
  }
}

function applyScannedFileToActiveTab(item) {
  const name = item.name;
  if (state.activeTab === 'tab-sales') {
    document.getElementById('sales-input-file').value = item.path;
    showToast(`已选定销售表: ${name}`, 'info');
  } else if (state.activeTab === 'tab-outbound') {
    if (name.includes('总账')) {
      document.getElementById('outbound-ledger-file').value = item.path;
      showToast(`已填入总账表: ${name}`, 'info');
    } else if (name.includes('模板')) {
      document.getElementById('outbound-template-file').value = item.path;
      showToast(`已填入凭证模板: ${name}`, 'info');
    } else {
      document.getElementById('outbound-sales-file').value = item.path;
      showToast(`已填入销售汇总表: ${name}`, 'info');
    }
  } else if (state.activeTab === 'tab-inbound') {
    if (name.includes('总账')) {
      document.getElementById('inbound-ledger-file').value = item.path;
      showToast(`已填入总账表: ${name}`, 'info');
    } else if (name.includes('模板')) {
      document.getElementById('inbound-template-file').value = item.path;
      showToast(`已填入凭证模板: ${name}`, 'info');
    } else {
      document.getElementById('inbound-file').value = item.path;
      showToast(`已填入入库单: ${name}`, 'info');
    }
  }
}

// ----------------------------------------------------
// 3. TAB 1: 销售汇总处理
// ----------------------------------------------------
function initTabSales() {
  const btnBrowse = document.getElementById('btn-browse-sales');
  const inputSales = document.getElementById('sales-input-file');
  const btnRun = document.getElementById('btn-run-sales');
  const inputSheet = document.getElementById('sales-sheet-name');
  const inputOutput = document.getElementById('sales-output-name');

  btnBrowse.onclick = async () => {
    const file = await pickExcelFile('选择原始销售明细表');
    if (file) inputSales.value = file;
  };

  setupDropzone(inputSales, inputSales.closest('.file-input-wrapper'));

  btnRun.onclick = async () => {
    const file = inputSales.value.trim();
    if (!file) {
      showToast('请先选择或拖入原始销售明细表', 'warning');
      return;
    }

    const sheetName = inputSheet.value.trim() || null;
    const output = inputOutput.value.trim() || null;

    showLoading('正在对销售明细进行去重、合并及多维度汇总计算...');
    try {
      const res = await invoke('execute_sales_process', {
        file,
        output,
        sheetName
      });

      hideLoading();
      if (!res || !res.success) {
        showToast(`销售汇总失败: ${res?.error || '未知错误'}`, 'error', 5000);
        return;
      }

      state.lastOutputs.sales = res.output_file;
      renderSalesResult(res);
      showToast('销售明细汇总处理完成！', 'success');

      // 智能联动：如果出库凭证面板的销售表为空，自动同步带入新汇总的文件
      const obSales = document.getElementById('outbound-sales-file');
      if (obSales && !obSales.value) {
        obSales.value = res.output_file;
      }
    } catch (err) {
      hideLoading();
      showToast(`执行异常: ${err}`, 'error');
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

  const totals = data.totals || {};
  document.getElementById('stat-sales-orig').textContent = totals.original_count?.toLocaleString() || '-';
  document.getElementById('stat-sales-unique').textContent = totals.unique_count?.toLocaleString() || '-';
  document.getElementById('stat-sales-qty').textContent = totals.total_qty?.toLocaleString() || '-';
  document.getElementById('stat-sales-in-amt').textContent = formatMoney(totals.total_in_amt);
  document.getElementById('stat-sales-retail-amt').textContent = formatMoney(totals.total_retail_amt);
}

// ----------------------------------------------------
// 4. TAB 2: 销售出库凭证生成
// ----------------------------------------------------
function initTabOutbound() {
  const inSales = document.getElementById('outbound-sales-file');
  const inLedger = document.getElementById('outbound-ledger-file');
  const inTemplate = document.getElementById('outbound-template-file');
  const inDate = document.getElementById('outbound-date');
  const chkFallback = document.getElementById('outbound-fallback-price');
  const btnRun = document.getElementById('btn-run-outbound');

  document.getElementById('btn-browse-outbound-sales').onclick = async () => {
    const file = await pickExcelFile('选择销售汇总表');
    if (file) inSales.value = file;
  };
  document.getElementById('btn-browse-outbound-ledger').onclick = async () => {
    const file = await pickExcelFile('选择数量金额总账表');
    if (file) inLedger.value = file;
  };
  document.getElementById('btn-browse-outbound-template').onclick = async () => {
    const file = await pickExcelFile('选择凭证导入模板');
    if (file) inTemplate.value = file;
  };

  setupDropzone(inSales, inSales.closest('.file-input-wrapper'));
  setupDropzone(inLedger, inLedger.closest('.file-input-wrapper'));
  setupDropzone(inTemplate, inTemplate.closest('.file-input-wrapper'));

  btnRun.onclick = async () => {
    const sales = inSales.value.trim();
    const ledger = inLedger.value.trim();
    const template = inTemplate.value.trim();

    if (!sales) return showToast('请指定销售汇总表', 'warning');
    if (!ledger) return showToast('请指定数量金额总账表', 'warning');
    if (!template) return showToast('请指定凭证导入模板', 'warning');

    const dateVal = inDate.value.trim() || null;
    const fallback = chkFallback.checked;

    showLoading('正在匹配总账存货编码与单价，生成销售出库凭证...');
    try {
      const res = await invoke('execute_outbound_voucher', {
        sales,
        ledger,
        template,
        output: null,
        date: dateVal,
        fallbackPrice: fallback,
        config: null
      });

      hideLoading();
      if (!res || !res.success) {
        showToast(`生成出库凭证失败: ${res?.error || '未知错误'}`, 'error', 6000);
        return;
      }

      state.lastOutputs.outbound = res.output_file;
      renderOutboundResult(res);
      showToast('销售出库凭证生成成功！', 'success');
    } catch (err) {
      hideLoading();
      showToast(`执行异常: ${err}`, 'error');
    }
  };

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

function renderOutboundResult(data) {
  const box = document.getElementById('outbound-result');
  box.classList.remove('hidden');

  document.getElementById('stat-outbound-total').textContent = data.total_items?.toLocaleString() || '-';
  document.getElementById('stat-outbound-rate').textContent = `${data.match_rate}%`;
  document.getElementById('stat-outbound-matched').textContent = data.matched_count?.toLocaleString() || '-';
  document.getElementById('stat-outbound-unmatched').textContent = data.unmatched_count?.toLocaleString() || '0';
  document.getElementById('stat-outbound-amt').textContent = formatMoney(data.total_credit_amt);

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
      const drugDisplayName = item.name || item.target_name;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-weight: 600; color: #fff;">${drugDisplayName}</td>
        <td>${item.spec || '-'}</td>
        <td>${item.factory || '-'}</td>
        <td style="font-family: monospace;">${item.qty}</td>
        <td style="font-family: monospace;">${item.in_price !== null ? '¥ ' + item.in_price : '-'}</td>
        <td><span class="diag-tag ${diag.type}">${diag.text}</span></td>
        <td style="text-align: center;">
          <button class="btn btn-outline btn-xs btn-jump-config" title="在字典配置中心为该药品设置编码">
            + 配置编码
          </button>
        </td>
      `;
      tr.querySelector('.btn-jump-config').onclick = () => jumpToConfigWithDrug(drugDisplayName);
      tbody.appendChild(tr);
    });

    // 绑定导出与复制按钮
    document.getElementById('btn-copy-outbound-unmatched').onclick = () => copyUnmatchedToClipboard(unmatched);
    document.getElementById('btn-export-outbound-unmatched').onclick = () => exportUnmatchedToCsv(unmatched, '销售出库未建档药品清单.csv');
  } else {
    unmatchedBox.classList.add('hidden');
  }
}

// ----------------------------------------------------
// 5. TAB 3: 药房入库凭证生成 (西药 / 中药)
// ----------------------------------------------------
function initTabInbound() {
  const inInbound = document.getElementById('inbound-file');
  const inLedger = document.getElementById('inbound-ledger-file');
  const inTemplate = document.getElementById('inbound-template-file');
  const inVoucherNo = document.getElementById('inbound-voucher-no');
  const inDate = document.getElementById('inbound-date');
  const btnRun = document.getElementById('btn-run-inbound');

  document.getElementById('btn-browse-inbound').onclick = async () => {
    const file = await pickExcelFile('选择药品入库单 (西药或中药)');
    if (file) inInbound.value = file;
  };
  document.getElementById('btn-browse-inbound-ledger').onclick = async () => {
    const file = await pickExcelFile('选择数量金额总账表');
    if (file) inLedger.value = file;
  };
  document.getElementById('btn-browse-inbound-template').onclick = async () => {
    const file = await pickExcelFile('选择凭证导入模板');
    if (file) inTemplate.value = file;
  };

  setupDropzone(inInbound, inInbound.closest('.file-input-wrapper'));
  setupDropzone(inLedger, inLedger.closest('.file-input-wrapper'));
  setupDropzone(inTemplate, inTemplate.closest('.file-input-wrapper'));

  btnRun.onclick = async () => {
    const inbound = inInbound.value.trim();
    const ledger = inLedger.value.trim();
    const template = inTemplate.value.trim();

    if (!inbound) return showToast('请指定药品入库单', 'warning');
    if (!ledger) return showToast('请指定数量金额总账表', 'warning');
    if (!template) return showToast('请指定凭证导入模板', 'warning');

    const voucherNo = inVoucherNo.value.trim() || '记';
    const dateVal = inDate.value.trim() || null;

    showLoading('正在解析入库单据、匹配供应商与存货编码并进行借贷平衡审计...');
    try {
      const res = await invoke('execute_inbound_voucher', {
        inbound,
        ledger,
        template,
        output: null,
        date: dateVal,
        voucherNo,
        config: null
      });

      hideLoading();
      if (!res || !res.success) {
        showToast(`生成入库凭证失败: ${res?.error || '未知错误'}`, 'error', 6000);
        return;
      }

      state.lastOutputs.inbound = res.output_file;
      renderInboundResult(res);
      showToast('药房入库凭证生成成功！', 'success');
    } catch (err) {
      hideLoading();
      showToast(`执行异常: ${err}`, 'error');
    }
  };

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
      <td style="font-weight: 600; color: #fff;">${s.supplier}</td>
      <td style="font-family: monospace; color: #38bdf8;">${s.code || '<span style="color:#f87171">未匹配编码</span>'}</td>
      <td>${s.brief || '-'}</td>
      <td style="font-family: monospace;">${s.count} 条</td>
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
      const drugDisplayName = item.target_name || item.name;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${item.supplier || '-'}</td>
        <td style="font-weight: 600; color: #fbbf24;">${drugDisplayName}</td>
        <td>${item.spec || '-'}</td>
        <td style="font-family: monospace;">${item.qty}</td>
        <td style="font-family: monospace;">¥ ${item.in_price}</td>
        <td style="font-family: monospace; font-weight: 600;">${formatMoney(item.in_amt)}</td>
        <td><span class="diag-tag ${diag.type}">${diag.text}</span></td>
        <td style="text-align: center;">
          <button class="btn btn-outline btn-xs btn-jump-config" title="在字典配置中心为该药品设置编码">
            + 配置编码
          </button>
        </td>
      `;
      tr.querySelector('.btn-jump-config').onclick = () => jumpToConfigWithDrug(drugDisplayName);
      tbody.appendChild(tr);
    });

    // 绑定导出与复制按钮
    document.getElementById('btn-copy-inbound-unmatched').onclick = () => copyUnmatchedToClipboard(unmatched);
    document.getElementById('btn-export-inbound-unmatched').onclick = () => exportUnmatchedToCsv(unmatched, '药品入库未建档药品清单.csv');
  } else {
    unmatchedBox.classList.add('hidden');
  }
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
    <td><input type="text" class="form-input input-sm vendor-fullname" value="${fullName}" placeholder="单据中的厂家全称..." /></td>
    <td><input type="text" class="form-input input-sm vendor-brief" value="${brief}" placeholder="总账规范简称..." /></td>
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
      <span>${drug}</span>
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
    <td><input type="text" class="form-input input-sm override-key" value="${drugName}" placeholder="药品目标名..." /></td>
    <td><input type="text" class="form-input input-sm override-val" value="${code}" placeholder="科目编码 (留空置空)" /></td>
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
      <span>${val}</span>
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
// 7. 全局 Tab 切换与应用初始化
// ----------------------------------------------------
function setupTabNavigation() {
  const tabs = document.querySelectorAll('.nav-tab');
  const panels = document.querySelectorAll('.tab-panel');

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const targetId = tab.dataset.tab;
      state.activeTab = targetId;

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
  initHelpModal();
  initTabSales();
  initTabOutbound();
  initTabInbound();
  initTabConfig();

  // 预载配置与自动扫描
  loadConfigData();
  triggerFileScan();
});

