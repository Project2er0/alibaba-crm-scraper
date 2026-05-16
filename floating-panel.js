/**
 * 阿里CRM客户采集 - 浮窗面板逻辑
 * 注入到页面中作为固定浮窗，不会因失焦而关闭
 * 包含：UI注入、采集逻辑、拖拽、关闭/最小化、XLSX导出
 */

// ============================================================
// 全局状态
// ============================================================
let allData = [];
let isCollecting = false;
let currentTabId = null;
let stopRequested = false;
let dedupEnabled = true; // 默认开启去重

const EXCEL_HEADERS = [
  '联系人姓名', '公司名称', '业务员', '客户阶段',
  '跟进状态', '跟进内容', '小记时间',
  '采购意向', '年采购额', '采购品类',
  '国家/地区', '所属客群', '客户来源', '商业类型', '建档时间'
];

const DATA_KEYS = [
  'contactName', 'companyName', 'owner', 'customerStage',
  'followStatus', 'followContent', 'followTime',
  'purchaseIntent', 'annualPurchase', 'purchaseCategory',
  'country', 'customerGroup', 'customerOrigin', 'bizType', 'createTime'
];

// ============================================================
// 数据去重（按联系人+公司名）
// ============================================================
function dedupData(data) {
  const seen = new Set();
  return data.filter(item => {
    const key = `${item.contactName || ''}|||${item.companyName || ''}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// ============================================================
// 注入浮窗 HTML（CSS 由 manifest 自动注入）
// ============================================================
function injectPanelHTML() {
  if (document.getElementById('crm-collector-panel')) return;

  const panel = document.createElement('div');
  panel.id = 'crm-collector-panel';
  panel.innerHTML = `
    <!-- Header（可拖拽） -->
    <div class="crm-header" id="crmHeader">
      <div class="crm-logo">📊</div>
      <div>
        <div class="crm-header-title">阿里CRM客户采集工具</div>
        <div class="crm-header-subtitle">阿里巴巴国际站 · 客户列表数据导出</div>
      </div>
      <button class="crm-minimize-btn" id="crmMinBtn" title="最小化">─</button>
      <button class="crm-close-btn" id="crmCloseBtn" title="关闭面板">✕</button>
    </div>

    <!-- 可滚动内容区 -->
    <div class="crm-body">
      <!-- Status -->
      <div id="crmStatus" class="crm-status info">请在阿里巴巴国际站CRM客户列表页使用</div>

      <!-- Stats -->
      <div class="crm-stats">
        <div class="crm-stats-item">
          <div class="crm-stats-num" id="crmDataCount">0</div>
          <div class="crm-stats-label">已采集条数</div>
        </div>
        <div class="crm-progress-wrapper" style="flex:1; padding: 0 16px;">
          <div class="crm-progress-track">
            <div class="crm-progress-bar" id="crmProgressBar"></div>
          </div>
          <div class="crm-progress-label" id="crmProgressLabel"></div>
        </div>
      </div>

      <!-- Buttons -->
      <div class="crm-btn-group">
        <button id="crmBtnStart" class="crm-btn-primary crm-btn-full">
          🚀 全部采集（自动翻页）
        </button>
        <button id="crmBtnStartCurrent" class="crm-btn-secondary">
          📋 仅采集当前页
        </button>
        <button id="crmBtnStop" class="crm-btn-stop" disabled>
          ⏹ 停止
        </button>
        <button id="crmBtnExport" class="crm-btn-success crm-btn-full" disabled>
          📥 导出 Excel
        </button>
        <button id="crmBtnClear" class="crm-btn-danger crm-btn-full">
          🗑 清空数据
        </button>
      </div>

      <!-- 手动翻页 -->
      <div class="crm-page-nav">
        <button id="crmBtnPrevPage">◀ 上一页</button>
        <button id="crmBtnNextPage">下一页 ▶</button>
      </div>


      <!-- Log -->
      <div class="crm-log-header">
        <span>运行日志</span>
      </div>
      <div class="crm-log-area" id="crmLogArea"></div>

      <!-- Tip -->
      <div class="crm-tip">
        💡 <b>使用说明：</b>先打开阿里巴巴国际站 CRM 客户列表页，再点击采集。
        全部采集会自动翻页，也可手动点击上一页/下一页后采集当前页。
      </div>
    </div>
  `;

  document.body.appendChild(panel);
}

// ============================================================
// 注入唤出按钮（浮窗关闭后显示）
// ============================================================
function injectToggleBtn() {
  if (document.getElementById('crm-collector-toggle')) return;

  const btn = document.createElement('button');
  btn.id = 'crm-collector-toggle';
  btn.title = '打开采集面板';
  btn.textContent = '📊';
  btn.addEventListener('click', () => showPanel());
  document.body.appendChild(btn);
}

// ============================================================
// 显示 / 隐藏面板
// ============================================================
function showPanel() {
  const panel = document.getElementById('crm-collector-panel');
  const toggle = document.getElementById('crm-collector-toggle');
  if (!panel) return;
  panel.classList.remove('crm-panel-hidden', 'crm-minimized');
  if (toggle) toggle.classList.remove('crm-toggle-visible');
}

function hidePanel() {
  const panel = document.getElementById('crm-collector-panel');
  const toggle = document.getElementById('crm-collector-toggle');
  if (!panel) return;

  // 采集进行中时禁止关闭
  if (isCollecting) {
    const status = document.getElementById('crmStatus');
    if (status) {
      status.textContent = '⚠️ 采集进行中，请等待采集完成后再关闭';
      status.className = 'crm-status error';
    }
    return;
  }

  panel.classList.add('crm-panel-hidden');
  if (toggle) toggle.classList.add('crm-toggle-visible');
}

function minimizePanel() {
  const panel = document.getElementById('crm-collector-panel');
  if (!panel) return;
  panel.classList.toggle('crm-minimized');
  const minBtn = document.getElementById('crmMinBtn');
  if (minBtn) {
    minBtn.textContent = panel.classList.contains('crm-minimized') ? '□' : '─';
    minBtn.title = panel.classList.contains('crm-minimized') ? '展开' : '最小化';
  }
}

// ============================================================
// 拖拽逻辑
// ============================================================
function enableDrag() {
  const header = document.getElementById('crmHeader');
  const panel = document.getElementById('crm-collector-panel');
  if (!header || !panel) return;

  let isDragging = false;
  let startX, startY, startLeft, startTop;

  header.addEventListener('mousedown', (e) => {
    // 不拖拽按钮区域
    if (e.target.closest('.crm-close-btn') || e.target.closest('.crm-minimize-btn')) return;

    isDragging = true;
    const rect = panel.getBoundingClientRect();
    startX = e.clientX;
    startY = e.clientY;
    startLeft = rect.left;
    startTop = rect.top;

    e.preventDefault();
  });

  document.addEventListener('mousemove', (e) => {
    if (!isDragging) return;

    const dx = e.clientX - startX;
    const dy = e.clientY - startY;

    let newLeft = startLeft + dx;
    let newTop = startTop + dy;

    // 边界限制
    const pw = panel.offsetWidth;
    const ph = panel.offsetHeight;
    newLeft = Math.max(0, Math.min(newLeft, window.innerWidth - pw));
    newTop = Math.max(0, Math.min(newTop, window.innerHeight - ph));

    panel.style.left = newLeft + 'px';
    panel.style.top = newTop + 'px';
    panel.style.right = 'auto';
  });

  document.addEventListener('mouseup', () => {
    isDragging = false;
  });
}

// ============================================================
// UI 工具函数
// ============================================================
function setStatus(text, type = 'info') {
  const el = document.getElementById('crmStatus');
  if (el) {
    el.textContent = text;
    el.className = `crm-status ${type}`;
  }
}

function setProgress(current, total) {
  const bar = document.getElementById('crmProgressBar');
  const label = document.getElementById('crmProgressLabel');
  const pct = total > 0 ? Math.min(100, Math.round(current / total * 100)) : 0;
  if (bar) bar.style.width = pct + '%';
  if (label) label.textContent = total > 0 ? `${current} / ${total} 页` : '';
}

function updateCount(count) {
  const el = document.getElementById('crmDataCount');
  if (el) el.textContent = count;
}

function setButtonState(collecting) {
  const ids = {
    crmBtnStart: collecting,
    crmBtnStartCurrent: collecting,
    crmBtnStop: !collecting,
    crmBtnExport: allData.length === 0,
    crmBtnPrevPage: collecting,
    crmBtnNextPage: collecting,
  };
  Object.entries(ids).forEach(([id, disabled]) => {
    const el = document.getElementById(id);
    if (el) el.disabled = disabled;
  });

  // 采集进行中，关闭按钮禁用
  const closeBtn = document.getElementById('crmCloseBtn');
  if (closeBtn) closeBtn.disabled = collecting;
}

function log(msg) {
  const el = document.getElementById('crmLogArea');
  if (!el) return;
  const line = document.createElement('div');
  line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}

// ============================================================
// 采集逻辑
// ============================================================
function wait(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function sendToContent(message, timeoutMs = 8000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('操作超时')), timeoutMs);
    try {
      let result;
      switch (message.action) {
        case 'SCRAPE_CURRENT_PAGE':
          result = { success: true, data: scrapeCurrentPage(), pagination: getPaginationInfo() };
          break;
        case 'CLICK_NEXT_PAGE':
          result = { success: message.selector ? clickPageBySelector(message.selector) : clickNextPage() };
          break;
        case 'CLICK_PREV_PAGE':
          result = { success: message.selector ? clickPageBySelector(message.selector) : clickPrevPage() };
          break;
        case 'GET_PAGINATION':
          result = { success: true, pagination: getPaginationInfo() };
          break;
        case 'JUMP_TO_PAGE':
          result = { success: jumpToPage(message.page) };
          break;
        default:
          result = { success: false, error: '未知操作' };
      }
      clearTimeout(timer);
      resolve(result);
    } catch (e) {
      clearTimeout(timer);
      reject(e);
    }
  });
}

async function scrapeSinglePage() {
  await wait(1500);
  const response = await sendToContent({ action: 'SCRAPE_CURRENT_PAGE' });
  if (!response || !response.success) {
    throw new Error(response?.error || '采集失败');
  }
  return response;
}

async function scrapeAllPages() {
  isCollecting = true;
  stopRequested = false;
  setButtonState(true);
  setStatus('正在采集...', 'info');
  log('开始多页采集...');

  try {
  dedupEnabled = true; // 始终去重

  const MAX_SAFE_PAGES = 500; // 安全上限
  let emptyPageCount = 0;     // 连续空页计数
  const EMPTY_STOP_THRESHOLD = 3; // 连续3页为空则停止

  // 先采集当前页，获取分页信息
  let firstResult;
  try {
    firstResult = await scrapeSinglePage();
  } catch (e) {
    setStatus(`采集出错: ${e.message}`, 'error');
    log(`❌ 错误: ${e.message}`);
    setButtonState(false);
    isCollecting = false;
    return;
  }

  const totalPages = firstResult.pagination.totalPages || 1;
  const totalRecords = firstResult.pagination.totalRecords || 0;

  // 第一页为空，直接结束
  if (firstResult.data.length === 0) {
    setStatus('当前页无数据，请检查筛选条件', 'error');
    log('当前页无数据，采集结束');
    setButtonState(false);
    isCollecting = false;
    return;
  }

  // 自动确定采集页范围
  const endPage = Math.min(totalPages, MAX_SAFE_PAGES);
  if (totalRecords > 5000) {
    log(`⚠️ 平台限制：共 ${totalRecords} 条记录，最多显示 5000 条`);
  }
  if (totalPages > MAX_SAFE_PAGES) {
    log(`⚠️ 总页数 ${totalPages} 超过安全上限 ${MAX_SAFE_PAGES}，将采集到第 ${MAX_SAFE_PAGES} 页`);
  }

  allData = allData.concat(firstResult.data);
  updateCount(allData.length);
  log(`第 1 页: 采集到 ${firstResult.data.length} 条`);

  setProgress(1, endPage);

  if (endPage <= 1) {
    log('仅1页，采集完成');
  } else {
    for (let targetPage = 2; targetPage <= endPage; targetPage++) {
      if (stopRequested) {
        log('用户已停止采集');
        break;
      }

      setStatus(`正在采集第 ${targetPage}/${endPage} 页...`, 'info');

      // 尝试翻页
      const clickResult = await sendToContent({ action: 'CLICK_NEXT_PAGE' });
      if (!clickResult || !clickResult.success) {
        log(`第 ${targetPage} 页：无法翻页（已到最后一页或翻页按钮不可用），自动停止`);
        break;
      }

      await wait(2500);

      // 采集当前页
      const result = await scrapeSinglePage();

      // ★ 关键：检测空页 → 自动停止
      if (result.data.length === 0) {
        emptyPageCount++;
        log(`第 ${targetPage} 页: 无数据（连续空页 ${emptyPageCount}/${EMPTY_STOP_THRESHOLD}）`);
        if (emptyPageCount >= EMPTY_STOP_THRESHOLD) {
          log(`连续 ${EMPTY_STOP_THRESHOLD} 页无数据，自动停止采集`);
          break;
        }
        setProgress(targetPage, endPage);
        continue; // 跳过此页，继续尝试下一页
      } else {
        emptyPageCount = 0; // 有数据就重置空页计数
      }

      allData = allData.concat(result.data);

      // 实时去重
      if (dedupEnabled && allData.length > 0) {
        const before = allData.length;
        allData = dedupData(allData);
        const removed = before - allData.length;
        if (removed > 0) {
          log(`  ↻ 去重：移除 ${removed} 条重复数据`);
        }
      }

      updateCount(allData.length);
      setProgress(targetPage, endPage);
      log(`第 ${targetPage} 页: 采集 ${result.data.length} 条，累计 ${allData.length} 条`);
    }
  }

  // 最终去重
  if (dedupEnabled && allData.length > 0) {
    const before = allData.length;
    allData = dedupData(allData);
    const removed = before - allData.length;
    if (removed > 0) {
      log(`↻ 最终去重：移除 ${removed} 条，剩余 ${allData.length} 条`);
    }
  }

  setStatus(`采集完成！共 ${allData.length} 条数据`, 'success');
  log(`✅ 采集完成，共 ${allData.length} 条`);
  setButtonState(false);
  document.getElementById('crmBtnExport').disabled = allData.length === 0;

  } catch (e) {
    setStatus(`采集出错: ${e.message}`, 'error');
    log(`❌ 错误: ${e.message}`);
    setButtonState(false);
  } finally {
    isCollecting = false;
  }
}

async function scrapeCurrentPageOnly() {
  isCollecting = true;
  setButtonState(true);
  setStatus('正在采集当前页...', 'info');
  dedupEnabled = true; // 始终去重

  try {
    const result = await scrapeSinglePage();

    allData = allData.concat(result.data);

    // 去重
    if (dedupEnabled) {
      const before = allData.length;
      allData = dedupData(allData);
      const removed = before - allData.length;
      if (removed > 0) {
        log(`↻ 去重：移除 ${removed} 条重复数据`);
      }
    }

    updateCount(allData.length);
    setStatus(`当前页采集完成，新增 ${result.data.length} 条，累计 ${allData.length} 条`, 'success');
    log(`当前页采集 ${result.data.length} 条，累计 ${allData.length} 条`);
    document.getElementById('crmBtnExport').disabled = allData.length === 0;
  } catch (e) {
    setStatus(`采集出错: ${e.message}`, 'error');
    log(`❌ 错误: ${e.message}`);
  } finally {
    isCollecting = false;
    setButtonState(false);
  }
}

// ============================================================
// 手动翻页
// ============================================================
async function manualPrevPage() {
  try {
    const res = await sendToContent({ action: 'CLICK_PREV_PAGE' });
    if (res && res.success) {
      log('↩ 点击网页上一页');
    } else {
      setStatus('已是第一页，无法往前翻', 'error');
    }
  } catch (e) {
    setStatus(`翻页失败: ${e.message}`, 'error');
  }
}

async function manualNextPage() {
  try {
    const res = await sendToContent({ action: 'CLICK_NEXT_PAGE' });
    if (res && res.success) {
      log('↪ 点击网页下一页');
    } else {
      setStatus('已是最后一页，无法往后翻', 'error');
    }
  } catch (e) {
    setStatus(`翻页失败: ${e.message}`, 'error');
  }
}

// ============================================================
// XLSX 导出
// ============================================================
async function exportToExcel() {
  if (allData.length === 0) {
    setStatus('没有可导出的数据', 'error');
    return;
  }

  try {
    setStatus('正在生成 Excel...', 'info');

    const now = new Date();
    const ts = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
    const filename = `阿里CRM客户列表_${ts}_共${allData.length}条.xlsx`;

    const rows = [EXCEL_HEADERS];
    allData.forEach(item => {
      rows.push(DATA_KEYS.map(key => item[key] || ''));
    });
    const colWidths = [
      { wch: 20 }, { wch: 25 }, { wch: 12 }, { wch: 12 },
      { wch: 10 }, { wch: 20 }, { wch: 10 },
      { wch: 10 }, { wch: 12 }, { wch: 15 },
      { wch: 12 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 12 }
    ];

    // 通过 background script 生成 Excel（不受页面 CSP 限制）
    let response;
    try {
      response = await chrome.runtime.sendMessage({
        action: 'EXPORT_XLSX',
        rows, colWidths, filename
      });
    } catch (e) {
      if (e.message.includes('Extension context invalidated') || e.message.includes('disconnected')) {
        throw new Error('插件已更新，请刷新页面后重试');
      }
      throw e;
    }

    if (!response || !response.success) {
      throw new Error(response?.error || '导出失败');
    }

    // 下载文件
    const binaryStr = atob(response.base64);
    const bytes = new Uint8Array(binaryStr.length);
    for (let i = 0; i < binaryStr.length; i++) {
      bytes[i] = binaryStr.charCodeAt(i);
    }
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = response.filename;
    a.click();
    URL.revokeObjectURL(url);

    setStatus(`✅ Excel已导出：${response.filename}`, 'success');
    log(`导出成功: ${response.filename}`);
  } catch (e) {
    setStatus(`导出失败: ${e.message}`, 'error');
    log(`❌ 导出错误: ${e.message}`);
  }
}

function clearData() {
  if (allData.length === 0) return;
  // 使用自定义确认代替 confirm()
  if (!window.confirm(`确定清空已采集的 ${allData.length} 条数据吗？`)) return;
  allData = [];
  updateCount(0);
  setProgress(0, 0);
  const logArea = document.getElementById('crmLogArea');
  if (logArea) logArea.innerHTML = '';
  setStatus('数据已清空', 'info');
  setButtonState(false);
}

// ============================================================
// 事件绑定
// ============================================================
function bindEvents() {
  document.getElementById('crmBtnStart')?.addEventListener('click', scrapeAllPages);
  document.getElementById('crmBtnStartCurrent')?.addEventListener('click', scrapeCurrentPageOnly);
  document.getElementById('crmBtnStop')?.addEventListener('click', () => {
    stopRequested = true;
    setStatus('正在停止...', 'info');
    log('用户请求停止采集');
  });
  document.getElementById('crmBtnExport')?.addEventListener('click', exportToExcel);
  document.getElementById('crmBtnClear')?.addEventListener('click', clearData);
  document.getElementById('crmBtnPrevPage')?.addEventListener('click', manualPrevPage);
  document.getElementById('crmBtnNextPage')?.addEventListener('click', manualNextPage);
  document.getElementById('crmCloseBtn')?.addEventListener('click', hidePanel);
  document.getElementById('crmMinBtn')?.addEventListener('click', minimizePanel);

  // 初始化
  setButtonState(false);
  setStatus('请在阿里巴巴国际站CRM客户列表页使用', 'info');
  log('浮窗插件已就绪，请在客户列表页点击采集');
}

// ============================================================
// 接收 popup 的 toggle 消息
// ============================================================
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'TOGGLE_PANEL') {
    const panel = document.getElementById('crm-collector-panel');
    if (panel) {
      if (panel.classList.contains('crm-panel-hidden')) {
        showPanel();
      } else {
        hidePanel();
      }
    }
    sendResponse({ success: true });
    return true;
  }

  if (message.action === 'GET_PANEL_STATE') {
    const panel = document.getElementById('crm-collector-panel');
    sendResponse({
      visible: panel && !panel.classList.contains('crm-panel-hidden')
    });
    return true;
  }
});

// ============================================================
// 初始化
// ============================================================
function initFloatingPanel() {
  injectPanelHTML();
  injectToggleBtn();
  enableDrag();
  bindEvents();
  console.log('[CRM采集插件] 浮窗面板已加载');
}

// 等待 DOM 就绪后初始化
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initFloatingPanel);
} else {
  initFloatingPanel();
}
