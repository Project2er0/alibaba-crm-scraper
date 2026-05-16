/**
 * 阿里巴巴CRM客户采集 - Popup 交互逻辑
 * 负责UI控制、多页采集协调、Excel导出
 */

// ============================================================
// 全局状态
// ============================================================
let allData = [];          // 累积的所有采集数据
let isCollecting = false;  // 是否正在采集
let currentTabId = null;   // 当前操作的Tab ID
let stopRequested = false; // 用户请求停止

// Excel 列头定义
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
// UI 工具函数
// ============================================================
function setStatus(text, type = 'info') {
  const el = document.getElementById('status');
  el.textContent = text;
  el.className = `status ${type}`;
}

function setProgress(current, total) {
  const bar = document.getElementById('progressBar');
  const label = document.getElementById('progressLabel');
  const pct = total > 0 ? Math.min(100, Math.round(current / total * 100)) : 0;
  bar.style.width = pct + '%';
  label.textContent = total > 0 ? `${current} / ${total} 页` : '';
}

function updateCount(count) {
  document.getElementById('dataCount').textContent = count;
}

function setButtonState(collecting) {
  document.getElementById('btnStart').disabled = collecting;
  document.getElementById('btnStop').disabled = !collecting;
  document.getElementById('btnStartCurrent').disabled = collecting;
  document.getElementById('btnExport').disabled = allData.length === 0;
  document.getElementById('btnPrevPage').disabled = collecting;
  document.getElementById('btnNextPage').disabled = collecting;
}

function log(msg) {
  const el = document.getElementById('logArea');
  const line = document.createElement('div');
  line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}

// ============================================================
// 采集逻辑
// ============================================================

async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

function sendToContent(tabId, message, timeoutMs = 8000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('消息超时')), timeoutMs);
    chrome.tabs.sendMessage(tabId, message, (response) => {
      clearTimeout(timer);
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        resolve(response);
      }
    });
  });
}

function wait(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function scrapeSinglePage(tabId) {
  await wait(1500);
  const response = await sendToContent(tabId, { action: 'SCRAPE_CURRENT_PAGE' });
  if (!response || !response.success) {
    throw new Error(response?.error || '采集失败');
  }
  return response;
}

async function scrapeAllPages() {
  isCollecting = true;
  stopRequested = false;
  allData = [];
  setButtonState(true);
  setStatus('正在采集...', 'info');
  log('开始多页采集...');

  try {
    const tab = await getActiveTab();
    currentTabId = tab.id;

    const firstResult = await scrapeSinglePage(currentTabId);
    allData = allData.concat(firstResult.data);
    updateCount(allData.length);
    log(`第 1 页: 采集到 ${firstResult.data.length} 条`);

    const totalPages = firstResult.pagination.totalPages || 1;
    setProgress(1, totalPages);

    if (totalPages <= 1) {
      log('仅1页，采集完成');
    } else {
      for (let page = 2; page <= totalPages; page++) {
        if (stopRequested) {
          log('用户已停止采集');
          break;
        }

        setStatus(`正在采集第 ${page}/${totalPages} 页...`, 'info');

        const clickResult = await sendToContent(currentTabId, { action: 'CLICK_NEXT_PAGE' });
        if (!clickResult || !clickResult.success) {
          log(`第 ${page} 页：无法翻页，停止`);
          break;
        }

        await wait(2500);

        const result = await scrapeSinglePage(currentTabId);
        allData = allData.concat(result.data);
        updateCount(allData.length);
        setProgress(page, totalPages);
        log(`第 ${page} 页: 采集到 ${result.data.length} 条，累计 ${allData.length} 条`);
      }
    }

    setStatus(`采集完成！共 ${allData.length} 条数据`, 'success');
    log(`✅ 采集完成，共 ${allData.length} 条`);
    setButtonState(false);
    document.getElementById('btnExport').disabled = allData.length === 0;

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

  try {
    const tab = await getActiveTab();
    currentTabId = tab.id;
    const result = await scrapeSinglePage(currentTabId);

    allData = allData.concat(result.data);
    updateCount(allData.length);
    setStatus(`当前页采集完成，新增 ${result.data.length} 条，累计 ${allData.length} 条`, 'success');
    log(`当前页采集 ${result.data.length} 条，累计 ${allData.length} 条`);
    document.getElementById('btnExport').disabled = allData.length === 0;
  } catch (e) {
    setStatus(`采集出错: ${e.message}`, 'error');
    log(`❌ 错误: ${e.message}`);
  } finally {
    isCollecting = false;
    setButtonState(false);
  }
}

// ============================================================
// 手动翻页：直接触发网页中的上一页/下一页按钮
// ============================================================

async function manualPrevPage() {
  try {
    const tab = await getActiveTab();
    if (!tab) return;
    const res = await sendToContent(tab.id, { action: 'CLICK_PREV_PAGE' });
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
    const tab = await getActiveTab();
    if (!tab) return;
    const res = await sendToContent(tab.id, { action: 'CLICK_NEXT_PAGE' });
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
// Excel 导出
// ============================================================
async function exportToExcel() {
  if (allData.length === 0) {
    setStatus('没有可导出的数据', 'error');
    return;
  }

  try {
    setStatus('正在生成 Excel...', 'info');

    const rows = [EXCEL_HEADERS];
    allData.forEach(item => {
      rows.push(DATA_KEYS.map(key => item[key] || ''));
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(rows);

    ws['!cols'] = [
      { wch: 20 }, { wch: 25 }, { wch: 25 }, { wch: 15 },
      { wch: 12 }, { wch: 15 }, { wch: 20 },
      { wch: 15 }, { wch: 10 }, { wch: 12 },
      { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 15 }
    ];
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };

    XLSX.utils.book_append_sheet(wb, ws, '客户列表');

    const now = new Date();
    const ts = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
    const filename = `阿里CRM客户列表_${ts}_共${allData.length}条.xlsx`;

    XLSX.writeFile(wb, filename);
    setStatus(`✅ Excel已导出：${filename}`, 'success');
    log(`导出成功: ${filename}`);
  } catch (e) {
    setStatus(`导出失败: ${e.message}`, 'error');
    log(`❌ 导出错误: ${e.message}`);
  }
}

function clearData() {
  if (allData.length === 0) return;
  if (confirm(`确定清空已采集的 ${allData.length} 条数据吗？`)) {
    allData = [];
    updateCount(0);
    setProgress(0, 0);
    document.getElementById('logArea').innerHTML = '';
    setStatus('数据已清空', 'info');
    setButtonState(false);
  }
}

// ============================================================
// 事件绑定
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('btnStart').addEventListener('click', scrapeAllPages);
  document.getElementById('btnStartCurrent').addEventListener('click', scrapeCurrentPageOnly);
  document.getElementById('btnStop').addEventListener('click', () => {
    stopRequested = true;
    setStatus('正在停止...', 'info');
    log('用户请求停止采集');
  });
  document.getElementById('btnExport').addEventListener('click', exportToExcel);
  document.getElementById('btnClear').addEventListener('click', clearData);

  // 手动翻页
  document.getElementById('btnPrevPage').addEventListener('click', manualPrevPage);
  document.getElementById('btnNextPage').addEventListener('click', manualNextPage);

  // 初始化
  setButtonState(false);
  setStatus('请在阿里巴巴国际站CRM客户列表页使用', 'info');
  log('插件已就绪，请在客户列表页点击采集');
});
