/**
 * 阿里巴巴CRM客户列表采集 - Content Script
 * 负责在页面中提取表格数据，通过消息与popup通信
 */

// ============================================================
// 数据提取核心函数
// ============================================================

/**
 * 从一行tr元素中提取结构化客户数据
 * 基于 2025-04 阿里CRM新表格结构（data-next-table-col 在 td 上）
 * col-0: 复选框 | col-1: 客户 | col-2: 业务员 | col-3: 客户阶段
 * col-4: 跟进状态 | col-5: 小记时间 | col-6: 采购意向（星级）
 * col-7: 年采购额 | col-8: 采购品类 | col-9: 国家/地区
 * col-10: 所属客群 | col-11: 客户来源 | col-12: 商业类型 | col-13: 建档时间
 * @param {HTMLElement} row - 表格行元素
 * @returns {Object} 客户数据对象
 */
function extractRowData(row) {
  const getCell = (col) => row.querySelector(`td[data-next-table-col="${col}"]`);
  const text = (el) => el ? el.textContent.trim() : '';

  // col-1: 联系人姓名 + 公司名
  const col1 = getCell(1);
  const contactName = text(col1?.querySelector('.column-component-company-name-main-contact-container .name span'));
  const companyName = text(col1?.querySelector('.column-component-company-name-companyName'));

  // col-2: 业务员
  const col2 = getCell(2);
  const owner = text(col2?.querySelector('.table-cell-content-render-container-trigger-container-item'));

  // col-3: 客户阶段
  const col3 = getCell(3);
  const customerStage = text(col3?.querySelector('.alicrm-customer-group-select-dropdown-trigger'));

  // col-4: 跟进状态（label + content）
  const col4 = getCell(4);
  const followStatus = text(col4?.querySelector('.common-columns-columns-recent-note-status-column-label'));
  const followContent = text(col4?.querySelector('.common-columns-columns-recent-note-status-column-content'));

  // col-5: 小记时间
  const col5 = getCell(5);
  const followTime = text(col5?.querySelector('.crm-c-customer-follow-time'));

  // col-6: 采购意向（星级 - 通过 overlay 宽度计算）
  const col6 = getCell(6);
  let purchaseIntent = '';
  if (col6) {
    const overlay = col6.querySelector('.next-rating-overlay');
    if (overlay) {
      const width = parseInt(overlay.style.width) || 0;
      if (width >= 50) purchaseIntent = '★★★ 高';
      else if (width >= 30) purchaseIntent = '★★ 中';
      else if (width >= 10) purchaseIntent = '★ 低';
      else purchaseIntent = '☆ 无';
    }
    // 备选：label 文本
    const label = text(col6?.querySelector('.alicrm-bc-rating .label'));
    if (!purchaseIntent && label) purchaseIntent = label;
  }

  // col-7: 年采购额
  const col7 = getCell(7);
  const annualPurchase = text(col7);

  // col-8: 采购品类
  const col8 = getCell(8);
  const purchaseCategory = text(col8?.querySelector('.purchase-category-handler-purchase-category-column .text'));

  // col-9: 国家/地区
  const col9 = getCell(9);
  const country = text(col9?.querySelector('.table-cell-content-render-container-trigger-container-item'));

  // col-10: 所属客群
  const col10 = getCell(10);
  const customerGroup = text(col10?.querySelector('.common-columns-columns-belong-group-column .text'));

  // col-11: 客户来源
  const col11 = getCell(11);
  const customerOrigin = text(col11?.querySelector('.common-columns-columns-customer-origin-label'));

  // col-12: 商业类型
  const col12 = getCell(12);
  const bizType = text(col12?.querySelector('.table-cell-content-render-container-trigger-container-item'));

  // col-13: 建档时间
  const col13 = getCell(13);
  const createTime = text(col13?.querySelector('.crm-c-customer-follow-time'));

  return {
    contactName,
    companyName,
    owner,
    customerStage,
    followStatus,
    followContent,
    followTime,
    purchaseIntent,
    annualPurchase,
    purchaseCategory,
    country,
    customerGroup,
    customerOrigin,
    bizType,
    createTime
  };
}

/**
 * 采集当前页面所有行的数据
 * @returns {Array} 数据数组
 */
function scrapeCurrentPage() {
  const rows = document.querySelectorAll('.next-table-body tbody tr.next-table-row');
  const data = [];

  rows.forEach((row, index) => {
    try {
      const rowData = extractRowData(row);
      // 过滤掉完全空白的行（至少有一个字段有值才保留）
      const hasAnyData = Object.values(rowData).some(v => v && v.trim() !== '');
      if (hasAnyData) {
        data.push(rowData);
      }
    } catch (e) {
      console.warn(`[CRM采集] 第${index + 1}行解析失败:`, e);
    }
  });

  console.log(`[CRM采集] 当前页采集到 ${data.length} 条数据`);
  return data;
}

/**
 * 获取当前分页信息
 * @returns {Object} { currentPage, totalPages, totalRecords }
 */
function getPaginationInfo() {
  let totalRecords = 0;
  let currentPage = 1;
  let totalPages = 1;

  const totalMatch = document.querySelector('.next-pagination-total, [class*="pagination"]');
  if (totalMatch) {
    const text = totalMatch.textContent;
    const match = text.match(/共(\d+)条/);
    if (match) totalRecords = parseInt(match[1]);
    const pageMatch = text.match(/(\d+)\/(\d+)/);
    if (pageMatch) {
      currentPage = parseInt(pageMatch[1]);
      totalPages = parseInt(pageMatch[2]);
    }
  }

  const pageText = Array.from(document.querySelectorAll('*'))
    .find(el => el.childNodes.length === 1 && /共\d+条/.test(el.textContent));
  if (pageText) {
    const m = pageText.textContent.match(/共(\d+)条/);
    if (m) totalRecords = parseInt(m[1]);
  }

  const activePage = document.querySelector('.next-pagination-item.next-current, .next-pagination-current');
  if (activePage) currentPage = parseInt(activePage.textContent) || 1;

  return { currentPage, totalPages, totalRecords };
}

/**
 * 点击下一页按钮（硬编码选择器）
 * @returns {boolean} 是否成功点击
 */
function clickNextPage() {
  // 策略1: 按钮容器（阿里 Fusion/Next 组件常见 class）
  const nextBtn = document.querySelector(
    '.next-pagination-next:not([disabled]), ' +
    'button[aria-label="next"]:not([disabled]), ' +
    '.next-pagination .next-btn-text:last-child:not([disabled])'
  );
  if (nextBtn && !nextBtn.disabled && !nextBtn.classList.contains('disabled')) {
    nextBtn.click();
    return true;
  }
  // 策略2: 按图标定位——找包含向右箭头的翻页按钮
  const arrowNext = document.querySelector(
    '.next-icon-arrow-right.next-icon-last'
  );
  if (arrowNext) {
    const clickable = arrowNext.closest('button, a, [role="button"], .next-pagination-next');
    if (clickable && !clickable.disabled && !clickable.classList.contains('disabled') && !clickable.classList.contains('next-pagination-next-disabled')) {
      clickable.click();
      return true;
    }
  }
  return false;
}

/**
 * 点击上一页按钮（硬编码选择器）
 * @returns {boolean} 是否成功点击
 */
function clickPrevPage() {
  const prevBtn = document.querySelector(
    '.next-pagination-prev:not([disabled]), ' +
    'button[aria-label="prev"]:not([disabled]), ' +
    'button[aria-label="previous"]:not([disabled])'
  );
  if (prevBtn && !prevBtn.disabled && !prevBtn.classList.contains('disabled')) {
    prevBtn.click();
    return true;
  }
  // 备用: 按图标定位——找包含向左箭头的翻页按钮
  const arrowPrev = document.querySelector(
    '.next-icon-arrow-left.next-icon-first'
  );
  if (arrowPrev) {
    const clickable = arrowPrev.closest('button, a, [role="button"], .next-pagination-prev');
    if (clickable && !clickable.disabled && !clickable.classList.contains('disabled') && !clickable.classList.contains('next-pagination-prev-disabled')) {
      clickable.click();
      return true;
    }
  }
  return false;
}

/**
 * 跳转到指定页（使用分页组件的页码输入框或直接点击页码）
 * @param {number} pageNum - 目标页码
 * @returns {boolean} 是否成功跳转
 */
function jumpToPage(pageNum) {
  // 策略1: 找分页组件中的输入框（阿里 Fusion Next UI 常见模式）
  const jumpInput = document.querySelector(
    '.next-pagination-jump-input input, ' +
    '.next-pagination-jump input, ' +
    'input.next-pagination-jump-input, ' +
    '.next-pagination input[type="text"], ' +
    '.next-pagination .next-input'
  );
  if (jumpInput) {
    jumpInput.value = pageNum;
    jumpInput.dispatchEvent(new Event('input', { bubbles: true }));
    jumpInput.dispatchEvent(new Event('change', { bubbles: true }));
    // 尝试找跳转按钮（通常是"跳至"按钮或回车）
    const jumpBtn = jumpInput.closest('.next-pagination-jump')?.querySelector('button, a, [role="button"]')
      || document.querySelector('.next-pagination-jump-go button, .next-pagination-jump-go a');
    if (jumpBtn) {
      jumpBtn.click();
      return true;
    }
    // 没有跳转按钮，按回车
    jumpInput.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));
    return true;
  }

  // 策略2: 直接点击页码
  const pageLinks = document.querySelectorAll(
    '.next-pagination-number, .next-pagination-item:not(.next-prev):not(.next-next):not(.next-current)'
  );
  for (const link of pageLinks) {
    if (parseInt(link.textContent) === pageNum) {
      link.click();
      return true;
    }
  }

  // 策略3: 点击最后一页来获取完整页码列表（某些分页组件懒加载页码）
  const lastBtn = document.querySelector('.next-pagination-last:not([disabled])');
  if (lastBtn && !lastBtn.classList.contains('disabled')) {
    lastBtn.click();
    return true;
  }

  return false;
}

/**
 * 用用户指定的选择器点击翻页
 */
function clickPageBySelector(selector) {
  if (!selector) return false;
  const btn = document.querySelector(selector);
  if (btn && !btn.disabled && !btn.classList.contains('disabled')) {
    btn.click();
    return true;
  }
  return false;
}

// ============================================================
// 消息监听 - 与popup通信
// ============================================================
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'PING') {
    sendResponse({ alive: true, url: window.location.href });
    return true;
  }

  if (message.action === 'SCRAPE_CURRENT_PAGE') {
    try {
      const data = scrapeCurrentPage();
      const pagination = getPaginationInfo();
      sendResponse({ success: true, data, pagination });
    } catch (e) {
      sendResponse({ success: false, error: e.message });
    }
    return true;
  }

  if (message.action === 'CLICK_NEXT_PAGE') {
    // 优先用用户自定义选择器，没有就用硬编码
    const clicked = message.selector
      ? clickPageBySelector(message.selector)
      : clickNextPage();
    sendResponse({ success: clicked });
    return true;
  }

  if (message.action === 'CLICK_PREV_PAGE') {
    const clicked = message.selector
      ? clickPageBySelector(message.selector)
      : clickPrevPage();
    sendResponse({ success: clicked });
    return true;
  }

  if (message.action === 'GET_PAGINATION') {
    const pagination = getPaginationInfo();
    sendResponse({ success: true, pagination });
    return true;
  }

  if (message.action === 'JUMP_TO_PAGE') {
    const success = jumpToPage(message.page);
    sendResponse({ success });
    return true;
  }
});

console.log('[CRM采集插件] Content script 已加载');
