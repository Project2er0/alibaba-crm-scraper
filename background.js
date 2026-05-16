// ============================================================
// Background Service Worker — 处理 XLSX 导出（不受页面 CSP 限制）
// 内联 XLSX 生成器，纯数据运算，不依赖 DOM
// ============================================================

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'EXPORT_XLSX') {
    try {
      const { rows, colWidths, filename } = message;

      // 用 XLSX.writeFile 的方式，但 SW 里没有 DOM
      // 所以我们手动构造 workbook 对象，然后直接调用内部 buildXlsx
      // xlsx.min.js 暴露的是 XLSX.writeFile(wb, filename)
      // writeFile 内部会调用 buildXlsx(aoa, sheetName, colWidths)
      // 我们绕过 writeFile，直接拿到 zip 数据

      // 方案：重新 fetch xlsx.min.js 源码，提取 buildXlsx 函数
      // 更简单的方案：直接在这里内联一个轻量 XLSX 生成器

      const zipData = buildXlsxInSW(rows, colWidths);

      // 转 base64 传输（避免 structured clone 对二进制数据的问题）
      let binary = '';
      const bytes = new Uint8Array(zipData);
      for (let i = 0; i < bytes.length; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      const base64 = btoa(binary);

      sendResponse({ success: true, base64, filename });
    } catch (e) {
      sendResponse({ success: false, error: e.message });
    }
    return true;
  }
});

// ============================================================
// 内联 XLSX 生成器（不依赖 DOM，适用于 Service Worker）
// ============================================================

function strToBytes(str) {
  return new TextEncoder().encode(str);
}

function uint32LE(n) {
  return [n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, (n >> 24) & 0xFF];
}
function uint16LE(n) {
  return [n & 0xFF, (n >> 8) & 0xFF];
}

const crcTable = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c;
  }
  return t;
})();

function crc32(data) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) crc = crcTable[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function buildZip(files) {
  const localHeaders = [];
  const centralHeaders = [];
  let offset = 0;

  files.forEach(file => {
    const nameBytes = strToBytes(file.name);
    const data = file.data instanceof Uint8Array ? file.data : strToBytes(file.data);
    const crc = crc32(data);
    const size = data.length;

    const local = new Uint8Array(30 + nameBytes.length + size);
    let p = 0;
    local[p++]=0x50;local[p++]=0x4B;local[p++]=0x03;local[p++]=0x04;
    local[p++]=0x14;local[p++]=0x00;
    local[p++]=0x00;local[p++]=0x00;
    local[p++]=0x00;local[p++]=0x00;
    local[p++]=0x00;local[p++]=0x00;local[p++]=0x00;local[p++]=0x00;
    uint32LE(crc).forEach(b => local[p++]=b);
    uint32LE(size).forEach(b => local[p++]=b);
    uint32LE(size).forEach(b => local[p++]=b);
    uint16LE(nameBytes.length).forEach(b => local[p++]=b);
    local[p++]=0x00;local[p++]=0x00;
    nameBytes.forEach(b => local[p++]=b);
    data.forEach(b => local[p++]=b);

    localHeaders.push({ local, offset, nameBytes, crc, size });
    offset += local.length;
  });

  const centralParts = [];
  localHeaders.forEach(h => {
    const cd = new Uint8Array(46 + h.nameBytes.length);
    let p = 0;
    cd[p++]=0x50;cd[p++]=0x4B;cd[p++]=0x01;cd[p++]=0x02;
    cd[p++]=0x14;cd[p++]=0x00;
    cd[p++]=0x14;cd[p++]=0x00;
    cd[p++]=0x00;cd[p++]=0x00;
    cd[p++]=0x00;cd[p++]=0x00;
    cd[p++]=0x00;cd[p++]=0x00;
    cd[p++]=0x00;cd[p++]=0x00;
    uint32LE(h.crc).forEach(b => cd[p++]=b);
    uint32LE(h.size).forEach(b => cd[p++]=b);
    uint32LE(h.size).forEach(b => cd[p++]=b);
    uint16LE(h.nameBytes.length).forEach(b => cd[p++]=b);
    cd[p++]=0x00;cd[p++]=0x00;
    cd[p++]=0x00;cd[p++]=0x00;
    cd[p++]=0x00;cd[p++]=0x00;
    cd[p++]=0x00;cd[p++]=0x00;
    cd[p++]=0x00;cd[p++]=0x00;cd[p++]=0x00;cd[p++]=0x00;
    uint32LE(h.offset).forEach(b => cd[p++]=b);
    h.nameBytes.forEach(b => cd[p++]=b);
    centralParts.push(cd);
  });

  const cdSize = centralParts.reduce((s, c) => s + c.length, 0);
  const cdOffset = offset;

  const eocd = new Uint8Array(22);
  let p = 0;
  eocd[p++]=0x50;eocd[p++]=0x4B;eocd[p++]=0x05;eocd[p++]=0x06;
  eocd[p++]=0x00;eocd[p++]=0x00;
  eocd[p++]=0x00;eocd[p++]=0x00;
  uint16LE(files.length).forEach(b => eocd[p++]=b);
  uint16LE(files.length).forEach(b => eocd[p++]=b);
  uint32LE(cdSize).forEach(b => eocd[p++]=b);
  uint32LE(cdOffset).forEach(b => eocd[p++]=b);
  eocd[p++]=0x00;eocd[p++]=0x00;

  const totalSize = localHeaders.reduce((s, h) => s + h.local.length, 0) + cdSize + 22;
  const result = new Uint8Array(totalSize);
  let pos = 0;
  localHeaders.forEach(h => { result.set(h.local, pos); pos += h.local.length; });
  centralParts.forEach(cd => { result.set(cd, pos); pos += cd.length; });
  result.set(eocd, pos);
  return result;
}

function xmlEsc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildXlsxInSW(aoa, colWidths) {
  const sheetName = '客户列表';

  const sharedStrings = [];
  const ssMap = {};
  function getSSI(s) {
    const key = String(s);
    if (ssMap[key] === undefined) {
      ssMap[key] = sharedStrings.length;
      sharedStrings.push(key);
    }
    return ssMap[key];
  }

  function colLetter(n) {
    let s = '';
    n++;
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  let sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  sheetXml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
  sheetXml += '<sheetViews><sheetView tabSelected="1" workbookViewId="0">';
  sheetXml += '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>';
  sheetXml += '</sheetView></sheetViews>';

  if (colWidths && colWidths.length > 0) {
    sheetXml += '<cols>';
    colWidths.forEach((w, i) => {
      // 兼容 {wch: 20} 和纯数字 20 两种格式
      const width = typeof w === 'object' ? (w.wch || w.width || 20) : w;
      sheetXml += `<col min="${i+1}" max="${i+1}" width="${width}" customWidth="1"/>`;
    });
    sheetXml += '</cols>';
  }

  sheetXml += '<sheetData>';
  aoa.forEach((row, ri) => {
    sheetXml += `<row r="${ri+1}">`;
    row.forEach((cell, ci) => {
      const addr = `${colLetter(ci)}${ri+1}`;
      const val = cell == null ? '' : String(cell);
      if (val === '') {
        sheetXml += `<c r="${addr}"/>`;
      } else {
        const si = getSSI(val);
        sheetXml += `<c r="${addr}" t="s"><v>${si}</v></c>`;
      }
    });
    sheetXml += '</row>';
  });
  sheetXml += '</sheetData></worksheet>';

  const ssXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">
${sharedStrings.map(s => `<si><t xml:space="preserve">${xmlEsc(s)}</t></si>`).join('\n')}
</sst>`;

  const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="${xmlEsc(sheetName)}" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

  const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

  const pkgRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;

  return buildZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: pkgRels },
    { name: 'xl/workbook.xml', data: wbXml },
    { name: 'xl/_rels/workbook.xml.rels', data: wbRels },
    { name: 'xl/worksheets/sheet1.xml', data: sheetXml },
    { name: 'xl/sharedStrings.xml', data: ssXml },
  ]);
}
