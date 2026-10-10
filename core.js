/*
 * IntegratorCore — 通用「模版 + 多份数据文件」整合核心
 * 1) 识别模版中的「目标/达成」板块（自动适配项目增减）
 * 2) 识别每份数据文件的表头/数据区/键列
 * 3) 按姓名等键自动匹配，写入模版的目标/达成单元格
 * 4) 直接改写模版 xlsx 内部 XML，完整保留格式、图片、公式
 *
 * 可在浏览器(挂到 window.IntegratorCore)或 Node(factory(JSZip, XLSX))中使用
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory;
  } else {
    root.IntegratorCore = factory(root.JSZip, root.XLSX);
  }
})(typeof self !== 'undefined' ? self : this, function (JSZip, XLSX) {
  'use strict';

  // ---------- 基础工具 ----------
  function colLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }
  function colNumber(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }
  function cellAddr(r, c) { return colLetter(c) + r; }
  function decodeXml(s) {
    return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  }
  function normKey(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/\s+/g, '').trim();
  }
  function normText(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/[\s\n\r]/g, '').toLowerCase();
  }
  function toNumber(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = String(v).replace(/,/g, '').replace(/%/g, '').trim();
    if (s === '') return null;
    const n = Number(s);
    return isFinite(n) ? n : null;
  }
  function lcs(a, b) {
    a = a || ''; b = b || '';
    if (!a || !b) return 0;
    let dp = new Array(b.length + 1).fill(0), best = 0;
    for (let i = 1; i <= a.length; i++) {
      let prev = 0;
      for (let j = 1; j <= b.length; j++) {
        const cur = dp[j];
        if (a[i - 1] === b[j - 1]) { dp[j] = prev + 1; if (dp[j] > best) best = dp[j]; }
        else dp[j] = 0;
        prev = cur;
      }
    }
    return best;
  }

  // ---------- 工作簿读取 ----------
  function readWorkbook(data, type) {
    return XLSX.read(data, { type: type || 'array', cellFormula: true, cellStyles: false });
  }

  // 展开合并单元格，返回二维数组（绝对 0 基索引）
  function expandGrid(ws) {
    if (!ws || !ws['!ref']) return { grid: [], range: { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } } };
    const range = XLSX.utils.decode_range(ws['!ref']);
    const grid = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        row[c] = cell ? cell.v : null;
      }
      grid[r] = row;
    }
    (ws['!merges'] || []).forEach(function (m) {
      const v = grid[m.s.r] ? grid[m.s.r][m.s.c] : null;
      for (let r = m.s.r; r <= m.e.r; r++) {
        if (!grid[r]) grid[r] = [];
        for (let c = m.s.c; c <= m.e.c; c++) grid[r][c] = v;
      }
    });
    return { grid: grid, range: range };
  }

  function nearestLabel(grid, row, col, fromRow) {
    for (let r = row - 1; r >= (fromRow || 0); r--) {
      const v = grid[r] ? grid[r][col] : null;
      if (v !== null && v !== undefined && String(v).trim() !== '') return String(v).trim();
    }
    return '';
  }

  // ---------- 模版识别 ----------
  // 返回 { sheetName, keyCol, keyRow, dataStart, lastRow, maxCol, items, extraCols }
  function detectTemplate(wb, sheetName) {
    const ws = wb.Sheets[sheetName];
    const ex = expandGrid(ws);
    const grid = ex.grid, range = ex.range;

    // 键列：优先精确 "姓名"
    let keyRow = -1, keyCol = -1;
    for (let r = range.s.r; r <= range.e.r && keyRow < 0; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const v = grid[r] ? grid[r][c] : null;
        if (v !== null && String(v).trim() === '姓名') { keyRow = r; keyCol = c; break; }
      }
    }
    if (keyCol < 0) {
      for (let r = range.s.r; r <= range.e.r && keyRow < 0; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const v = grid[r] ? grid[r][c] : null;
          if (v !== null && String(v).indexOf('姓名') >= 0) { keyRow = r; keyCol = c; break; }
        }
      }
    }
    if (keyCol < 0) throw new Error('模版中未找到「姓名」列');

    // 数据起始行
    let dataStart = -1;
    for (let r = keyRow + 1; r <= range.e.r; r++) {
      const v = grid[r] ? grid[r][keyCol] : null;
      if (v !== null && String(v).trim() !== '') { dataStart = r; break; }
    }
    let lastRow = dataStart;
    for (let r = dataStart; r <= range.e.r; r++) {
      const v = grid[r] ? grid[r][keyCol] : null;
      if (v !== null && String(v).trim() !== '') lastRow = r;
    }

    // 目标/达成 板块（在表头行 keyRow 上成对出现）
    const items = [];
    const itemCols = {};
    const maxCol = range.e.c;
    function headerAt(r, c) { const v = grid[r] ? grid[r][c] : null; return v === null || v === undefined ? '' : String(v).trim(); }
    for (let c = range.s.c; c <= maxCol; c++) {
      const hs = headerAt(keyRow, c);
      if (hs.indexOf('目标') < 0 || hs.indexOf('率') >= 0) continue;
      let d = -1;
      for (let x = c + 1; x <= maxCol; x++) {
        const as = headerAt(keyRow, x);
        if (as === '达成' || (as.indexOf('达成') >= 0 && as.indexOf('率') < 0)) { d = x; break; }
      }
      if (d < 0) continue;
      let name = '';
      for (let r = keyRow - 1; r >= range.s.r; r--) {
        const nv = grid[r] ? grid[r][c] : null;
        if (nv !== null && String(nv).trim() !== '' && isNaN(Number(nv))) { name = String(nv).trim(); break; }
      }
      // 紧跟在「达成」后的 达成率 / 绩效 列（用于除零容错）
      let ratioCol = null, perfCol = null, e = d + 1;
      if (e <= maxCol && headerAt(keyRow, e).indexOf('率') >= 0) { ratioCol = e; e++; }
      if (e <= maxCol && headerAt(keyRow, e).indexOf('绩效') >= 0) { perfCol = e; e++; }
      items.push({
        name: name || ('板块' + (items.length + 1)),
        targetCol: c + 1, actualCol: d + 1,
        ratioCol: ratioCol !== null ? ratioCol + 1 : null,
        perfCol: perfCol !== null ? perfCol + 1 : null
      });
      for (let k = c; k < e; k++) itemCols[k] = true;
      c = e - 1;
    }

    // 其它可填列（无公式、非键列、非板块列、有表头）
    const extraCols = [];
    for (let c = keyCol + 1; c <= maxCol; c++) {
      if (itemCols[c]) continue;
      const hdr = nearestLabel(grid, keyRow, c, range.s.r);
      if (!hdr) continue;
      let hasFormula = false, numericOk = true;
      for (let r = dataStart; r <= lastRow; r++) {
        const cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        if (!cell) continue;
        if (cell.f) { hasFormula = true; break; }
        if (cell.v !== null && cell.v !== undefined && cell.v !== '' &&
            typeof cell.v === 'string' && toNumber(cell.v) === null) { numericOk = false; break; }
      }
      if (hasFormula || !numericOk) continue;
      extraCols.push({ name: hdr, col: c + 1 });
    }

    return {
      sheetName: sheetName, keyCol: keyCol + 1, keyRow: keyRow + 1,
      dataStart: dataStart + 1, lastRow: lastRow + 1, maxCol: maxCol + 1,
      items: items, extraCols: extraCols
    };
  }

  // ---------- 数据文件识别 ----------
  // 返回 { sheets:[{name, keyCol, keyRow, dataStart, columns:[{col,label}], rows:[{key,values}]}] }
  function detectData(wb, fileName, sheetKeysOverride) {
    const out = [];
    wb.SheetNames.forEach(function (sn) {
      const res = detectDataSheet(wb.Sheets[sn], sn, fileName);
      if (res) out.push(res);
    });
    return { sheets: out };
  }

  function detectDataSheet(ws, sheetName, fileName) {
    const ex = expandGrid(ws);
    const grid = ex.grid, range = ex.range;

    let keyRow = -1, keyCol = -1;
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 5) && keyRow < 0; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const v = grid[r] ? grid[r][c] : null;
        if (v !== null && String(v).trim() === '姓名') { keyRow = r; keyCol = c; break; }
      }
    }
    if (keyCol < 0) return null;

    let dataStart = -1;
    for (let r = keyRow + 1; r <= range.e.r; r++) {
      const v = grid[r] ? grid[r][keyCol] : null;
      if (v !== null && String(v).trim() !== '' && String(v).trim() !== '姓名') { dataStart = r; break; }
    }
    if (dataStart < 0) return null;

    const headerRows = [];
    for (let r = range.s.r; r < dataStart; r++) headerRows.push(r);

    const columns = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      if (c === keyCol) continue;
      const parts = [];
      headerRows.forEach(function (r) {
        const v = grid[r] ? grid[r][c] : null;
        if (v !== null && v !== undefined && String(v).trim() !== '') parts.push(String(v).trim());
      });
      const label = parts.join(' ');
      if (label) columns.push({ col: c + 1, label: label });
    }

    const rows = [];
    for (let r = dataStart; r <= range.e.r; r++) {
      const kv = grid[r] ? grid[r][keyCol] : null;
      if (kv === null || String(kv).trim() === '') continue;
      const values = {};
      columns.forEach(function (col) { values[col.col] = grid[r] ? grid[r][col.col - 1] : null; });
      rows.push({ key: String(kv).trim(), values: values });
    }

    return {
      fileName: fileName, sheetName: sheetName,
      keyCol: keyCol + 1, keyRow: keyRow + 1, dataStart: dataStart + 1,
      columns: columns, rows: rows
    };
  }

  // ---------- 自动映射 ----------
  // dataSheets: 来自 detectData 的 sheets；返回 { assignments, sheetKey }
  function sheetKeyOf(ds, idx) { return idx + '::' + ds.fileName + '::' + ds.sheetName; }

  function autoMap(template, dataSheets) {
    // 展平所有来源列
    const srcCols = [];
    dataSheets.forEach(function (ds, idx) {
      const sk = sheetKeyOf(ds, idx);
      ds.columns.forEach(function (col) {
        srcCols.push({ sheetKey: sk, col: col.col, label: col.label, ds: ds });
      });
    });

    function best(itemName, role) {
      let bestCol = null, bestScore = 0;
      srcCols.forEach(function (sc) {
        const lbl = normText(sc.label);
        const roleOk = role === 'target'
          ? lbl.indexOf('目标') >= 0
          : (lbl.indexOf('达成') >= 0 && lbl.indexOf('率') < 0);
        if (!roleOk) return;
        const sc1 = lcs(normText(itemName), lbl);
        if (sc1 > bestScore) { bestScore = sc1; bestCol = sc; }
      });
      return bestScore > 0 ? bestCol : null;
    }

    const fieldMap = {};
    template.items.forEach(function (it, i) {
      const t = best(it.name, 'target');
      const a = best(it.name, 'actual');
      fieldMap['item' + i + '_target'] = t ? { sheetKey: t.sheetKey, col: t.col } : null;
      fieldMap['item' + i + '_actual'] = a ? { sheetKey: a.sheetKey, col: a.col } : null;
    });
    template.extraCols.forEach(function (ec, i) {
      let bestCol = null, bestScore = 0;
      srcCols.forEach(function (sc) {
        const s = lcs(normText(ec.name), normText(sc.label));
        if (s > bestScore) { bestScore = s; bestCol = sc; }
      });
      fieldMap['extra' + i] = bestScore >= 3 ? { sheetKey: bestCol.sheetKey, col: bestCol.col } : null;
    });

    return { srcCols: srcCols, fieldMap: fieldMap, sheetKeyOf: sheetKeyOf };
  }

  // ---------- 生成结果 ----------
  // bindings: [{ templateCol, source:{sheetKey,col}|null }]
  function buildResult(template, dataSheets, bindings) {
    // 每个数据 sheet 各自建立 姓名 -> 行 的索引
    const nameIdxByKey = {};
    dataSheets.forEach(function (ds, idx) {
      const sk = sheetKeyOf(ds, idx);
      const m = {};
      ds.rows.forEach(function (row) {
        const k = normKey(row.key);
        if (k && !(k in m)) m[k] = row;
      });
      nameIdxByKey[sk] = m;
    });

    const updates = [];
    const matched = [];
    const unmatched = [];
    const seen = {};
    const rowResults = [];

    for (let r = template.dataStart; r <= template.lastRow; r++) {
      const name = templateName(template, r);
      const k = normKey(name);
      let anyHit = false;
      const cells = [];
      bindings.forEach(function (b) {
        let val = null;
        if (b.source && k) {
          const idx = nameIdxByKey[b.source.sheetKey];
          if (idx && idx[k]) { val = idx[k].values[b.source.col]; anyHit = true; }
        }
        val = toNumber(val);
        cells.push({ col: b.templateCol, value: val });
        if (val !== null) updates.push({ row: r, col: b.templateCol, addr: cellAddr(r, b.templateCol), value: val });
      });
      rowResults.push({ row: r, name: name, matched: anyHit, cells: cells });
      if (k && !seen[k]) {
        seen[k] = true;
        (anyHit ? matched : unmatched).push(name);
      }
    }

    return { updates: updates, rowResults: rowResults, matched: matched, unmatched: unmatched };
  }

  // 需要从模版网格取姓名——由调用方通过 template._grid 传入
  function templateName(template, r) {
    const g = template._grid;
    if (!g) return '';
    const row = g[r - 1];
    return row ? row[template.keyCol - 1] : '';
  }

  function attachGrid(template, wb) {
    const ex = expandGrid(wb.Sheets[template.sheetName]);
    template._grid = ex.grid;
    return template;
  }

  // 各列的表头标签（用于「按列拆分」的默认选项）
  function columnLabels(template) {
    const g = template._grid || [];
    const out = [];
    for (let c = 1; c <= template.maxCol; c++) {
      let label = '';
      for (let r = template.keyRow - 1; r >= 0; r--) {
        const row = g[r];
        const v = row ? row[c - 1] : null;
        if (v !== null && v !== undefined && String(v).trim() !== '') { label = String(v).trim(); break; }
      }
      out.push({ col: c, label: label });
    }
    return out;
  }

  // ---------- 写回 xlsx（只改目标 sheet 的 XML） ----------
  async function resolveSheetPath(zip, sheetName) {
    const wbxml = await zip.file('xl/workbook.xml').async('string');
    const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const sheets = [];
    let m, re = /<sheet\b[^>]*>/g;
    while ((m = re.exec(wbxml))) {
      const tag = m[0];
      const name = (tag.match(/name="([^"]*)"/) || [])[1];
      const rid = (tag.match(/r:id="([^"]*)"/) || [])[1];
      if (name !== undefined && rid) sheets.push({ name: decodeXml(name), rid: rid });
    }
    const relMap = {};
    re = /<Relationship\b[^>]*>/g;
    while ((m = re.exec(rels))) {
      const tag = m[0];
      const id = (tag.match(/Id="([^"]*)"/) || [])[1];
      const target = (tag.match(/Target="([^"]*)"/) || [])[1];
      if (id && target) relMap[id] = target;
    }
    const s = sheets.find(function (x) { return x.name === sheetName; });
    if (!s || !relMap[s.rid]) throw new Error('无法定位工作表 XML：' + sheetName);
    return 'xl/' + relMap[s.rid].replace(/^\/?/, '').replace(/^xl\//, '');
  }

  function buildColStyle(sheetXml) {
    const map = {};
    const re = /<c r="([A-Z]+)\d+"[^>]*?\bs="(\d+)"/g;
    let m;
    while ((m = re.exec(sheetXml))) { if (!(m[1] in map)) map[m[1]] = m[2]; }
    return map;
  }

  function insertCellOrdered(rowXml, colL, newCell) {
    const target = colNumber(colL);
    const re = /<c r="([A-Z]+)\d+"/g;
    let m, insertAt = -1;
    while ((m = re.exec(rowXml))) {
      if (colNumber(m[1]) > target) { insertAt = m.index; break; }
    }
    if (insertAt < 0) {
      const close = rowXml.indexOf('</row>');
      insertAt = close >= 0 ? close : rowXml.length;
    }
    return rowXml.slice(0, insertAt) + newCell + rowXml.slice(insertAt);
  }

  function setCellInRow(rowXml, addr, value, colStyle) {
    const colL = addr.match(/^[A-Z]+/)[0];
    const re = new RegExp('<c r="' + addr + '"((?:\\s[^>]*?)?)(/>|>)');
    const m = re.exec(rowXml);
    const valTxt = String(value);
    if (!m) {
      const attrs = ' s="' + (colStyle[colL] || '0') + '"';
      return insertCellOrdered(rowXml, colL, '<c r="' + addr + '"' + attrs + '><v>' + valTxt + '</v></c>');
    }
    let attrs = (m[1] || '').replace(/\st="[^"]*"/, '');
    if (m[2] === '/>') {
      return rowXml.slice(0, m.index) + '<c r="' + addr + '"' + attrs + '><v>' + valTxt + '</v></c>' + rowXml.slice(m.index + m[0].length);
    }
    const close = rowXml.indexOf('</c>', m.index);
    return rowXml.slice(0, m.index) + '<c r="' + addr + '"' + attrs + '><v>' + valTxt + '</v></c>' + rowXml.slice(close + 4);
  }

  function updateRow(sheetXml, rowNum, cellUpdates, colStyle) {
    const rowRe = new RegExp('<row\\b[^>]*\\br="' + rowNum + '"');
    const rm = rowRe.exec(sheetXml);
    if (!rm) {
      // 行不存在则新建，插入到 </sheetData> 之前
      let rowXml = '<row r="' + rowNum + '"></row>';
      cellUpdates.forEach(function (u) { rowXml = setCellInRow(rowXml, u.addr, u.value, colStyle); });
      const sd = sheetXml.lastIndexOf('</sheetData>');
      if (sd < 0) return sheetXml;
      return sheetXml.slice(0, sd) + rowXml + sheetXml.slice(sd);
    }
    const start = rm.index;
    const tagEnd = sheetXml.indexOf('>', start);
    const tag = sheetXml.slice(start, tagEnd + 1);
    const selfClose = /\/>$/.test(tag);
    const end = selfClose ? tagEnd + 1 : sheetXml.indexOf('</row>', start) + 6;
    let rowXml = selfClose ? tag.replace(/\/>$/, '></row>') : sheetXml.slice(start, end);
    cellUpdates.forEach(function (u) { rowXml = setCellInRow(rowXml, u.addr, u.value, colStyle); });
    return sheetXml.slice(0, start) + rowXml + sheetXml.slice(end);
  }

  // 把某一列公式单元格替换为 IFERROR(原公式,0)，避免空/0 目标导致 #DIV/0!
  // 共享公式只需改主单元格，其余成员自动继承
  function wrapFormulaColumn(sheetXml, colL) {
    const re = new RegExp('<c r="' + colL + '(\\d+)"([^>]*)>((?:(?!</c>)[\\s\\S])*?)</c>', 'g');
    return sheetXml.replace(re, function (m, row, cattrs, inner) {
      const fm = /<f([^>]*)>([\s\S]*?)<\/f>/.exec(inner);
      if (!fm) return m;
      if (/IFERROR/i.test(fm[2])) return m;
      const attrs = cattrs.replace(/\s*t="[^"]*"/, '');
      return '<c r="' + colL + row + '"' + attrs + '><f' + fm[1] + '>IFERROR(' + fm[2] + ',0)</f></c>';
    });
  }

  async function writeXlsx(templateBuffer, sheetName, updates, options) {
    options = options || {};
    const outType = options.outType || (typeof window !== 'undefined' ? 'blob' : 'nodebuffer');
    const zip = await JSZip.loadAsync(templateBuffer);
    const sheetPath = await resolveSheetPath(zip, sheetName);
    let sheetXml = await zip.file(sheetPath).async('string');
    const colStyle = buildColStyle(sheetXml);

    if (options.ratioCols && options.guardDiv !== false) {
      options.ratioCols.forEach(function (n) { sheetXml = wrapFormulaColumn(sheetXml, colLetter(n)); });
    }

    const byRow = {};
    updates.forEach(function (u) {
      const row = u.row || Number(u.addr.replace(/[A-Z]+/g, ''));
      (byRow[row] = byRow[row] || []).push({ addr: u.addr, value: u.value });
    });
    Object.keys(byRow).forEach(function (rowNum) {
      sheetXml = updateRow(sheetXml, Number(rowNum), byRow[rowNum], colStyle);
    });
    zip.file(sheetPath, sheetXml);

    // 强制打开时整体重算，确保公式列显示最新结果
    let wbxml = await zip.file('xl/workbook.xml').async('string');
    if (/<calcPr\b/.test(wbxml)) {
      wbxml = wbxml.replace(/<calcPr\b([^>]*?)(\/?)>/, function (m, attrs, sc) {
        return '<calcPr' + attrs.replace(/\s*fullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"' + (sc ? '/>' : '>');
      });
    } else {
      wbxml = wbxml.replace('</workbook>', '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>');
    }
    zip.file('xl/workbook.xml', wbxml);

    return await zip.generateAsync({ type: outType || (typeof window !== 'undefined' ? 'blob' : 'nodebuffer'), compression: 'DEFLATE' });
  }

  // ---------- 按列拆分（如按地市拆成多份表） ----------
  function escapeRegExp(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  // 公式内的相对引用按行/列偏移；带 $ 的绝对引用不移动
  function shiftRefs(formula, dRow, dCol) {
    return String(formula).replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g, function (m, d1, col, d2, row) {
      let r = parseInt(row, 10) + (d2 ? 0 : dRow);
      let c = colNumber(col.toUpperCase()) + (d1 ? 0 : dCol);
      if (r < 1 || c < 1) return m;
      return d1 + colLetter(c) + d2 + r;
    });
  }

  // 把共享公式(si)展开成每个单元格的独立公式，便于行删除/重排
  function expandSharedFormulas(xml) {
    const masters = {};
    let m;
    const masterRe = /<c r="([A-Z]+)(\d+)"([^>]*)>\s*<f t="shared"([^>]*?[^/>])>([\s\S]*?)<\/f>/g;
    while ((m = masterRe.exec(xml))) {
      const si = (m[4].match(/si="([^"]*)"/) || [])[1];
      if (si !== undefined) masters[si] = { col: m[1], row: parseInt(m[2], 10), formula: m[5] };
    }
    xml = xml.replace(/<c r="([A-Z]+)(\d+)"([^>]*)>\s*<f t="shared"([^>]*?[^/>])>([\s\S]*?)<\/f>/g,
      function (mm, col, row, attrs, sharedAttrs, formula) { return '<c r="' + col + row + '"' + attrs + '><f>' + formula + '</f>'; });
    xml = xml.replace(/<c r="([A-Z]+)(\d+)"([^>]*)>\s*<f t="shared"([^>]*?)\/>/g,
      function (mm, col, row, attrs, sharedAttrs) {
        const si = (sharedAttrs.match(/si="([^"]*)"/) || [])[1];
        const master = masters[si];
        if (!master) return mm;
        const f = shiftRefs(master.formula, parseInt(row, 10) - master.row, colNumber(col) - colNumber(master.col));
        return '<c r="' + col + row + '"' + attrs + '><f>' + f + '</f>';
      });
    return xml;
  }

  function shiftRowXml(rowXml, oldR, newR) {
    const delta = newR - oldR;
    rowXml = rowXml.replace(/(<row\b[^>]*?)\br="\d+"/, '$1r="' + newR + '"');
    rowXml = rowXml.replace(/<c r="([A-Z]+)\d+"/g, '<c r="$1' + newR + '"');
    rowXml = rowXml.replace(/(<f[^>]*>)([\s\S]*?)(<\/f>)/g, function (m, open, body, close) {
      return open + shiftRefs(body, delta, 0) + close;
    });
    if (delta !== 0) rowXml = rowXml.replace(/(<f[^>]*?\bref="[A-Z]+)\d+"/g, '$1' + newR + '"');
    return rowXml;
  }

  function buildSplitSheetXml(sheetXml, opts) {
    const headerEnd = opts.headerEnd, dataStart = opts.dataStart, keptRows = opts.keptRows;
    let xml = expandSharedFormulas(sheetXml);
    const rows = {};
    const rowRe = /<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>|<row\b[^>]*\br="(\d+)"[^>]*\/>/g;
    let m;
    while ((m = rowRe.exec(xml))) rows[parseInt(m[1] || m[2], 10)] = m[0];

    const newRef = {};
    let nr = dataStart;
    keptRows.forEach(function (r) { if (rows[r] && newRef[r] === undefined) { newRef[r] = nr; nr++; } });

    const parts = [];
    Object.keys(rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
      let n;
      if (r <= headerEnd) n = r;
      else if (newRef[r] !== undefined) n = newRef[r];
      else return;
      parts.push(shiftRowXml(rows[r], r, n));
    });

    const sdStart = xml.indexOf('<sheetData>');
    const sdEnd = xml.indexOf('</sheetData>');
    xml = xml.slice(0, sdStart) + '<sheetData>' + parts.join('') + '</sheetData>' + xml.slice(sdEnd + '</sheetData>'.length);

    let lastCol = 'AH';
    const dm = xml.match(/<dimension ref="[A-Z]+\d+:([A-Z]+)\d+"\/>/);
    if (dm) lastCol = dm[1];
    const newLast = nr - 1;
    xml = xml.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:' + lastCol + newLast + '"/>');
    xml = xml.replace(/<autoFilter ref="[^"]*"/, '<autoFilter ref="A3:' + lastCol + newLast + '"');
    return { xml: xml, lastCol: lastCol, newLast: newLast };
  }

  // resultBuffer: 已整合好的结果文件；groupCol: 分组列(1基)；返回 [{group, buffer}]
  async function splitByColumn(resultBuffer, sheetName, options) {
    options = options || {};
    const outType = options.outType || (typeof window !== 'undefined' ? 'blob' : 'nodebuffer');
    const headerEnd = options.headerEnd, dataStart = options.dataStart, lastRow = options.lastRow;
    const groupCol = options.groupCol;

    const wb = readWorkbook(new Uint8Array(resultBuffer), 'array');
    const ws = wb.Sheets[sheetName];
    const order = [];
    const byGroup = {};
    for (let r = dataStart; r <= lastRow; r++) {
      const cell = ws[XLSX.utils.encode_cell({ r: r - 1, c: groupCol - 1 })];
      const g = (cell && cell.v !== null && cell.v !== undefined) ? String(cell.v).trim() : '未分类';
      if (!byGroup[g]) { byGroup[g] = []; order.push(g); }
      byGroup[g].push(r);
    }

    const zip0 = await JSZip.loadAsync(resultBuffer);
    const sheetPath = await resolveSheetPath(zip0, sheetName);
    const sheetXml = await zip0.file(sheetPath).async('string');

    const out = [];
    for (const g of order) {
      const built = buildSplitSheetXml(sheetXml, { headerEnd: headerEnd, dataStart: dataStart, keptRows: byGroup[g] });
      const zip = await JSZip.loadAsync(resultBuffer);
      zip.file(sheetPath, built.xml);
      let wbxml = await zip.file('xl/workbook.xml').async('string');
      const nameRe = new RegExp(escapeRegExp(sheetName) + '!\\$A\\$3:\\$[A-Z]+\\$\\d+');
      wbxml = wbxml.replace(nameRe, sheetName + '!$A$3:$' + built.lastCol + '$' + built.newLast);
      zip.file('xl/workbook.xml', wbxml);
      const buf = await zip.generateAsync({ type: outType, compression: 'DEFLATE' });
      out.push({ group: g, buffer: buf, count: byGroup[g].length });
    }
    return out;
  }

  return {
    readWorkbook: readWorkbook,
    expandGrid: expandGrid,
    detectTemplate: detectTemplate,
    detectData: detectData,
    autoMap: autoMap,
    buildResult: buildResult,
    attachGrid: attachGrid,
    columnLabels: columnLabels,
    sheetKeyOf: sheetKeyOf,
    writeXlsx: writeXlsx,
    splitByColumn: splitByColumn,
    toNumber: toNumber,
    normKey: normKey,
    normText: normText,
    colLetter: colLetter,
    colNumber: colNumber,
    cellAddr: cellAddr
  };
}); 
