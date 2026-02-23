(() => {
  'use strict';

  const COURSEWORK_FIRST_ROW = 2;
  const EXAM_FIRST_ROW = 39;
  const MAX_TEMPLATE_CW = 4;

  const EXAM_MODULE_CODE_CELL = 'C3';
  const EXAM_MODULE_NAME_CELL = 'D3';
  const EXAM_ORGANISER_CELL = 'D6';
  const EXAM_CW_HEADER_FIRST_COL = 'D';
  const EXAM_CW_HEADER_ROW = 13;
  const EXAM_CW_WEIGHT_ROW = 14;

  const COURSEWORK_MODULE_CODE_CELL = 'A1';
  const COURSEWORK_CW_HEADER_FIRST_COL = 'B';
  const COURSEWORK_CW_HEADER_ROW = 1;

  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const NS_CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';

  const els = {
    buildExport: document.getElementById('build-export-file'),
    buildModuleCode: document.getElementById('build-module-code'),
    buildModuleName: document.getElementById('build-module-name'),
    buildModuleOrganiser: document.getElementById('build-module-organiser'),
    buildStudentIdCol: document.getElementById('build-student-id-col'),
    buildCwCols: document.getElementById('build-cw-cols'),
    buildBtn: document.getElementById('build-btn'),

    mergeExport: document.getElementById('merge-export-file'),
    mergeWorkbook: document.getElementById('merge-workbook-file'),
    mergeModuleCode: document.getElementById('merge-module-code'),
    mergeStudentIdCol: document.getElementById('merge-student-id-col'),
    mergeTargetCol: document.getElementById('merge-target-col'),
    mergeBlankUnmatched: document.getElementById('merge-blank-unmatched'),
    mergeBtn: document.getElementById('merge-btn'),

    log: document.getElementById('log'),
    clearLogBtn: document.getElementById('clear-log-btn'),
  };

  function timestamp() {
    const now = new Date();
    return now.toLocaleTimeString();
  }

  function log(message) {
    const lines = String(message).split(/\r?\n/);
    const stamped = lines.map((line) => `${timestamp()}  ${line}`).join('\n');
    els.log.textContent += `${stamped}\n`;
    els.log.scrollTop = els.log.scrollHeight;
  }

  function setBusy(busy) {
    els.buildBtn.disabled = busy;
    els.mergeBtn.disabled = busy;
  }

  function findFieldCaseInsensitive(fieldnames, wanted) {
    const wantedLower = String(wanted || '').trim().toLowerCase();
    if (!wantedLower) return null;
    for (const name of fieldnames) {
      if (String(name).trim().toLowerCase() === wantedLower) return name;
    }
    return null;
  }

  function normalizeRegNo(value) {
    if (value === null || value === undefined) return '';
    let text = String(value).trim();
    if (!text) return '';
    if (text.startsWith("'")) text = text.slice(1);
    if (text.startsWith('#')) text = text.slice(1);
    return text.trim();
  }

  function normalizeModuleCode(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/\s+/g, '').trim().toUpperCase();
  }

  function extractModuleCodeFromChildCourse(value) {
    const text = String(value || '').trim();
    if (!text) return '';
    if (text.includes('{') && text.includes('}')) {
      const pieces = [...text.matchAll(/\{([^}]*)\}/g)].map((m) => normalizeModuleCode(m[1]));
      const found = pieces.find(Boolean);
      if (found) return found;
    }
    return normalizeModuleCode(text);
  }

  function parseNumber(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === 'number') {
      return Number.isFinite(value) ? value : null;
    }

    let text = String(value).trim();
    if (!text) return null;
    text = text.replace(/,/g, '');
    if (text.endsWith('%')) text = text.slice(0, -1).trim();
    const num = Number(text);
    return Number.isFinite(num) ? num : null;
  }

  function formatMark(value) {
    const rounded = Math.round(Number(value) * 100) / 100;
    const fixed = rounded.toFixed(2);
    return fixed.replace(/\.00$/, '').replace(/(\.\d)0$/, '$1');
  }

  function formatRegForTemplate(value) {
    const norm = normalizeRegNo(value);
    return norm ? `#${norm}` : '';
  }

  function detectStudentIdColumn(fieldnames, preferred) {
    if (preferred) {
      const exact = findFieldCaseInsensitive(fieldnames, preferred);
      if (exact) return exact;
      throw new Error(`Student ID column '${preferred}' not found in headers`);
    }

    const patterns = [
      /\bstudent\s*id\b/i,
      /\bstudent\s*number\b/i,
      /\bregistration\s*number\b/i,
      /\bregistration\b/i,
      /\breg\s*no\b/i,
    ];

    for (const pattern of patterns) {
      for (const field of fieldnames) {
        if (pattern.test(field)) return field;
      }
    }

    throw new Error('Could not auto-detect Student ID column. Set it manually in Advanced options.');
  }

  function detectChildCourseColumn(fieldnames) {
    const patterns = [/\bchild\s*course\b/i, /\bchild\s*course\s*id\b/i];
    for (const pattern of patterns) {
      for (const field of fieldnames) {
        if (pattern.test(field)) return field;
      }
    }
    return null;
  }

  function parseCourseworkIndexForModule(field, moduleCode) {
    const match = String(field || '').match(/^\s*\{0*(\d{1,3})\}\s*\{([^}]+)\}/i);
    if (!match) return null;
    const idx = Number(match[1]);
    const mod = normalizeModuleCode(match[2]);
    if (mod !== normalizeModuleCode(moduleCode)) return null;
    return idx;
  }

  function detectCourseworkColumnsForModule(fieldnames, moduleCode, preferred) {
    if (preferred) {
      const rawItems = preferred
        .split(',')
        .map((x) => x.trim())
        .filter(Boolean);
      if (!rawItems.length) {
        throw new Error('CW columns override cannot be empty when provided.');
      }
      if (rawItems.length > MAX_TEMPLATE_CW) {
        throw new Error(`Template supports at most ${MAX_TEMPLATE_CW} coursework columns.`);
      }
      const resolved = [];
      for (const item of rawItems) {
        const exact = findFieldCaseInsensitive(fieldnames, item);
        if (!exact) throw new Error(`Coursework column '${item}' not found in headers`);
        resolved.push(exact);
      }
      return { columns: resolved, labels: [...rawItems] };
    }

    const detectedByIndex = new Map();
    for (const field of fieldnames) {
      const idx = parseCourseworkIndexForModule(field, moduleCode);
      if (idx !== null && !detectedByIndex.has(idx)) {
        detectedByIndex.set(idx, field);
      }
    }

    if (detectedByIndex.size) {
      const ordered = [...detectedByIndex.keys()].sort((a, b) => a - b);
      if (ordered.length > MAX_TEMPLATE_CW) {
        throw new Error(
          `Detected ${ordered.length} coursework items for ${moduleCode}, but template supports at most ${MAX_TEMPLATE_CW}.`
        );
      }
      return {
        columns: ordered.map((i) => detectedByIndex.get(i)),
        labels: ordered.map((i) => `CW ${String(i).padStart(3, '0')}`),
      };
    }

    throw new Error(
      `Could not auto-detect coursework columns for module '${moduleCode}'. ` +
      'Expected Blackboard headers like {001}{MODULECODE}. ' +
      'Use CW columns override if your site uses different header names.'
    );
  }

  function detectTargetColumnByModule(fieldnames, moduleCode, preferred) {
    if (preferred) {
      const exact = findFieldCaseInsensitive(fieldnames, preferred);
      if (exact) return exact;
      throw new Error(`Target column '${preferred}' not found in headers`);
    }

    const mod = normalizeModuleCode(moduleCode);
    if (!mod) throw new Error('Module code is required to auto-detect target column.');

    const candidates = fieldnames.filter((f) => normalizeModuleCode(f).includes(mod));
    if (!candidates.length) {
      throw new Error(`Could not find Blackboard target column containing module code '${moduleCode}'.`);
    }

    const starred = candidates.filter((f) => String(f).includes('**'));
    const pool = starred.length ? starred : candidates;

    const scored = pool.map((field, idx) => {
      let score = 0;
      if (/\*\*\s*0*\d+\b/i.test(field)) score += 3;
      if (/\b003\b/i.test(field)) score += 2;
      if (/\bdr\d+\b/i.test(field)) score -= 1;
      return { field, score, idx };
    });

    scored.sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      return a.idx - b.idx;
    });

    if (scored.length > 1 && scored[0].score === scored[1].score) {
      throw new Error(
        `Multiple target columns match module ${moduleCode}. Set Target column override explicitly.`
      );
    }

    return scored[0].field;
  }

  function extractCourseworkModuleCounts(fieldnames) {
    const counts = new Map();
    for (const field of fieldnames) {
      const m = String(field || '').match(/^\s*\{0*(\d{1,3})\}\s*\{([^}]+)\}/i);
      if (!m) continue;
      const module = normalizeModuleCode(m[2]);
      if (!module) continue;
      counts.set(module, (counts.get(module) || 0) + 1);
    }
    return counts;
  }

  function pickUniqueHighestByCount(candidates, counts, contextLabel) {
    if (!candidates.length) return '';
    let bestModule = '';
    let bestCount = -1;
    let tie = false;
    for (const module of candidates) {
      const count = counts.get(module) || 0;
      if (count > bestCount) {
        bestModule = module;
        bestCount = count;
        tie = false;
      } else if (count === bestCount) {
        tie = true;
      }
    }
    if (tie) {
      throw new Error(
        `Could not auto-infer module code (${contextLabel}) because multiple modules are equally likely. ` +
        `Set module code manually.`
      );
    }
    return bestModule;
  }

  function inferDefaultModuleCode(fieldnames, rows, childCourseCol) {
    const moduleCounts = extractCourseworkModuleCounts(fieldnames);
    const modules = [...moduleCounts.keys()];
    if (!modules.length) {
      throw new Error(
        'Could not auto-infer module code from Blackboard headers. ' +
        'Expected coursework headers like {001}{MODULECODE}.'
      );
    }

    if (modules.length === 1) return modules[0];

    const childModules = new Set();
    if (childCourseCol) {
      for (const row of rows) {
        const cm = extractModuleCodeFromChildCourse(row[childCourseCol]);
        if (cm) childModules.add(cm);
      }
    }

    const nonChild = modules.filter((m) => !childModules.has(m));
    if (nonChild.length === 1) return nonChild[0];
    if (nonChild.length > 1) {
      return pickUniqueHighestByCount(nonChild, moduleCounts, 'non-child module selection');
    }
    return pickUniqueHighestByCount(modules, moduleCounts, 'module selection');
  }

  function countDelimiterOutsideQuotes(line, delim) {
    let count = 0;
    let inQuotes = false;
    for (let i = 0; i < line.length; i += 1) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') {
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (!inQuotes && ch === delim) {
        count += 1;
      }
    }
    return count;
  }

  function detectDelimiter(text) {
    const firstLine = text.split(/\r?\n/).find((line) => line.trim().length > 0) || '';
    const candidates = [',', '\t', ';', '|'];
    let best = ',';
    let bestScore = -1;
    for (const delim of candidates) {
      const score = countDelimiterOutsideQuotes(firstLine, delim);
      if (score > bestScore) {
        best = delim;
        bestScore = score;
      }
    }
    return best;
  }

  function detectCsvEncoding(bytes) {
    if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return 'utf-16le';
    if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return 'utf-16be';

    const head = bytes.slice(0, 2048);
    let nulCount = 0;
    for (const b of head) if (b === 0) nulCount += 1;
    if (head.length && nulCount / head.length > 0.2) return 'utf-16le';

    try {
      // Validate UTF-8 strictly.
      new TextDecoder('utf-8', { fatal: true }).decode(head);
      return 'utf-8';
    } catch {
      return 'windows-1252';
    }
  }

  function decodeBytes(bytes, encoding) {
    let normalized = encoding;
    if (normalized === 'utf-16') normalized = 'utf-16le';
    if (normalized === 'utf-16be') {
      // Convert BE to LE bytes for reliable decoding.
      const swapped = new Uint8Array(bytes.length);
      for (let i = 0; i + 1 < bytes.length; i += 2) {
        swapped[i] = bytes[i + 1];
        swapped[i + 1] = bytes[i];
      }
      return new TextDecoder('utf-16le').decode(swapped);
    }
    return new TextDecoder(normalized).decode(bytes);
  }

  function encodeUtf16Le(text, includeBom = true) {
    const out = new Uint8Array((text.length * 2) + (includeBom ? 2 : 0));
    let offset = 0;
    if (includeBom) {
      out[0] = 0xff;
      out[1] = 0xfe;
      offset = 2;
    }
    for (let i = 0; i < text.length; i += 1) {
      const code = text.charCodeAt(i);
      out[offset++] = code & 0xff;
      out[offset++] = (code >> 8) & 0xff;
    }
    return out;
  }

  function encodeOutputText(text, sourceEncoding) {
    const enc = String(sourceEncoding || '').toLowerCase();
    if (enc.startsWith('utf-16')) return encodeUtf16Le(text, true);
    const utf8 = new TextEncoder().encode(text);
    const out = new Uint8Array(utf8.length + 3);
    out[0] = 0xef;
    out[1] = 0xbb;
    out[2] = 0xbf;
    out.set(utf8, 3);
    return out;
  }

  function parseDelimited(text, delimiter) {
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;

    for (let i = 0; i < text.length; i += 1) {
      const ch = text[i];

      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') {
            field += '"';
            i += 1;
          } else {
            inQuotes = false;
          }
        } else {
          field += ch;
        }
        continue;
      }

      if (ch === '"') {
        inQuotes = true;
        continue;
      }
      if (ch === delimiter) {
        row.push(field);
        field = '';
        continue;
      }
      if (ch === '\n') {
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
        continue;
      }
      if (ch === '\r') {
        if (text[i + 1] === '\n') i += 1;
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
        continue;
      }

      field += ch;
    }

    row.push(field);
    if (row.length > 1 || row[0] !== '' || rows.length === 0) rows.push(row);

    if (rows.length && rows[0].length) {
      rows[0][0] = String(rows[0][0]).replace(/^\uFEFF/, '');
    }

    return rows;
  }

  async function readBlackboardExport(file) {
    const arrayBuffer = await file.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    const encoding = detectCsvEncoding(bytes);
    const text = decodeBytes(bytes, encoding);
    const delimiter = detectDelimiter(text);
    const matrix = parseDelimited(text, delimiter);
    if (!matrix.length) throw new Error('No header row found in Blackboard export.');

    const headers = matrix[0].map((h) => String(h || '').trim());
    if (!headers.some(Boolean)) throw new Error('Header row is empty in Blackboard export.');

    const rows = [];
    for (let i = 1; i < matrix.length; i += 1) {
      const arr = matrix[i];
      const rowObj = {};
      let hasAny = false;
      for (let j = 0; j < headers.length; j += 1) {
        const key = headers[j];
        const val = arr[j] !== undefined ? String(arr[j]) : '';
        rowObj[key] = val;
        if (val.trim() !== '') hasAny = true;
      }
      if (hasAny) rows.push(rowObj);
    }

    return {
      headers,
      rows,
      delimiter,
      encoding,
      filename: file.name,
    };
  }

  function csvEscape(value, delimiter) {
    const text = value === null || value === undefined ? '' : String(value);
    if (text.includes('"') || text.includes('\n') || text.includes('\r') || text.includes(delimiter)) {
      return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
  }

  function serializeDelimited(headers, rows, delimiter) {
    const lines = [];
    lines.push(headers.map((h) => csvEscape(h, delimiter)).join(delimiter));
    for (const row of rows) {
      lines.push(headers.map((h) => csvEscape(row[h] ?? '', delimiter)).join(delimiter));
    }
    return lines.join('\r\n');
  }

  function base64ToBytes(base64) {
    const binary = atob(base64);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) out[i] = binary.charCodeAt(i);
    return out;
  }

  function loadTemplateWorkbook() {
    if (!window.TEMPLATE_XLSX_BASE64) {
      throw new Error('Embedded template not found.');
    }
    const bytes = base64ToBytes(window.TEMPLATE_XLSX_BASE64);
    return XLSX.read(bytes, {
      type: 'array',
      cellFormula: true,
      cellStyles: true,
      cellNF: true,
      raw: false,
    });
  }

  async function loadTemplateZip() {
    if (!window.TEMPLATE_XLSX_BASE64) {
      throw new Error('Embedded template not found.');
    }
    if (!window.JSZip) {
      throw new Error('JSZip is required but not loaded.');
    }
    const bytes = base64ToBytes(window.TEMPLATE_XLSX_BASE64);
    return window.JSZip.loadAsync(bytes);
  }

  function parseXml(xmlText) {
    const doc = new DOMParser().parseFromString(xmlText, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) {
      throw new Error('Failed to parse workbook XML.');
    }
    return doc;
  }

  function serializeXml(doc) {
    const xml = new XMLSerializer().serializeToString(doc);
    if (/^\s*<\?xml\b/i.test(xml)) {
      return xml;
    }
    return `<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n${xml}`;
  }

  function normalizeSheetPart(target) {
    let clean = String(target || '').trim();
    if (!clean) return '';
    clean = clean.replace(/^\/+/, '');
    if (!clean.startsWith('xl/')) clean = `xl/${clean}`;
    clean = clean.replace(/\/{2,}/g, '/');
    return clean;
  }

  function splitCellRef(ref) {
    const m = String(ref || '').match(/^([A-Za-z]+)(\d+)$/);
    if (!m) throw new Error(`Invalid cell reference: ${ref}`);
    return { col: m[1].toUpperCase(), row: Number(m[2]) };
  }

  function colToIdx(col) {
    return XLSX.utils.decode_col(String(col).toUpperCase()) + 1;
  }

  function idxToCol(idx) {
    return XLSX.utils.encode_col(idx - 1);
  }

  function createWorksheetEditor(doc) {
    const sheetData = doc.getElementsByTagNameNS(NS_MAIN, 'sheetData')[0];
    if (!sheetData) throw new Error('Worksheet is missing sheetData.');

    const rowMap = new Map();
    for (const row of Array.from(sheetData.children)) {
      if (row.localName !== 'row') continue;
      const idx = Number(row.getAttribute('r') || '0');
      if (idx > 0) rowMap.set(idx, row);
    }

    function maxRow() {
      if (!rowMap.size) return 0;
      return Math.max(...rowMap.keys());
    }

    function getRow(rowIdx) {
      return rowMap.get(rowIdx) || null;
    }

    function ensureRow(rowIdx) {
      const existing = rowMap.get(rowIdx);
      if (existing) return existing;

      const row = doc.createElementNS(NS_MAIN, 'row');
      row.setAttribute('r', String(rowIdx));

      let inserted = false;
      for (const child of Array.from(sheetData.children)) {
        if (child.localName !== 'row') continue;
        const idx = Number(child.getAttribute('r') || '0');
        if (idx > rowIdx) {
          sheetData.insertBefore(row, child);
          inserted = true;
          break;
        }
      }
      if (!inserted) sheetData.appendChild(row);

      rowMap.set(rowIdx, row);
      return row;
    }

    function findCellInRow(row, ref) {
      for (const cell of Array.from(row.children)) {
        if (cell.localName === 'c' && cell.getAttribute('r') === ref) return cell;
      }
      return null;
    }

    function ensureCell(rowIdx, col) {
      const ref = `${col.toUpperCase()}${rowIdx}`;
      const row = ensureRow(rowIdx);
      const found = findCellInRow(row, ref);
      if (found) return found;

      const cell = doc.createElementNS(NS_MAIN, 'c');
      cell.setAttribute('r', ref);
      const targetColIdx = colToIdx(col);

      let inserted = false;
      for (const existing of Array.from(row.children)) {
        if (existing.localName !== 'c') continue;
        const existingRef = existing.getAttribute('r') || '';
        const m = existingRef.match(/^([A-Za-z]+)\d+$/);
        if (!m) continue;
        const existingColIdx = colToIdx(m[1]);
        if (existingColIdx > targetColIdx) {
          row.insertBefore(cell, existing);
          inserted = true;
          break;
        }
      }
      if (!inserted) row.appendChild(cell);
      return cell;
    }

    function clearValueLike(cell, removeFormula) {
      for (const child of Array.from(cell.children)) {
        const name = child.localName;
        if (name === 'v' || name === 'is' || (removeFormula && name === 'f')) {
          cell.removeChild(child);
        }
      }
      cell.removeAttribute('t');
    }

    function clearCell(rowIdx, col) {
      const cell = ensureCell(rowIdx, col);
      clearValueLike(cell, true);
    }

    function setCellValue(rowIdx, col, value) {
      const cell = ensureCell(rowIdx, col);
      clearValueLike(cell, true);

      if (value === null || value === undefined) return;

      if (typeof value === 'string') {
        if (!value.trim()) return;
        cell.setAttribute('t', 'inlineStr');
        const isEl = doc.createElementNS(NS_MAIN, 'is');
        const tEl = doc.createElementNS(NS_MAIN, 't');
        tEl.textContent = value;
        if (value !== value.trim()) {
          tEl.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        }
        isEl.appendChild(tEl);
        cell.appendChild(isEl);
        return;
      }

      const number = parseNumber(value);
      if (number === null) return;
      const vEl = doc.createElementNS(NS_MAIN, 'v');
      vEl.textContent = Number.isInteger(number) ? String(number) : String(number);
      cell.appendChild(vEl);
    }

    function clearFormulaCachedValues() {
      const cells = doc.getElementsByTagNameNS(NS_MAIN, 'c');
      for (const cell of Array.from(cells)) {
        const f = Array.from(cell.children).find((x) => x.localName === 'f');
        if (!f) continue;
        const v = Array.from(cell.children).find((x) => x.localName === 'v');
        if (v) cell.removeChild(v);
      }
    }

    return { doc, maxRow, clearCell, setCellValue, clearFormulaCachedValues };
  }

  async function openTemplateXmlWorkbook() {
    const zip = await loadTemplateZip();
    const workbookDoc = parseXml(await zip.file('xl/workbook.xml').async('string'));
    const workbookRelsDoc = parseXml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
    const contentTypesDoc = parseXml(await zip.file('[Content_Types].xml').async('string'));

    const worksheetEditors = new Map();
    let sheetMap = new Map();

    function relTargetsById() {
      const map = new Map();
      const rels = workbookRelsDoc.getElementsByTagNameNS(NS_PKG_REL, 'Relationship');
      for (const rel of Array.from(rels)) {
        const id = rel.getAttribute('Id') || '';
        const target = rel.getAttribute('Target') || '';
        if (id && target) map.set(id, target);
      }
      return map;
    }

    function refreshSheetMap() {
      const relMap = relTargetsById();
      const next = new Map();
      const sheets = workbookDoc.getElementsByTagNameNS(NS_MAIN, 'sheet');
      for (const sheet of Array.from(sheets)) {
        const name = sheet.getAttribute('name') || '';
        const rid = sheet.getAttributeNS(NS_DOC_REL, 'id') || sheet.getAttribute('r:id') || '';
        if (!name || !rid) continue;
        const target = relMap.get(rid);
        if (!target) continue;
        next.set(name, normalizeSheetPart(target));
      }
      sheetMap = next;
    }

    function sheetPartByName(name) {
      refreshSheetMap();
      const part = sheetMap.get(name);
      if (!part) throw new Error(`Sheet '${name}' not found in workbook.`);
      return part;
    }

    async function worksheetEditorByName(name) {
      const part = sheetPartByName(name);
      if (!worksheetEditors.has(part)) {
        const xml = await zip.file(part).async('string');
        worksheetEditors.set(part, createWorksheetEditor(parseXml(xml)));
      }
      return worksheetEditors.get(part);
    }

    function nextSheetPartNumber() {
      let maxN = 0;
      for (const key of Object.keys(zip.files)) {
        const m = key.match(/^xl\/worksheets\/sheet(\d+)\.xml$/);
        if (!m) continue;
        const n = Number(m[1]);
        if (n > maxN) maxN = n;
      }
      return maxN + 1;
    }

    function nextRid() {
      let maxN = 0;
      const rels = workbookRelsDoc.getElementsByTagNameNS(NS_PKG_REL, 'Relationship');
      for (const rel of Array.from(rels)) {
        const id = rel.getAttribute('Id') || '';
        const m = id.match(/^rId(\d+)$/);
        if (!m) continue;
        const n = Number(m[1]);
        if (n > maxN) maxN = n;
      }
      return `rId${maxN + 1}`;
    }

    function nextSheetId() {
      let maxN = 0;
      const sheets = workbookDoc.getElementsByTagNameNS(NS_MAIN, 'sheet');
      for (const sheet of Array.from(sheets)) {
        const n = Number(sheet.getAttribute('sheetId') || '0');
        if (n > maxN) maxN = n;
      }
      return maxN + 1;
    }

    function ensureContentTypeOverride(partName) {
      const overrides = contentTypesDoc.getElementsByTagNameNS(NS_CONTENT_TYPES, 'Override');
      for (const ov of Array.from(overrides)) {
        if ((ov.getAttribute('PartName') || '') === `/${partName}`) return;
      }
      const override = contentTypesDoc.createElementNS(NS_CONTENT_TYPES, 'Override');
      override.setAttribute('PartName', `/${partName}`);
      override.setAttribute(
        'ContentType',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
      );
      contentTypesDoc.documentElement.appendChild(override);
    }

    async function cloneSheet(sourceName, newName, rewriteCourseworkRefs) {
      refreshSheetMap();
      if (sheetMap.has(newName)) return;

      const sourcePart = sheetPartByName(sourceName);
      let sourceXml = await zip.file(sourcePart).async('string');
      if (rewriteCourseworkRefs) {
        sourceXml = sourceXml.replace(/'Coursework'!/g, "'Coursework Child'!");
        sourceXml = sourceXml.replace(/(^|[^A-Za-z0-9_'])Coursework!/g, "$1'Coursework Child'!");
      }

      const partNumber = nextSheetPartNumber();
      const newPart = `xl/worksheets/sheet${partNumber}.xml`;
      zip.file(newPart, sourceXml);

      const rid = nextRid();
      const rel = workbookRelsDoc.createElementNS(NS_PKG_REL, 'Relationship');
      rel.setAttribute('Id', rid);
      rel.setAttribute(
        'Type',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
      );
      rel.setAttribute('Target', `worksheets/sheet${partNumber}.xml`);
      workbookRelsDoc.documentElement.appendChild(rel);

      const sheetsParent = workbookDoc.getElementsByTagNameNS(NS_MAIN, 'sheets')[0];
      if (!sheetsParent) throw new Error('Workbook is missing <sheets>.');
      const newSheet = workbookDoc.createElementNS(NS_MAIN, 'sheet');
      newSheet.setAttribute('name', newName);
      newSheet.setAttribute('sheetId', String(nextSheetId()));
      newSheet.setAttributeNS(NS_DOC_REL, 'r:id', rid);
      sheetsParent.appendChild(newSheet);

      ensureContentTypeOverride(newPart);
      refreshSheetMap();
    }

    async function setForceRecalcOnOpen() {
      let calcPr = workbookDoc.getElementsByTagNameNS(NS_MAIN, 'calcPr')[0];
      if (!calcPr) {
        calcPr = workbookDoc.createElementNS(NS_MAIN, 'calcPr');
        workbookDoc.documentElement.appendChild(calcPr);
      }
      calcPr.setAttribute('fullCalcOnLoad', '1');
      calcPr.setAttribute('forceFullCalc', '1');
    }

    async function saveAsUint8Array() {
      zip.file('xl/workbook.xml', serializeXml(workbookDoc));
      zip.file('xl/_rels/workbook.xml.rels', serializeXml(workbookRelsDoc));
      zip.file('[Content_Types].xml', serializeXml(contentTypesDoc));
      for (const [part, editor] of worksheetEditors.entries()) {
        zip.file(part, serializeXml(editor.doc));
      }
      return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    }

    refreshSheetMap();
    return {
      worksheetEditorByName,
      cloneSheet,
      setForceRecalcOnOpen,
      saveAsUint8Array,
    };
  }

  function setSheetMetadataXml(examEditor, courseworkEditor, moduleCode, moduleName, moduleOrganiser) {
    const examMod = splitCellRef(EXAM_MODULE_CODE_CELL);
    examEditor.setCellValue(examMod.row, examMod.col, moduleCode);

    const examName = splitCellRef(EXAM_MODULE_NAME_CELL);
    examEditor.setCellValue(examName.row, examName.col, moduleName || '');

    const examOrg = splitCellRef(EXAM_ORGANISER_CELL);
    examEditor.setCellValue(examOrg.row, examOrg.col, moduleOrganiser || '');

    const cwMod = splitCellRef(COURSEWORK_MODULE_CODE_CELL);
    courseworkEditor.setCellValue(cwMod.row, cwMod.col, moduleCode);
  }

  function setCourseworkLayoutXml(examEditor, courseworkEditor, labels) {
    const examStart = colToIdx(EXAM_CW_HEADER_FIRST_COL);
    const cwStart = colToIdx(COURSEWORK_CW_HEADER_FIRST_COL);
    for (let i = 0; i < MAX_TEMPLATE_CW; i += 1) {
      const label = i < labels.length ? labels[i] : '';
      const examCol = idxToCol(examStart + i);
      const cwCol = idxToCol(cwStart + i);
      examEditor.setCellValue(EXAM_CW_HEADER_ROW, examCol, label);
      courseworkEditor.setCellValue(COURSEWORK_CW_HEADER_ROW, cwCol, label);
    }
  }

  async function populateCohortSheetsXml(builder, cohort, moduleName, moduleOrganiser) {
    const courseworkEditor = await builder.worksheetEditorByName(cohort.courseworkSheetName);
    const examEditor = await builder.worksheetEditorByName(cohort.examSheetName);

    setSheetMetadataXml(examEditor, courseworkEditor, cohort.moduleCode, moduleName, moduleOrganiser);
    setCourseworkLayoutXml(examEditor, courseworkEditor, cohort.courseworkLabels);

    const capacity = courseworkEditor.maxRow() - COURSEWORK_FIRST_ROW + 1;
    if (capacity <= 0) {
      throw new Error(`Template sheet '${cohort.courseworkSheetName}' has no student rows.`);
    }
    if (cohort.students.length > capacity) {
      throw new Error(
        `${cohort.name} cohort has ${cohort.students.length} students but template capacity is ${capacity}.`
      );
    }

    for (let idx = 0; idx < capacity; idx += 1) {
      const cwRow = COURSEWORK_FIRST_ROW + idx;
      const examRow = EXAM_FIRST_ROW + idx;
      const student = cohort.students[idx] || null;

      if (student) {
        courseworkEditor.setCellValue(cwRow, 'A', formatRegForTemplate(student.studentId));
        for (let cwIdx = 0; cwIdx < MAX_TEMPLATE_CW; cwIdx += 1) {
          const col = idxToCol(colToIdx('B') + cwIdx);
          courseworkEditor.setCellValue(cwRow, col, student.coursework[cwIdx]);
        }
      } else {
        courseworkEditor.clearCell(cwRow, 'A');
        for (let cwIdx = 0; cwIdx < MAX_TEMPLATE_CW; cwIdx += 1) {
          const col = idxToCol(colToIdx('B') + cwIdx);
          courseworkEditor.clearCell(cwRow, col);
        }
      }

      for (const qCol of ['D', 'E', 'F', 'G', 'H', 'I']) {
        examEditor.clearCell(examRow, qCol);
      }
    }

    courseworkEditor.clearFormulaCachedValues();
    examEditor.clearFormulaCachedValues();
  }

  function replaceExt(filename, newExtWithDot) {
    const idx = filename.lastIndexOf('.');
    const base = idx >= 0 ? filename.slice(0, idx) : filename;
    return `${base}${newExtWithDot}`;
  }

  function uploadNameFromExport(filename) {
    const idx = filename.lastIndexOf('.');
    const base = idx >= 0 ? filename.slice(0, idx) : filename;
    return `${base}_upload.xls`;
  }

  function downloadBytes(bytes, filename, mimeType) {
    const blob = new Blob([bytes], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function getSheet(workbook, name) {
    const ws = workbook.Sheets[name];
    if (!ws) throw new Error(`Sheet '${name}' not found in workbook.`);
    return ws;
  }

  function sheetMaxRow(ws) {
    const ref = ws['!ref'];
    if (!ref) return 1;
    return XLSX.utils.decode_range(ref).e.r + 1;
  }

  function updateSheetRef(ws, addr) {
    const cell = XLSX.utils.decode_cell(addr);
    if (!ws['!ref']) {
      ws['!ref'] = `${addr}:${addr}`;
      return;
    }
    const range = XLSX.utils.decode_range(ws['!ref']);
    if (cell.r < range.s.r) range.s.r = cell.r;
    if (cell.c < range.s.c) range.s.c = cell.c;
    if (cell.r > range.e.r) range.e.r = cell.r;
    if (cell.c > range.e.c) range.e.c = cell.c;
    ws['!ref'] = XLSX.utils.encode_range(range);
  }

  function getCellValue(ws, addr) {
    const cell = ws[addr];
    if (!cell) return '';
    if (cell.v === undefined || cell.v === null) return '';
    return cell.v;
  }

  function clearCell(ws, addr) {
    const cell = ws[addr];
    if (!cell) return;
    if (cell.f) {
      delete cell.v;
      delete cell.w;
      return;
    }
    cell.t = 's';
    cell.v = '';
    delete cell.w;
  }

  function setCellValue(ws, addr, value) {
    if (value === null || value === undefined || value === '') {
      clearCell(ws, addr);
      return;
    }

    const existing = ws[addr] ? { ...ws[addr] } : {};
    delete existing.f;
    delete existing.w;

    if (typeof value === 'number' && Number.isFinite(value)) {
      existing.t = 'n';
      existing.v = value;
    } else {
      existing.t = 's';
      existing.v = String(value);
    }

    ws[addr] = existing;
    updateSheetRef(ws, addr);
  }

  function clearFormulaCachedValues(ws) {
    for (const key of Object.keys(ws)) {
      if (key.startsWith('!')) continue;
      const cell = ws[key];
      if (cell && typeof cell.f === 'string') {
        delete cell.v;
        delete cell.w;
      }
    }
  }

  function colOffset(startCol, offset) {
    const idx = XLSX.utils.decode_col(startCol) + offset;
    return XLSX.utils.encode_col(idx);
  }

  function cloneSheet(workbook, sourceName, newName, rewriteCourseworkRefs = false) {
    if (workbook.Sheets[newName]) return;
    const sourceSheet = workbook.Sheets[sourceName];
    if (!sourceSheet) throw new Error(`Cannot clone missing sheet '${sourceName}'.`);

    const sourceIndex = workbook.SheetNames.indexOf(sourceName);
    const clone = JSON.parse(JSON.stringify(sourceSheet));

    if (rewriteCourseworkRefs) {
      for (const key of Object.keys(clone)) {
        if (key.startsWith('!')) continue;
        const cell = clone[key];
        if (!cell || typeof cell.f !== 'string') continue;
        let f = cell.f;
        f = f.replace(/'Coursework'!/g, "'Coursework Child'!");
        f = f.replace(/(^|[^A-Za-z0-9_'])Coursework!/g, "$1'Coursework Child'!");
        cell.f = f;
      }
    }

    workbook.Sheets[newName] = clone;
    workbook.SheetNames.push(newName);

    if (workbook.Workbook && Array.isArray(workbook.Workbook.Sheets)) {
      const sourceMeta = workbook.Workbook.Sheets[sourceIndex] || { name: sourceName };
      const metaClone = JSON.parse(JSON.stringify(sourceMeta));
      metaClone.name = newName;
      workbook.Workbook.Sheets.push(metaClone);
    }
  }

  function setSheetMetadata(examSheet, courseworkSheet, moduleCode, moduleName, moduleOrganiser) {
    setCellValue(examSheet, EXAM_MODULE_CODE_CELL, moduleCode);
    if (moduleName !== null && moduleName !== undefined) {
      setCellValue(examSheet, EXAM_MODULE_NAME_CELL, moduleName);
    }
    if (moduleOrganiser !== null && moduleOrganiser !== undefined) {
      setCellValue(examSheet, EXAM_ORGANISER_CELL, moduleOrganiser);
    }
    setCellValue(courseworkSheet, COURSEWORK_MODULE_CODE_CELL, moduleCode);
  }

  function setCourseworkLayout(examSheet, courseworkSheet, labels) {
    for (let i = 0; i < MAX_TEMPLATE_CW; i += 1) {
      const label = i < labels.length ? labels[i] : '';
      const examCol = colOffset(EXAM_CW_HEADER_FIRST_COL, i);
      const cwCol = colOffset(COURSEWORK_CW_HEADER_FIRST_COL, i);

      setCellValue(examSheet, `${examCol}${EXAM_CW_HEADER_ROW}`, label);
      setCellValue(courseworkSheet, `${cwCol}${COURSEWORK_CW_HEADER_ROW}`, label);
    }
  }

  function buildStudentRecords(rows, studentIdCol, cwCols) {
    const records = [];
    for (const row of rows) {
      const sid = normalizeRegNo(row[studentIdCol]);
      if (!sid) continue;
      const coursework = [];
      for (const col of cwCols) coursework.push(parseNumber(row[col]));
      while (coursework.length < MAX_TEMPLATE_CW) coursework.push(null);
      records.push({ studentId: sid, coursework: coursework.slice(0, MAX_TEMPLATE_CW) });
    }
    return records;
  }

  function populateCohortSheets(workbook, cohort, moduleName, moduleOrganiser) {
    const courseworkSheet = getSheet(workbook, cohort.courseworkSheetName);
    const examSheet = getSheet(workbook, cohort.examSheetName);

    setSheetMetadata(examSheet, courseworkSheet, cohort.moduleCode, moduleName, moduleOrganiser);
    setCourseworkLayout(examSheet, courseworkSheet, cohort.courseworkLabels);

    const capacity = sheetMaxRow(courseworkSheet) - COURSEWORK_FIRST_ROW + 1;
    if (capacity <= 0) {
      throw new Error(`Template sheet '${cohort.courseworkSheetName}' has no student rows.`);
    }
    if (cohort.students.length > capacity) {
      throw new Error(
        `${cohort.name} cohort has ${cohort.students.length} students but template capacity is ${capacity}.`
      );
    }

    for (let idx = 0; idx < capacity; idx += 1) {
      const cwRow = COURSEWORK_FIRST_ROW + idx;
      const examRow = EXAM_FIRST_ROW + idx;
      const student = cohort.students[idx] || null;

      if (student) {
        setCellValue(courseworkSheet, `A${cwRow}`, formatRegForTemplate(student.studentId));
        for (let cwIdx = 0; cwIdx < MAX_TEMPLATE_CW; cwIdx += 1) {
          const col = colOffset('B', cwIdx);
          const mark = student.coursework[cwIdx];
          setCellValue(courseworkSheet, `${col}${cwRow}`, mark);
        }
      } else {
        clearCell(courseworkSheet, `A${cwRow}`);
        for (let cwIdx = 0; cwIdx < MAX_TEMPLATE_CW; cwIdx += 1) {
          const col = colOffset('B', cwIdx);
          clearCell(courseworkSheet, `${col}${cwRow}`);
        }
      }

      for (const qCol of ['D', 'E', 'F', 'G', 'H', 'I']) {
        clearCell(examSheet, `${qCol}${examRow}`);
      }
    }

    clearFormulaCachedValues(courseworkSheet);
    clearFormulaCachedValues(examSheet);
  }

  function extractMarksWithModulesFromWorkbook(workbook) {
    const pairs = [['Coursework', 'Exam']];
    const hasCwChild = Boolean(workbook.Sheets['Coursework Child']);
    const hasExamChild = Boolean(workbook.Sheets['Exam Child']);
    if (hasCwChild && hasExamChild) {
      pairs.push(['Coursework Child', 'Exam Child']);
    } else if (hasCwChild || hasExamChild) {
      throw new Error("Workbook has only one child sheet. Expected both 'Coursework Child' and 'Exam Child'.");
    }

    const marks = new Map();

    for (const [cwName, examName] of pairs) {
      const cwSheet = getSheet(workbook, cwName);
      const examSheet = getSheet(workbook, examName);
      const moduleCode = normalizeModuleCode(getCellValue(examSheet, EXAM_MODULE_CODE_CELL));
      if (!moduleCode) {
        throw new Error(`Module code missing in ${examName}!${EXAM_MODULE_CODE_CELL}`);
      }

      const maxRows = Math.max(
        sheetMaxRow(cwSheet) - COURSEWORK_FIRST_ROW + 1,
        sheetMaxRow(examSheet) - EXAM_FIRST_ROW + 1
      );

      for (let idx = 0; idx < maxRows; idx += 1) {
        const cwRow = COURSEWORK_FIRST_ROW + idx;
        const examRow = EXAM_FIRST_ROW + idx;

        const sid = normalizeRegNo(getCellValue(cwSheet, `A${cwRow}`));
        if (!sid) continue;

        const markNum = parseNumber(getCellValue(examSheet, `M${examRow}`));
        if (markNum === null) continue;

        const mark = formatMark(markNum);
        const key = `${sid}||${moduleCode}`;
        const prev = marks.get(key);
        if (prev && prev !== mark) {
          throw new Error(`Conflicting marks for Student ID ${sid} in module ${moduleCode}.`);
        }
        marks.set(key, mark);
      }
    }

    if (!marks.size) {
      throw new Error('No moderated marks found in workbook. Ensure Exam column M is populated.');
    }

    return marks;
  }

  function splitRowsByChildCourse(rows, childCourseCol) {
    if (!childCourseCol) return { mainRows: [...rows], childRows: [], childModuleCodes: new Set() };

    const mainRows = [];
    const childRows = [];
    const childModuleCodes = new Set();

    for (const row of rows) {
      const childModule = extractModuleCodeFromChildCourse(row[childCourseCol]);
      if (childModule) {
        childRows.push(row);
        childModuleCodes.add(childModule);
      } else {
        mainRows.push(row);
      }
    }

    return { mainRows, childRows, childModuleCodes };
  }

  async function runBuild() {
    const exportFile = els.buildExport.files[0];
    if (!exportFile) throw new Error('Blackboard export file is required.');

    let moduleCode = normalizeModuleCode(els.buildModuleCode.value);

    const moduleName = els.buildModuleName.value.trim();
    const moduleOrganiser = els.buildModuleOrganiser.value.trim();

    log('[build] reading Blackboard export');
    const exportData = await readBlackboardExport(exportFile);

    const studentIdCol = detectStudentIdColumn(exportData.headers, els.buildStudentIdCol.value.trim());
    const childCourseCol = detectChildCourseColumn(exportData.headers);

    const filteredRows = exportData.rows.filter((row) => normalizeRegNo(row[studentIdCol]));
    const { mainRows, childRows, childModuleCodes } = splitRowsByChildCourse(filteredRows, childCourseCol);

    if (!moduleCode) {
      moduleCode = inferDefaultModuleCode(exportData.headers, filteredRows, childCourseCol);
      els.buildModuleCode.value = moduleCode;
      log(`[build] inferred module code: ${moduleCode}`);
    }

    const mainCw = detectCourseworkColumnsForModule(
      exportData.headers,
      moduleCode,
      els.buildCwCols.value.trim()
    );
    const mainStudents = buildStudentRecords(mainRows, studentIdCol, mainCw.columns);

    const cohorts = [
      {
        name: 'Main',
        courseworkSheetName: 'Coursework',
        examSheetName: 'Exam',
        moduleCode,
        students: mainStudents,
        courseworkColumns: mainCw.columns,
        courseworkLabels: mainCw.labels,
      },
    ];

    if (childRows.length) {
      if (!childModuleCodes.size) {
        throw new Error('Child course rows exist but child module code could not be extracted.');
      }
      if (childModuleCodes.size > 1) {
        throw new Error(
          `Multiple child module codes found in one export: ${[...childModuleCodes].join(', ')}`
        );
      }
      const childModule = [...childModuleCodes][0];
      const childCw = detectCourseworkColumnsForModule(exportData.headers, childModule, '');
      const childStudents = buildStudentRecords(childRows, studentIdCol, childCw.columns);
      cohorts.push({
        name: 'Child',
        courseworkSheetName: 'Coursework Child',
        examSheetName: 'Exam Child',
        moduleCode: childModule,
        students: childStudents,
        courseworkColumns: childCw.columns,
        courseworkLabels: childCw.labels,
      });
    }

    const totalStudents = cohorts.reduce((sum, c) => sum + c.students.length, 0);
    if (totalStudents === 0) throw new Error('No student records with Student ID were found.');

    log('[build] creating workbook from embedded template');
    const builder = await openTemplateXmlWorkbook();

    const needChild = cohorts.some((c) => c.name === 'Child');
    if (needChild) {
      await builder.cloneSheet('Coursework', 'Coursework Child', false);
      await builder.cloneSheet('Exam', 'Exam Child', true);
    }

    for (const cohort of cohorts) {
      await populateCohortSheetsXml(builder, cohort, moduleName, moduleOrganiser);
    }

    await builder.setForceRecalcOnOpen();
    const output = await builder.saveAsUint8Array();

    const outName = replaceExt(exportFile.name, '.xlsx');
    downloadBytes(output, outName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    log(`[build] wrote workbook: ${outName}`);
    log(`[build] student id column: ${studentIdCol}`);
    if (childCourseCol) log(`[build] child course column: ${childCourseCol}`);
    for (const cohort of cohorts) {
      log(
        `[build] ${cohort.name} module=${cohort.moduleCode} courseworks=${cohort.courseworkColumns.length} students=${cohort.students.length}`
      );
    }
  }

  async function runMerge() {
    const exportFile = els.mergeExport.files[0];
    const workbookFile = els.mergeWorkbook.files[0];
    if (!exportFile) throw new Error('Original Blackboard export file is required.');
    if (!workbookFile) throw new Error('Filled moderation workbook is required.');

    log('[merge] reading Blackboard export');
    const exportData = await readBlackboardExport(exportFile);

    log('[merge] reading filled workbook');
    const wbBytes = await workbookFile.arrayBuffer();
    const workbook = XLSX.read(wbBytes, {
      type: 'array',
      cellFormula: true,
      cellStyles: true,
      cellNF: true,
      raw: false,
    });

    let defaultModule = normalizeModuleCode(els.mergeModuleCode.value);
    if (!defaultModule) {
      defaultModule = normalizeModuleCode(getCellValue(getSheet(workbook, 'Exam'), EXAM_MODULE_CODE_CELL));
    }
    if (!defaultModule) {
      defaultModule = inferDefaultModuleCode(exportData.headers, exportData.rows, detectChildCourseColumn(exportData.headers));
      log(`[merge] inferred module code from Blackboard export: ${defaultModule}`);
    }
    if (!defaultModule) {
      throw new Error('Could not infer module code. Enter it manually.');
    }
    if (!els.mergeModuleCode.value.trim()) {
      els.mergeModuleCode.value = defaultModule;
    }

    const studentIdCol = detectStudentIdColumn(exportData.headers, els.mergeStudentIdCol.value.trim());
    const childCourseCol = detectChildCourseColumn(exportData.headers);
    const marksBySidModule = extractMarksWithModulesFromWorkbook(workbook);

    const modulesForTargets = new Set([defaultModule]);
    if (childCourseCol) {
      for (const row of exportData.rows) {
        const cm = extractModuleCodeFromChildCourse(row[childCourseCol]);
        if (cm) modulesForTargets.add(cm);
      }
    }
    for (const key of marksBySidModule.keys()) {
      const module = key.split('||')[1] || '';
      if (module) modulesForTargets.add(module);
    }

    const targetOverride = els.mergeTargetCol.value.trim();
    const targetByModule = new Map();
    for (const moduleCode of [...modulesForTargets].filter(Boolean)) {
      targetByModule.set(
        moduleCode,
        detectTargetColumnByModule(exportData.headers, moduleCode, targetOverride)
      );
    }

    let matched = 0;
    const blankUnmatched = els.mergeBlankUnmatched.checked;

    for (const row of exportData.rows) {
      const sid = normalizeRegNo(row[studentIdCol]);
      if (!sid) continue;

      let rowModule = defaultModule;
      if (childCourseCol) {
        const cm = extractModuleCodeFromChildCourse(row[childCourseCol]);
        if (cm) rowModule = cm;
      }

      let markModule = rowModule;
      let mark = marksBySidModule.get(`${sid}||${rowModule}`) || '';

      if (!mark) {
        const candidates = [];
        for (const [key, val] of marksBySidModule.entries()) {
          const [kSid, kModule] = key.split('||');
          if (kSid === sid) candidates.push([kModule, val]);
        }
        if (candidates.length === 1) {
          [markModule, mark] = candidates[0];
        }
      }

      if (!mark) {
        if (blankUnmatched) {
          const target = targetByModule.get(rowModule);
          if (target) row[target] = '';
        }
        continue;
      }

      const target = targetByModule.get(markModule) || targetByModule.get(rowModule);
      if (!target) {
        throw new Error(`Could not resolve target column for module ${markModule || rowModule}.`);
      }

      row[target] = mark;
      matched += 1;
    }

    const outputText = serializeDelimited(exportData.headers, exportData.rows, exportData.delimiter);
    const outputBytes = encodeOutputText(outputText, exportData.encoding);
    const outputName = uploadNameFromExport(exportData.filename);
    downloadBytes(outputBytes, outputName, 'text/csv;charset=utf-8');

    log(`[merge] wrote upload file: ${outputName} (CSV content)`);
    log(`[merge] rows in export: ${exportData.rows.length}`);
    log(`[merge] marks extracted from workbook: ${marksBySidModule.size}`);
    log(`[merge] rows updated: ${matched}`);
    log(`[merge] student id column: ${studentIdCol}`);
    for (const [mod, target] of targetByModule.entries()) {
      log(`[merge] target ${mod}: ${target}`);
    }
  }

  async function guarded(taskName, fn) {
    try {
      setBusy(true);
      log(`[${taskName}] started`);
      await fn();
      log(`[${taskName}] completed`);
      alert(`${taskName} completed successfully.`);
    } catch (err) {
      const message = err && err.message ? err.message : String(err);
      log(`[${taskName}] ERROR: ${message}`);
      alert(`${taskName} failed:\n${message}`);
    } finally {
      setBusy(false);
    }
  }

  els.buildBtn.addEventListener('click', () => guarded('build', runBuild));
  els.mergeBtn.addEventListener('click', () => guarded('merge', runMerge));
  els.clearLogBtn.addEventListener('click', () => {
    els.log.textContent = '';
  });

  log('Ready. All processing happens in your browser on this machine.');
})();
