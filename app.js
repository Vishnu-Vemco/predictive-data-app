const state = {
  fileName: '',
  sourceType: '',
  sheetName: '',
  headers: [],
  rows: [],
  resultRows: [],
  predictions: [],
  reviewItems: [],
  targetMappings: {},
  logs: [],
};

const elements = {
  fileInput: document.getElementById('fileInput'),
  dropZone: document.getElementById('dropZone'),
  fileName: document.getElementById('fileName'),
  dateColumn: document.getElementById('dateColumn'),
  targetColumns: document.getElementById('targetColumns'),
  historyColumns: document.getElementById('historyColumns'),
  zeroMissing: document.getElementById('zeroMissing'),
  roundingMode: document.getElementById('roundingMode'),
  runButton: document.getElementById('runButton'),
  downloadCsvButton: document.getElementById('downloadCsvButton'),
  downloadReportButton: document.getElementById('downloadReportButton'),
  clearLogButton: document.getElementById('clearLogButton'),
  logWindow: document.getElementById('logWindow'),
  previewTable: document.getElementById('previewTable'),
  previewNote: document.getElementById('previewNote'),
  rowCount: document.getElementById('rowCount'),
  missingCount: document.getElementById('missingCount'),
  predictedCount: document.getElementById('predictedCount'),
  metricRows: document.getElementById('metricRows'),
  metricMissing: document.getElementById('metricMissing'),
  metricPredicted: document.getElementById('metricPredicted'),
  metricReview: document.getElementById('metricReview'),
};

const HISTORY_KEYWORDS = [
  'previous',
  'prev',
  'last',
  'prior',
  'week',
  'month',
  'year',
  'historical',
  'history',
  'old',
  'lag',
  'same',
  'baseline',
  'average',
  'avg',
];

const TARGET_KEYWORDS = [
  'count',
  'people',
  'visitor',
  'traffic',
  'entry',
  'entries',
  'in',
  'out',
  'total',
  'current',
  'new',
];

const XLSX_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XLSX_OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XLSX_PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XLSX_DATE_NUM_FORMATS = new Set([14, 15, 16, 17, 22, 27, 30, 36, 45, 46, 47, 50, 57]);
const SUMMARY_ROW_LABELS = new Set(['total', 'average', 'minimum', 'maximum', 'min', 'max']);

function nowStamp() {
  return new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}

function addLog(message, level = 'info') {
  const prefix = level === 'warn' ? 'WARN' : level === 'error' ? 'ERROR' : 'INFO';
  state.logs.push(`[${nowStamp()}] ${prefix} ${message}`);
  elements.logWindow.textContent = state.logs.join('\n');
  elements.logWindow.scrollTop = elements.logWindow.scrollHeight;
}

function resetLog(message = 'Waiting for CSV upload...') {
  state.logs = [message];
  elements.logWindow.textContent = message;
}

function parseCSV(text) {
  const normalized = text.replace(/^\uFEFF/, '');
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;

  for (let index = 0; index < normalized.length; index += 1) {
    const char = normalized[index];
    const next = normalized[index + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        field += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === ',' && !inQuotes) {
      row.push(field);
      field = '';
      continue;
    }

    if ((char === '\n' || char === '\r') && !inQuotes) {
      if (char === '\r' && next === '\n') {
        index += 1;
      }
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      continue;
    }

    field += char;
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows.filter((line) => line.some((cell) => String(cell).trim() !== ''));
}

function findZipEndOfCentralDirectory(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const minimumOffset = Math.max(0, bytes.length - 66000);

  for (let offset = bytes.length - 22; offset >= minimumOffset; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) {
      return offset;
    }
  }

  throw new Error('Could not read the Excel workbook package.');
}

async function inflateZipEntry(compressedBytes) {
  if (typeof DecompressionStream === 'undefined') {
    throw new Error('This browser cannot decompress XLSX files. Please use a current Chrome or Edge browser, or export as CSV.');
  }

  const stream = new Blob([compressedBytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  const buffer = await new Response(stream).arrayBuffer();
  return new Uint8Array(buffer);
}

async function unzipXLSXEntries(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const view = new DataView(arrayBuffer);
  const decoder = new TextDecoder('utf-8');
  const entries = new Map();
  const endOffset = findZipEndOfCentralDirectory(bytes);
  const totalEntries = view.getUint16(endOffset + 10, true);
  const centralDirectoryOffset = view.getUint32(endOffset + 16, true);
  let offset = centralDirectoryOffset;

  for (let index = 0; index < totalEntries; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) {
      throw new Error('The Excel workbook package is not readable.');
    }

    const compressionMethod = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const fileName = decoder.decode(bytes.slice(offset + 46, offset + 46 + fileNameLength));

    if (view.getUint32(localHeaderOffset, true) !== 0x04034b50) {
      throw new Error(`The workbook entry "${fileName}" is not readable.`);
    }

    const localFileNameLength = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + localFileNameLength + localExtraLength;
    const compressedBytes = bytes.slice(dataStart, dataStart + compressedSize);

    if (compressionMethod === 0) {
      entries.set(fileName, decoder.decode(compressedBytes));
    } else if (compressionMethod === 8) {
      const decompressedBytes = await inflateZipEntry(compressedBytes);
      entries.set(fileName, decoder.decode(decompressedBytes));
    }

    offset += 46 + fileNameLength + extraLength + commentLength;
  }

  return entries;
}

function parseXml(xml, label) {
  const documentXml = new DOMParser().parseFromString(xml, 'application/xml');
  const parserError = documentXml.getElementsByTagName('parsererror')[0];
  if (parserError) {
    throw new Error(`Could not parse ${label}.`);
  }
  return documentXml;
}

function xmlElements(parent, namespace, localName) {
  const namespaced = Array.from(parent.getElementsByTagNameNS(namespace, localName));
  if (namespaced.length > 0) {
    return namespaced;
  }
  return Array.from(parent.getElementsByTagName(localName));
}

function xmlText(parent, namespace, localName) {
  const element = xmlElements(parent, namespace, localName)[0];
  return element ? element.textContent || '' : '';
}

function normalizeXlsxPath(basePath, target) {
  if (!target) {
    return '';
  }

  const combined = target.startsWith('/')
    ? target.replace(/^\/+/, '')
    : `${basePath.split('/').slice(0, -1).join('/')}/${target}`;

  const parts = [];
  combined.split('/').forEach((part) => {
    if (!part || part === '.') {
      return;
    }
    if (part === '..') {
      parts.pop();
      return;
    }
    parts.push(part);
  });

  return parts.join('/');
}

function parseRelationships(xml) {
  const relationships = new Map();
  const documentXml = parseXml(xml, 'workbook relationships');

  xmlElements(documentXml, XLSX_PACKAGE_REL_NS, 'Relationship').forEach((relationship) => {
    relationships.set(relationship.getAttribute('Id'), relationship.getAttribute('Target'));
  });

  return relationships;
}

function parseSharedStrings(xml) {
  if (!xml) {
    return [];
  }

  const documentXml = parseXml(xml, 'shared strings');
  return xmlElements(documentXml, XLSX_MAIN_NS, 'si').map((item) => {
    const textNodes = xmlElements(item, XLSX_MAIN_NS, 't');
    return textNodes.map((node) => node.textContent || '').join('');
  });
}

function parseStyles(xml) {
  const styles = {
    cellFormats: [],
    customNumberFormats: new Map(),
  };

  if (!xml) {
    return styles;
  }

  const documentXml = parseXml(xml, 'workbook styles');
  xmlElements(documentXml, XLSX_MAIN_NS, 'numFmt').forEach((format) => {
    styles.customNumberFormats.set(Number(format.getAttribute('numFmtId')), format.getAttribute('formatCode') || '');
  });

  const cellFormats = xmlElements(documentXml, XLSX_MAIN_NS, 'cellXfs')[0];
  if (cellFormats) {
    Array.from(cellFormats.children)
      .filter((child) => child.localName === 'xf')
      .forEach((format) => {
        styles.cellFormats.push(Number(format.getAttribute('numFmtId') || 0));
      });
  }

  return styles;
}

function isCustomDateFormat(formatCode) {
  const clean = formatCode
    .replace(/\[[^\]]+\]/g, '')
    .replace(/"[^"]*"/g, '')
    .toLowerCase();

  return /[dy]/.test(clean) && /m/.test(clean);
}

function isDateStyle(styleIndex, styles) {
  if (styleIndex === null || styleIndex === undefined || styleIndex === '') {
    return false;
  }

  const numberFormatId = styles.cellFormats[Number(styleIndex)];
  if (XLSX_DATE_NUM_FORMATS.has(numberFormatId)) {
    return true;
  }

  return isCustomDateFormat(styles.customNumberFormats.get(numberFormatId) || '');
}


function excelSerialToDateString(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return value;
  }

  const millisecondsPerDay = 24 * 60 * 60 * 1000;
  const date = new Date(Date.UTC(1899, 11, 30) + number * millisecondsPerDay);

  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  const hours = String(date.getUTCHours()).padStart(2, '0');
  const minutes = String(date.getUTCMinutes()).padStart(2, '0');

  return `${year}-${month}-${day} ${hours}:${minutes}`;
}


function columnIndexFromCellRef(cellRef) {
  const match = String(cellRef || '').match(/^([A-Z]+)/i);
  if (!match) {
    return 0;
  }

  return match[1]
    .toUpperCase()
    .split('')
    .reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0) - 1;
}

function parseXlsxCellValue(cell, sharedStrings, styles) {
  const cellType = cell.getAttribute('t');
  const styleIndex = cell.getAttribute('s');
  const rawValue = xmlText(cell, XLSX_MAIN_NS, 'v');

  if (cellType === 's' && rawValue !== '') {
    return sharedStrings[Number(rawValue)] ?? '';
  }

  if (cellType === 'inlineStr') {
    return xmlElements(cell, XLSX_MAIN_NS, 't').map((node) => node.textContent || '').join('');
  }

  if (isDateStyle(styleIndex, styles) && rawValue !== '') {
    return excelSerialToDateString(rawValue);
  }

  return rawValue;
}

function parseXlsxSheet(xml, sharedStrings, styles) {
  const documentXml = parseXml(xml, 'worksheet');
  const rows = [];

  xmlElements(documentXml, XLSX_MAIN_NS, 'row').forEach((rowElement) => {
    const row = [];
    xmlElements(rowElement, XLSX_MAIN_NS, 'c').forEach((cell) => {
      const cellRef = cell.getAttribute('r') || '';
      const columnIndex = columnIndexFromCellRef(cellRef);
      while (row.length <= columnIndex) {
        row.push('');
      }
      row[columnIndex] = parseXlsxCellValue(cell, sharedStrings, styles);
    });
    rows.push(row);
  });

  return rows.filter((line) => line.some((cell) => String(cell).trim() !== ''));
}

async function parseXLSX(arrayBuffer) {
  const entries = await unzipXLSXEntries(arrayBuffer);
  const workbookXml = entries.get('xl/workbook.xml');
  const workbookRelsXml = entries.get('xl/_rels/workbook.xml.rels');

  if (!workbookXml || !workbookRelsXml) {
    throw new Error('The XLSX file is missing workbook metadata.');
  }

  const workbook = parseXml(workbookXml, 'workbook');
  const workbookRelationships = parseRelationships(workbookRelsXml);
  const sheets = xmlElements(workbook, XLSX_MAIN_NS, 'sheet');
  const firstSheet = sheets[0];

  if (!firstSheet) {
    throw new Error('The XLSX file does not contain any sheets.');
  }

  const relationshipId = firstSheet.getAttributeNS(XLSX_OFFICE_REL_NS, 'id') || firstSheet.getAttribute('r:id');
  const sheetTarget = workbookRelationships.get(relationshipId);
  const sheetPath = normalizeXlsxPath('xl/workbook.xml', sheetTarget);
  const sheetXml = entries.get(sheetPath);

  if (!sheetXml) {
    throw new Error(`The worksheet "${firstSheet.getAttribute('name') || 'Sheet 1'}" could not be read.`);
  }

  return {
    rows: parseXlsxSheet(sheetXml, parseSharedStrings(entries.get('xl/sharedStrings.xml')), parseStyles(entries.get('xl/styles.xml'))),
    sheetName: firstSheet.getAttribute('name') || 'Sheet 1',
  };
}

async function readUploadedFile(file) {
  const name = file.name.toLowerCase();

  if (name.endsWith('.xlsx')) {
    const workbook = await parseXLSX(await file.arrayBuffer());
    return {
      parsedRows: workbook.rows,
      sourceType: 'XLSX',
      sheetName: workbook.sheetName,
    };
  }

  if (name.endsWith('.csv') || file.type === 'text/csv') {
    return {
      parsedRows: parseCSV(await file.text()),
      sourceType: 'CSV',
      sheetName: '',
    };
  }

  throw new Error('Please upload a CSV or XLSX export file.');
}

function makeUniqueHeaders(rawHeaders) {
  const seen = new Map();
  return rawHeaders.map((header, index) => {
    const fallback = `Column ${index + 1}`;
    const clean = String(header || fallback).trim() || fallback;
    const count = seen.get(clean) || 0;
    seen.set(clean, count + 1);
    return count === 0 ? clean : `${clean} (${count + 1})`;
  });
}


function isIndexHeader(header) {
  const lower = String(header || '').trim().toLowerCase();
  return lower === '' || lower === 'index' || lower === 'idx' || lower === '#' || lower.startsWith('unnamed') || /^index( \(\d+\))?$/.test(lower);
}

function shouldRemoveSourceColumn(header, index) {
  return isIndexHeader(header);
}

function formatDate(day, month, year) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${String(day).padStart(2, '0')} ${months[Number(month) - 1] || month} ${year}`;
}

function splitDateTime(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, '0');
    const day = String(value.getDate()).padStart(2, '0');
    const hours = String(value.getHours()).padStart(2, '0');
    const minutes = String(value.getMinutes()).padStart(2, '0');
    return { date: formatDate(day, month, year), time: `${hours}:${minutes}` };
  }

  const text = String(value || '').trim();

  let match = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::\d{2})?)?/);
  if (match) {
    return {
      date: formatDate(match[3], match[2], match[1]),
      time: `${match[4] || '00'}:${match[5] || '00'}`,
    };
  }

  match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ T](\d{2}):(\d{2})(?::\d{2})?)?/);
  if (match) {
    return {
      date: formatDate(match[1], match[2], match[3]),
      time: `${match[4] || '00'}:${match[5] || '00'}`,
    };
  }

  return { date: text, time: '00:00' };
}

function rowsToObjects(parsedRows) {
  if (parsedRows.length === 0) {
    return { headers: [], rows: [] };
  }

  const rawHeaders = makeUniqueHeaders(parsedRows[0]);
  const keptColumns = rawHeaders
    .map((header, index) => ({ header, index }))
    .filter((column) => !shouldRemoveSourceColumn(column.header, column.index));

  let headers = keptColumns.map((column) => column.header);
  const firstKeptColumnIndex = keptColumns[0]?.index ?? 0;

  const hasDateTime = parsedRows.slice(1, 30).some((row) => {
    const value = String(row[firstKeptColumnIndex] ?? '').trim();
    return /^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2})?/.test(value) || /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}/.test(value);
  });

  if (hasDateTime) {
    headers[0] = 'Date';
    if (!headers.includes('Time')) {
      headers.splice(1, 0, 'Time');
    }
  }

  const rows = parsedRows.slice(1).map((cells, rowIndex) => {
    const row = { __rowNumber: rowIndex + 2 };

    headers.forEach((header, outputIndex) => {
      if (hasDateTime && outputIndex === 0) {
        const { date, time } = splitDateTime(cells[firstKeptColumnIndex]);
        row.Date = date;
        row.Time = time;
        return;
      }

      if (hasDateTime && outputIndex === 1 && header === 'Time') {
        return;
      }

      const keptIndex = hasDateTime && outputIndex > 1 ? outputIndex - 1 : outputIndex;
      const sourceColumnIndex = keptColumns[keptIndex]?.index;
      row[header] = sourceColumnIndex !== undefined ? cells[sourceColumnIndex] ?? '' : '';
    });

    return row;
  });

  const cleanedRows = rows.filter((row) => {
    return !Object.values(row).some((val) =>
      SUMMARY_ROW_LABELS.has(String(val).trim().toLowerCase())
    );
  });

  return { headers, rows: cleanedRows };
}


function parseNumber(value) {
  if (value === null || value === undefined) {
    return null;
  }

  const raw = String(value).trim();
  if (raw === '') {
    return null;
  }

  const cleaned = raw.replace(/,/g, '');
  const number = Number(cleaned);
  return Number.isFinite(number) ? number : null;
}

function isZeroValue(value) {
  const number = parseNumber(value);
  return number !== null && number === 0;
}

function numericProfile(header) {
  const sample = state.rows.slice(0, 200);
  if (sample.length === 0) {
    return { count: 0, ratio: 0 };
  }

  const count = sample.reduce((total, row) => (parseNumber(row[header]) !== null ? total + 1 : total), 0);
  return { count, ratio: count / sample.length };
}

function looksLikeDateHeader(header) {
  const lower = header.toLowerCase();
  return lower.includes('date') || lower.includes('time') || lower.includes('timestamp') || lower.includes('period');
}

function looksLikeHistoryHeader(header) {
  const lower = header.toLowerCase();
  return HISTORY_KEYWORDS.some((keyword) => lower.includes(keyword));
}

function scoreTargetHeader(header) {
  const lower = header.toLowerCase();
  let score = 0;

  TARGET_KEYWORDS.forEach((keyword) => {
    if (lower === keyword) {
      score += 8;
    } else if (lower.includes(keyword)) {
      score += 4;
    }
  });

  if (looksLikeHistoryHeader(header)) {
    score -= 6;
  }

  if (lower.includes('predict') || lower.includes('forecast') || lower.includes('estimate')) {
    score -= 6;
  }

  return score;
}


function cleanHeadingForMatch(header) {
  return String(header || '')
    .toLowerCase()
    .replace(/[▲▼]/g, ' ')
    .replace(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}\s*-\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}/gi, ' ')
    .replace(/\d{4}-\d{2}-\d{2}\s*-\s*\d{4}-\d{2}-\d{2}/g, ' ')
    .replace(/\bsame\s+weekday\b/g, ' ')
    .replace(/\bsame\s+date\b/g, ' ')
    .replace(/\bprevious\s+week\b/g, ' ')
    .replace(/\bprevious\s+month\b/g, ' ')
    .replace(/\bprevious\b/g, ' ')
    .replace(/\bweek\b/g, ' ')
    .replace(/\bmonth\b/g, ' ')
    .replace(/\byear\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function headingTokens(header) {
  const stop = new Set(['same', 'date', 'weekday', 'previous', 'week', 'month', 'year', 'index', 'total']);
  return cleanHeadingForMatch(header)
    .split(/\s+/)
    .filter((token) => token && token.length > 1 && !stop.has(token));
}

function tokenSimilarity(a, b) {
  const aTokens = new Set(headingTokens(a));
  const bTokens = new Set(headingTokens(b));

  if (aTokens.size === 0 || bTokens.size === 0) {
    return 0;
  }

  let intersection = 0;
  aTokens.forEach((token) => {
    if (bTokens.has(token)) {
      intersection += 1;
    }
  });

  return intersection / Math.max(aTokens.size, bTokens.size);
}

function sideToken(header) {
  const clean = cleanHeadingForMatch(header);
  if (clean.includes('left')) return 'left';
  if (clean.includes('right')) return 'right';
  return '';
}

function isTotalHeader(header) {
  return /\btotal\b/i.test(String(header || ''));
}

function isPredictionTargetHeader(header) {
  if (looksLikeDateHeader(header) || looksLikeHistoryHeader(header) || isTotalHeader(header) || isIndexHeader(header)) {
    return false;
  }

  const profile = numericProfile(header);
  return profile.ratio >= 0.35 && scoreTargetHeader(header) > 0;
}

function guessColumns() {
  const numericHeaders = state.headers
    .map((header) => ({ header, profile: numericProfile(header) }))
    .filter((item) => item.profile.ratio >= 0.35);

  const dateColumn = state.headers.find(looksLikeDateHeader) || '';

  const targetColumns = numericHeaders
    .map((item) => ({ ...item, score: scoreTargetHeader(item.header) }))
    .filter((item) => isPredictionTargetHeader(item.header))
    .sort((a, b) => b.score - a.score || b.profile.ratio - a.profile.ratio)
    .map((item) => item.header);

  const fallbackTargets = targetColumns.length > 0
    ? targetColumns
    : numericHeaders
        .filter((item) => !looksLikeDateHeader(item.header) && !looksLikeHistoryHeader(item.header) && !isTotalHeader(item.header))
        .slice(0, 2)
        .map((item) => item.header);

  const historyColumns = numericHeaders
    .filter((item) => !fallbackTargets.includes(item.header) && looksLikeHistoryHeader(item.header) && !isTotalHeader(item.header))
    .map((item) => item.header);

  return {
    dateColumn,
    targetColumns: fallbackTargets,
    historyColumns,
    numericHeaders: numericHeaders.map((item) => item.header),
  };
}

function populateSelect(select, headers, selectedValue, allowBlank = true) {
  select.innerHTML = '';

  if (allowBlank) {
    const blank = document.createElement('option');
    blank.value = '';
    blank.textContent = 'Not available';
    select.appendChild(blank);
  }

  headers.forEach((header) => {
    const option = document.createElement('option');
    option.value = header;
    option.textContent = header;
    select.appendChild(option);
  });

  select.value = selectedValue || '';
}

function populateTargetColumns(headers, selectedHeaders) {
  elements.targetColumns.innerHTML = '';

  if (headers.length === 0) {
    const empty = document.createElement('span');
    empty.className = 'muted-line';
    empty.textContent = 'No numeric target columns detected.';
    elements.targetColumns.appendChild(empty);
    return;
  }

  headers.forEach((header) => {
    const label = document.createElement('label');
    const input = document.createElement('input');
    const text = document.createElement('span');

    input.type = 'checkbox';
    input.value = header;
    input.checked = selectedHeaders.includes(header);
    text.textContent = header;
    text.title = header;

    label.append(input, text);
    elements.targetColumns.appendChild(label);
  });
}

function getSelectedTargetColumns() {
  return Array.from(elements.targetColumns.querySelectorAll('input[type="checkbox"]:checked')).map((input) => input.value);
}

function populateHistoryColumns(headers, selectedHeaders) {
  elements.historyColumns.innerHTML = '';

  if (headers.length === 0) {
    const empty = document.createElement('span');
    empty.className = 'muted-line';
    empty.textContent = 'History columns will be mapped automatically from matching headings.';
    elements.historyColumns.appendChild(empty);
    return;
  }

  headers.forEach((header) => {
    const label = document.createElement('label');
    const input = document.createElement('input');
    const text = document.createElement('span');

    input.type = 'checkbox';
    input.value = header;
    input.checked = selectedHeaders.includes(header);
    text.textContent = header;
    text.title = header;

    label.append(input, text);
    elements.historyColumns.appendChild(label);
  });
}

function getSelectedHistoryColumns() {
  return Array.from(elements.historyColumns.querySelectorAll('input[type="checkbox"]:checked')).map((input) => input.value);
}

function scoreHistoryForTarget(targetHeader, historyHeader) {
  if (!looksLikeHistoryHeader(historyHeader) || isTotalHeader(historyHeader) || isIndexHeader(historyHeader)) {
    return 0;
  }

  let score = tokenSimilarity(targetHeader, historyHeader);

  const targetSide = sideToken(targetHeader);
  const historySide = sideToken(historyHeader);
  if (targetSide && historySide && targetSide === historySide) {
    score += 0.3;
  }
  if (targetSide && historySide && targetSide !== historySide) {
    score -= 0.4;
  }

  const targetClean = cleanHeadingForMatch(targetHeader);
  const historyClean = cleanHeadingForMatch(historyHeader);
  if (targetClean && historyClean.includes(targetClean)) {
    score += 0.3;
  }

  if (/previous\s+week|same\s+weekday/i.test(historyHeader)) {
    score += 0.1;
  }
  if (/previous\s+month|same\s+date/i.test(historyHeader)) {
    score += 0.1;
  }

  return score;
}

function buildTargetHistoryMappings(targetColumns, candidateHistoryColumns) {
  const mappings = {};

  targetColumns.forEach((targetColumn) => {
    const ranked = candidateHistoryColumns
      .filter((header) => header !== targetColumn)
      .map((historyColumn) => ({
        historyColumn,
        score: scoreHistoryForTarget(targetColumn, historyColumn),
      }))
      .filter((item) => item.score >= 0.45)
      .sort((a, b) => b.score - a.score);

    mappings[targetColumn] = ranked.slice(0, 4).map((item) => item.historyColumn);
  });

  return mappings;
}

function selectedTargetHistoryColumns(targetColumn) {
  return state.targetMappings[targetColumn] || [];
}

function chooseAutoMethod() {
  const targetColumns = getSelectedTargetColumns();

  if (targetColumns.length === 0) {
    return 'weighted';
  }

  const missingRows = state.rows.filter((row) =>
    targetColumns.some((targetColumn) => shouldPredict(row, targetColumn))
  );

  if (missingRows.length === 0) {
    return 'weighted';
  }

  const totalCells = Math.max(1, state.rows.length * targetColumns.length);
  const missingCells = missingRows.reduce((count, row) => {
    return count + targetColumns.filter((targetColumn) => shouldPredict(row, targetColumn)).length;
  }, 0);
  const missingRatio = missingCells / totalCells;

  const historyValues = targetColumns.flatMap((targetColumn) => {
    const historyColumns = selectedTargetHistoryColumns(targetColumn);
    return missingRows.flatMap((row) =>
      historyColumns
        .map((column) => parseNumber(row[column]))
        .filter((value) => value !== null && value > 0)
    );
  });

  if (historyValues.length === 0) {
    return 'median';
  }

  const avg = average(historyValues);
  const max = Math.max(...historyValues);

  if (avg > 0 && max / avg > 2) {
    return 'median';
  }

  if (missingRatio > 0.2) {
    return 'average';
  }

  return 'weighted';
}

function methodName() {
  return document.querySelector('input[name="method"]:checked')?.value || 'weighted';
}

function roundPrediction(value) {
  const mode = elements.roundingMode.value;
  if (mode === 'integer') {
    return Math.round(value);
  }
  if (mode === 'oneDecimal') {
    return Math.round(value * 10) / 10;
  }
  return Math.round(value * 1000) / 1000;
}

function median(values) {
  if (values.length === 0) {
    return null;
  }

  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
}

function average(values) {
  if (values.length === 0) {
    return null;
  }

  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function trimmedAverage(values) {
  if (values.length <= 2) {
    return average(values);
  }

  const sorted = [...values].sort((a, b) => a - b);
  const trimCount = Math.floor(sorted.length * 0.15);
  const trimmed = sorted.slice(trimCount, sorted.length - trimCount);
  return average(trimmed.length > 0 ? trimmed : sorted);
}

function weightForHeader(header) {
  const normalized = header.toLowerCase().replace(/[^a-z0-9]+/g, ' ');
  if (/\bweek\b|\bweekly\b/.test(normalized)) {
    return 0.45;
  }
  if (/\bmonth\b|\bmonthly\b/.test(normalized)) {
    return 0.35;
  }
  if (/\byear\b|\byearly\b/.test(normalized)) {
    return 0.2;
  }
  if (/\baverage\b|\bavg\b|\bbaseline\b/.test(normalized)) {
    return 0.3;
  }
  return 0.25;
}

function predictFromHistory(row, historyColumns, method) {
  const values = historyColumns
    .map((column) => ({ column, value: parseNumber(row[column]) }))
    .filter((item) => item.value !== null && item.value > 0);

  if (values.length === 0) {
    return { value: null, sourceValues: [] };
  }

  if (method === 'median') {
    return {
      value: median(values.map((item) => item.value)),
      sourceValues: values,
    };
  }

  if (method === 'average') {
    return {
      value: average(values.map((item) => item.value)),
      sourceValues: values,
    };
  }

  const totalWeight = values.reduce((sum, item) => sum + weightForHeader(item.column), 0);
  const weighted = values.reduce((sum, item) => sum + item.value * weightForHeader(item.column), 0) / totalWeight;
  const robust = trimmedAverage(values.map((item) => item.value));

  return {
    value: values.length >= 3 ? weighted * 0.7 + robust * 0.3 : weighted,
    sourceValues: values,
  };
}

function isSummaryRow(row) {
  const dateColumn = elements.dateColumn.value;
  if (!dateColumn) {
    return false;
  }

  const label = String(row[dateColumn] ?? '').trim().toLowerCase();
  return SUMMARY_ROW_LABELS.has(label);
}

function resetPredictionState() {
  state.resultRows = state.rows.map((row) => ({ ...row }));
  state.predictions = [];
  state.reviewItems = [];
  elements.downloadCsvButton.disabled = true;
  elements.downloadReportButton.disabled = true;
}

function shouldPredict(row, targetColumn) {
  if (!targetColumn) {
    return false;
  }

  if (isSummaryRow(row)) {
    return false;
  }

  if (elements.zeroMissing.checked) {
    return isZeroValue(row[targetColumn]);
  }

  return parseNumber(row[targetColumn]) === null;
}


function runPrediction() {
  const dateColumn = elements.dateColumn.value;
  const targetColumns = getSelectedTargetColumns();
  const method = methodName();

  if (targetColumns.length === 0) {
    addLog('Select at least one count column before running prediction.', 'error');
    return;
  }

  elements.runButton.disabled = true;

  state.predictions = [];
  state.reviewItems = [];
  state.resultRows = state.rows.map((row) => ({ ...row }));

  targetColumns.forEach((targetColumn) => {
    if (!state.targetMappings[targetColumn] || state.targetMappings[targetColumn].length === 0) {
      state.targetMappings[targetColumn] =
        buildTargetHistoryMappings([targetColumn], getSelectedHistoryColumns())[targetColumn] || [];
    }
  });

  addLog(`Running ${method} prediction for ${targetColumns.length} target column(s).`);

  targetColumns.forEach((targetColumn) => {
    const historyColumns = selectedTargetHistoryColumns(targetColumn);

    if (historyColumns.length === 0) {
      addLog(`No matching history columns found for "${targetColumn}".`, 'warn');
      return;
    }

    addLog(`"${targetColumn}" mapped to ${historyColumns.length} history column(s).`);

    state.resultRows.forEach((row) => {
      if (!shouldPredict(row, targetColumn)) {
        return;
      }

      const prediction = predictFromHistory(row, historyColumns, method);
      const dateLabel = dateColumn
        ? row[dateColumn] || `row ${row.__rowNumber}`
        : `row ${row.__rowNumber}`;

      if (prediction.value === null) {
        state.reviewItems.push({
          rowNumber: row.__rowNumber,
          date: dateLabel,
          targetColumn,
          reason: `No positive previous data values found for ${targetColumn}`,
        });

        row.__reviewCells = row.__reviewCells || {};
        row.__reviewCells[targetColumn] = true;
        row.__predictionStatus = 'review';

        return;
      }

      const rounded = roundPrediction(prediction.value);
      const originalValue = row[targetColumn];

      row[targetColumn] = String(rounded);
      row.__predictedCells = row.__predictedCells || {};
      row.__predictedCells[targetColumn] = true;
      row.__predictionStatus = 'predicted';

      const sources = prediction.sourceValues
        .map((item) => `${item.column}=${item.value}`)
        .join(', ');

      state.predictions.push({
        rowNumber: row.__rowNumber,
        date: dateLabel,
        targetColumn,
        originalValue,
        predictedValue: rounded,
        method,
        sources,
      });
    });
  });

  if (state.predictions.length === 0 && state.reviewItems.length === 0) {
    addLog('No missing values were found with the current settings.');
  } else {
    addLog(
      `Prediction complete. ${state.predictions.length} value(s) predicted, ${state.reviewItems.length} need review.`
    );
  }

  updateMetrics();
  renderPreview();

  elements.downloadCsvButton.disabled = state.predictions.length === 0;
  elements.downloadReportButton.disabled =
    state.predictions.length === 0 && state.reviewItems.length === 0;

  elements.runButton.disabled = false;
}

function updateMetrics() {
  const targetColumns = getSelectedTargetColumns();
  const scannableRows = state.rows.filter((row) => !isSummaryRow(row)).length;
  const missing = targetColumns.reduce((count, targetColumn) => {
    return count + state.rows.filter((row) => shouldPredict(row, targetColumn)).length;
  }, 0);

  elements.rowCount.textContent = state.rows.length ? `${scannableRows} data rows` : 'No file loaded';
  elements.missingCount.textContent = `${missing} missing`;
  elements.predictedCount.textContent = `${state.predictions.length} predicted`;
  elements.metricRows.textContent = String(scannableRows);
  elements.metricMissing.textContent = String(missing);
  elements.metricPredicted.textContent = String(state.predictions.length);
  elements.metricReview.textContent = String(state.reviewItems.length);
}


function visiblePreviewRows() {
  if (state.resultRows.length === 0) {
    return [];
  }

  const highlighted = state.resultRows.filter((row) => row.__predictionStatus);
  return highlighted.length > 0 ? highlighted.slice(0, 80) : state.resultRows.slice(0, 80);
}

function renderPreview() {
  const thead = elements.previewTable.querySelector('thead');
  const tbody = elements.previewTable.querySelector('tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  if (state.rows.length === 0) {
    tbody.innerHTML = '<tr><td class="empty-state">No data loaded yet.</td></tr>';
    elements.previewNote.textContent = 'Upload a CSV or XLSX to begin.';
    return;
  }

  const rows = visiblePreviewRows();
  const selectedTargets = getSelectedTargetColumns();
  const importantHeaders = [
    ...state.headers.filter((header) => looksLikeDateHeader(header) || header === 'Date' || header === 'Time'),
    ...selectedTargets,
  ];
  const displayHeaders = [...new Set(importantHeaders.length > 0 ? importantHeaders : state.headers.slice(0, 12))].slice(0, 14);

  const headerRow = document.createElement('tr');
  displayHeaders.forEach((header) => {
    const th = document.createElement('th');
    th.textContent = header;
    th.title = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    if (row.__predictionStatus === 'predicted') {
      tr.classList.add('predicted-row');
    }
    if (row.__predictionStatus === 'review') {
      tr.classList.add('review-row');
    }

    displayHeaders.forEach((header) => {
      const td = document.createElement('td');
      td.textContent = row[header] ?? '';
      td.title = row[header] ?? '';

      if (row.__predictedCells?.[header]) {
        td.classList.add('predicted-cell');
      }
      if (row.__reviewCells?.[header]) {
        td.classList.add('review-cell');
      }

      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  const highlightedCount = state.predictions.length + state.reviewItems.length;
  elements.previewNote.textContent = highlightedCount
    ? `Showing ${rows.length} affected row(s).`
    : `Showing first ${rows.length} row(s).`;
}

function escapeCSVCell(value) {
  const text = value === null || value === undefined ? '' : String(value);
  if (/[",\r\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function buildCSV(rows) {
  const lines = [];
  lines.push(state.headers.map(escapeCSVCell).join(','));
  rows.forEach((row) => {
    lines.push(state.headers.map((header) => escapeCSVCell(row[header])).join(','));
  });
  return lines.join('\r\n');
}


function buildReportCSV() {
  const headers = ['row_number', 'date', 'target_column', 'status', 'original_value', 'predicted_value', 'method', 'sources_or_reason'];
  const lines = [headers.join(',')];

  state.predictions.forEach((item) => {
    lines.push([
      item.rowNumber,
      item.date,
      item.targetColumn,
      'predicted',
      item.originalValue,
      item.predictedValue,
      item.method,
      item.sources,
    ].map(escapeCSVCell).join(','));
  });

  state.reviewItems.forEach((item) => {
    lines.push([
      item.rowNumber,
      item.date,
      item.targetColumn,
      'needs_review',
      '',
      '',
      '',
      item.reason,
    ].map(escapeCSVCell).join(','));
  });

  return lines.join('\r\n');
}

function downloadText(filename, text) {
  const blob = new Blob([text], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}


async function downloadTrueXlsx(filename, rows) {
  if (typeof ExcelJS === 'undefined') {
    throw new Error('ExcelJS is not loaded. Check the ExcelJS script in index.html.');
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Predicted Data');

  worksheet.addRow(state.headers);

  worksheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFEAEAEA' },
    };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
  });

  rows.forEach((row) => {
    const excelRow = worksheet.addRow(
      state.headers.map((header) => row[header] ?? '')
    );

    state.headers.forEach((header, index) => {
      const cell = excelRow.getCell(index + 1);

      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };

      if (row.__predictedCells?.[header]) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFC000' },
        };

        cell.font = {
          bold: true,
          color: { argb: 'FF000000' },
        };
      }

      if (row.__reviewCells?.[header]) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFE0B2' },
        };
      }
    });
  });

  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  worksheet.columns.forEach((column) => {
    let maxLength = 12;

    column.eachCell({ includeEmpty: true }, (cell) => {
      const value = cell.value ? String(cell.value) : '';
      maxLength = Math.max(maxLength, value.length + 2);
    });

    column.width = Math.min(maxLength, 40);
  });

  const buffer = await workbook.xlsx.writeBuffer();

  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');

  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();

  URL.revokeObjectURL(url);
}

function baseFileName() {
  return (state.fileName || 'predictive-data').replace(/\.(csv|xlsx)$/i, '');
}


async function handleFile(file) {
  if (!file) {
    return;
  }

  try {
    resetLog(`Loading "${file.name}"...`);
    const uploaded = await readUploadedFile(file);
    const { headers, rows } = rowsToObjects(uploaded.parsedRows);

    if (headers.length === 0 || rows.length === 0) {
      throw new Error('The export file does not contain headers and data rows.');
    }

    state.fileName = file.name;
    state.sourceType = uploaded.sourceType;
    state.sheetName = uploaded.sheetName;
    state.headers = headers;
    state.rows = rows;
    resetPredictionState();
    state.logs = [];

    const guesses = guessColumns();
    const historyCandidates = guesses.historyColumns.length > 0
      ? guesses.historyColumns
      : guesses.numericHeaders.filter((header) => looksLikeHistoryHeader(header) && !isTotalHeader(header));

    populateSelect(elements.dateColumn, headers, guesses.dateColumn);
    populateTargetColumns(guesses.numericHeaders.filter((header) => !looksLikeDateHeader(header) && !looksLikeHistoryHeader(header) && !isTotalHeader(header)), guesses.targetColumns);
    populateHistoryColumns(historyCandidates, historyCandidates);

    state.targetMappings = buildTargetHistoryMappings(guesses.targetColumns, historyCandidates);

    const recommendedMethod = chooseAutoMethod();
    const methodInput = document.querySelector(`input[name="method"][value="${recommendedMethod}"]`);
    if (methodInput) {
      methodInput.checked = true;
    }

    elements.fileName.textContent = file.name;
    elements.runButton.disabled = false;
    elements.downloadCsvButton.disabled = true;
    elements.downloadReportButton.disabled = true;

    addLog(`Loaded "${file.name}" as ${uploaded.sourceType} with ${rows.length} row(s) and ${headers.length} column(s).`);
    if (uploaded.sheetName) {
      addLog(`Using workbook sheet: "${uploaded.sheetName}".`);
    }

    const selectedTargets = getSelectedTargetColumns();
    if (selectedTargets.length > 0) {
      addLog(`Detected target columns: ${selectedTargets.join(' | ')}.`);
      selectedTargets.forEach((targetColumn) => {
        const mapped = selectedTargetHistoryColumns(targetColumn);
        addLog(`Mapped "${targetColumn}" to ${mapped.length} history column(s): ${mapped.join(' | ') || 'none found'}.`);
      });
    } else {
      addLog('No target columns were auto-selected. Please choose them manually.', 'warn');
    }

    addLog(`Recommended prediction method selected: ${recommendedMethod}. You can change it before running prediction.`);

    updateMetrics();
    renderPreview();
  } catch (error) {
    resetLog(`ERROR ${error.message}`);
    elements.runButton.disabled = true;
    elements.downloadCsvButton.disabled = true;
    elements.downloadReportButton.disabled = true;
  }
}


function syncTargetMappingsAfterTargetChange() {
  const targetColumns = getSelectedTargetColumns();
  const historyCandidates = getSelectedHistoryColumns();

  state.targetMappings = buildTargetHistoryMappings(targetColumns, historyCandidates);

  const recommendedMethod = chooseAutoMethod();
  const methodInput = document.querySelector(`input[name="method"][value="${recommendedMethod}"]`);
  if (methodInput) {
    methodInput.checked = true;
  }

  resetPredictionState();
  updateMetrics();
  renderPreview();
}

function syncHistoryAfterTargetChange() {
  syncTargetMappingsAfterTargetChange();
}

function syncAfterDateColumnChange() {
  resetPredictionState();
  updateMetrics();
  renderPreview();
}

elements.fileInput.addEventListener('change', (event) => {
  handleFile(event.target.files[0]);
});

elements.dropZone.addEventListener('dragover', (event) => {
  event.preventDefault();
  elements.dropZone.classList.add('dragging');
});

elements.dropZone.addEventListener('dragleave', () => {
  elements.dropZone.classList.remove('dragging');
});

elements.dropZone.addEventListener('drop', (event) => {
  event.preventDefault();
  elements.dropZone.classList.remove('dragging');
  handleFile(event.dataTransfer.files[0]);
});

elements.dateColumn.addEventListener('change', syncAfterDateColumnChange);
elements.targetColumns.addEventListener('change', syncTargetMappingsAfterTargetChange);
elements.zeroMissing.addEventListener('change', () => {
  syncTargetMappingsAfterTargetChange();
});
elements.runButton.addEventListener('click', runPrediction);
elements.clearLogButton.addEventListener('click', () => resetLog('Log cleared.'));

elements.downloadCsvButton.addEventListener('click', async () => {
  try {
    addLog('Preparing Excel download...');
    await downloadTrueXlsx(`${baseFileName()}-predicted.xlsx`, state.resultRows);
    addLog('Excel file downloaded successfully.');
  } catch (error) {
    addLog(`Excel download failed: ${error.message}`, 'error');
    console.error(error);
  }
});

elements.downloadReportButton.addEventListener('click', () => {
  downloadText(`${baseFileName()}-prediction-report.csv`, buildReportCSV());
});


