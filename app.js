const state = {
  fileName: '',
  sourceType: '',
  sheetName: '',
  headers: [],
  rows: [],
  resultRows: [],
  predictions: [],
  reviewItems: [],
  logs: [],
};

const elements = {
  fileInput: document.getElementById('fileInput'),
  dropZone: document.getElementById('dropZone'),
  fileName: document.getElementById('fileName'),
  dateColumn: document.getElementById('dateColumn'),
  targetColumn: document.getElementById('targetColumn'),
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

function rowsToObjects(parsedRows) {
  if (parsedRows.length === 0) {
    return { headers: [], rows: [] };
  }

  const headers = makeUniqueHeaders(parsedRows[0]);
  const rows = parsedRows.slice(1).map((cells, rowIndex) => {
    const row = { __rowNumber: rowIndex + 2 };
    headers.forEach((header, columnIndex) => {
      row[header] = cells[columnIndex] ?? '';
    });
    return row;
  });

  return { headers, rows };
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

function guessColumns() {
  const numericHeaders = state.headers
    .map((header) => ({ header, profile: numericProfile(header) }))
    .filter((item) => item.profile.ratio >= 0.35);

  const dateColumn = state.headers.find(looksLikeDateHeader) || '';

  const targetColumn = numericHeaders
    .map((item) => ({ ...item, score: scoreTargetHeader(item.header) }))
    .sort((a, b) => b.score - a.score || b.profile.ratio - a.profile.ratio)[0]?.header || numericHeaders[0]?.header || '';

  const historyColumns = numericHeaders
    .filter((item) => item.header !== targetColumn)
    .map((item) => ({
      header: item.header,
      score: looksLikeHistoryHeader(item.header) ? 10 : 0,
      ratio: item.profile.ratio,
    }))
    .sort((a, b) => b.score - a.score || b.ratio - a.ratio)
    .filter((item, index) => item.score > 0 || index < 3)
    .map((item) => item.header);

  return { dateColumn, targetColumn, historyColumns, numericHeaders: numericHeaders.map((item) => item.header) };
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

function populateHistoryColumns(headers, selectedHeaders) {
  elements.historyColumns.innerHTML = '';

  if (headers.length === 0) {
    const empty = document.createElement('span');
    empty.className = 'muted-line';
    empty.textContent = 'No numeric history columns detected.';
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

function chooseAutoMethod() {
  const targetColumn = elements.targetColumn.value;
  const historyColumns = getSelectedHistoryColumns();

  if (!targetColumn || historyColumns.length === 0 || state.rows.length === 0) {
    return 'weighted';
  }

  const missingRows = state.rows.filter((row) => shouldPredict(row, targetColumn));
  if (missingRows.length === 0) {
    return 'weighted';
  }

  const missingRatio = missingRows.length / state.rows.length;
  const historyValues = missingRows.flatMap((row) =>
    historyColumns
      .map((column) => parseNumber(row[column]))
      .filter((value) => value !== null && value > 0)
  );

  if (historyValues.length === 0) {
    return 'median';
  }

  const avg = average(historyValues);
  const max = Math.max(...historyValues);
  const hasOutlier = avg > 0 && max / avg > 2;

  if (hasOutlier) {
    return 'median';
  }

  if (missingRatio > 0.2) {
    return 'average';
  }

  if (historyColumns.length >= 3) {
    return 'weighted';
  }

  return 'average';
}

function selectedPredictionMethod() {
  const selected = methodName();
  return selected === 'auto' ? chooseAutoMethod() : selected;
}

function methodName() {
  return document.querySelector('input[name="method"]:checked')?.value || 'auto';
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
  const targetColumn = elements.targetColumn.value;
  const historyColumns = getSelectedHistoryColumns();
  const method = selectedPredictionMethod();

  if (!targetColumn) {
    addLog('Select the count column before running prediction.', 'error');
    return;
  }

  if (historyColumns.length === 0) {
    addLog('Select at least one previous data column.', 'error');
    return;
  }

  state.predictions = [];
  state.reviewItems = [];
  state.resultRows = state.rows.map((row) => ({ ...row }));

  addLog(`Running ${method} prediction on "${targetColumn}" using ${historyColumns.length} history column(s).`);

  state.resultRows.forEach((row) => {
    if (!shouldPredict(row, targetColumn)) {
      return;
    }

    const prediction = predictFromHistory(row, historyColumns, method);
    const dateLabel = dateColumn ? row[dateColumn] || `row ${row.__rowNumber}` : `row ${row.__rowNumber}`;

    if (prediction.value === null) {
      state.reviewItems.push({
        rowNumber: row.__rowNumber,
        date: dateLabel,
        reason: 'No positive previous data values found',
      });
      row.__predictionStatus = 'review';
      addLog(`Missing count found at ${dateLabel}, but no usable previous data was available.`, 'warn');
      return;
    }

    const rounded = roundPrediction(prediction.value);
    row.__originalValue = row[targetColumn];
    row[targetColumn] = String(rounded);
    row.__predictedValue = rounded;
    row.__predictionStatus = 'predicted';

    const sources = prediction.sourceValues
      .map((item) => `${item.column}=${item.value}`)
      .join(', ');

    state.predictions.push({
      rowNumber: row.__rowNumber,
      date: dateLabel,
      originalValue: row.__originalValue,
      predictedValue: rounded,
      method,
      sources,
    });

    addLog(`Predicted ${targetColumn} at ${dateLabel}: ${row.__originalValue} -> ${rounded} (${sources}).`);
  });

  if (state.predictions.length === 0 && state.reviewItems.length === 0) {
    addLog('No missing values were found with the current settings.');
  }

  updateMetrics();
  renderPreview();
  elements.downloadCsvButton.disabled = state.predictions.length === 0;
  elements.downloadReportButton.disabled = state.predictions.length === 0 && state.reviewItems.length === 0;
}

function updateMetrics() {
  const targetColumn = elements.targetColumn.value;
  const scannableRows = state.rows.filter((row) => !isSummaryRow(row)).length;
  const missing = targetColumn
    ? state.rows.filter((row) => shouldPredict(row, targetColumn)).length
    : 0;

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
  const displayHeaders = state.headers.slice(0, 10);
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
  const headers = ['row_number', 'date', 'status', 'original_value', 'predicted_value', 'method', 'sources_or_reason'];
  const lines = [headers.join(',')];

  state.predictions.forEach((item) => {
    lines.push([
      item.rowNumber,
      item.date,
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
    populateSelect(elements.dateColumn, headers, guesses.dateColumn);
    populateSelect(elements.targetColumn, headers, guesses.targetColumn, false);
    populateHistoryColumns(guesses.numericHeaders.filter((header) => header !== guesses.targetColumn), guesses.historyColumns);

    const recommendedMethod = chooseAutoMethod();
    addLog(`Recommended prediction method: ${recommendedMethod}. Auto mode will use this unless you manually choose another method.`);

    elements.fileName.textContent = file.name;
    elements.runButton.disabled = false;
    elements.downloadCsvButton.disabled = true;
    elements.downloadReportButton.disabled = true;

    addLog(`Loaded "${file.name}" as ${uploaded.sourceType} with ${rows.length} row(s) and ${headers.length} column(s).`);
    if (uploaded.sheetName) {
      addLog(`Using workbook sheet: "${uploaded.sheetName}".`);
    }
    if (guesses.targetColumn) {
      addLog(`Detected count column: "${guesses.targetColumn}".`);
    }
    if (guesses.historyColumns.length > 0) {
      addLog(`Detected previous data columns: ${guesses.historyColumns.join(', ')}.`);
    } else {
      addLog('No previous data columns were auto-selected. Please choose them manually.', 'warn');
    }

    updateMetrics();
    renderPreview();
  } catch (error) {
    resetLog(`ERROR ${error.message}`);
    elements.runButton.disabled = true;
    elements.downloadCsvButton.disabled = true;
    elements.downloadReportButton.disabled = true;
  }
}

function syncHistoryAfterTargetChange() {
  const targetColumn = elements.targetColumn.value;
  const selected = getSelectedHistoryColumns().filter((header) => header !== targetColumn);
  const numericHeaders = state.headers.filter((header) => numericProfile(header).ratio >= 0.35 && header !== targetColumn);
  const autoSelected = selected.length > 0 ? selected : numericHeaders.filter(looksLikeHistoryHeader);
  populateHistoryColumns(numericHeaders, autoSelected);
  resetPredictionState();
  updateMetrics();
  renderPreview();
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
elements.targetColumn.addEventListener('change', syncHistoryAfterTargetChange);
elements.zeroMissing.addEventListener('change', updateMetrics);
elements.runButton.addEventListener('click', runPrediction);
elements.clearLogButton.addEventListener('click', () => resetLog('Log cleared.'));

elements.downloadCsvButton.addEventListener('click', () => {
  downloadText(`${baseFileName()}-predicted.csv`, buildCSV(state.resultRows));
});

elements.downloadReportButton.addEventListener('click', () => {
  downloadText(`${baseFileName()}-prediction-report.csv`, buildReportCSV());
});


elements.downloadCsvButton.addEventListener('click', async () => {
  try {
    if (typeof ExcelJS === 'undefined') {
      addLog('Excel export library is not loaded. Check the ExcelJS script in index.html.', 'error');
      return;
    }
    if (!state.resultRows || state.resultRows.length === 0) {
      addLog('No predicted output is available to download.', 'error');
      return;
    }
    addLog('Preparing Excel download...');
    await downloadTrueXlsx(`${baseFileName()}-predicted.xlsx`, state.resultRows);
    addLog('Excel file downloaded successfully.');
  } catch (error) {
    addLog(`Excel download failed: ${error.message}`, 'error');
    console.error(error);
  }
});
