"use strict";

const SAMPLE_FILE_NAME = "סקר הסקרים - 27.4.xlsx";
const TITLE_FONT_FAMILY = '"FbPractica", "Fb Practica", "Heebo", Arial, "Noto Sans Hebrew", "Segoe UI", sans-serif';
const GRAPH_FONT_FAMILY =
  '"FbPracticaNarrow", "Fb Practica Narrow", "FbPractica Narrow", "FbPractica", "Fb Practica", "Heebo", Arial, "Noto Sans Hebrew", "Segoe UI", sans-serif';
const REQUIRED_CANVAS_FONTS = [
  {
    name: "FbPractica",
    family: "FbPractica",
    aliases: ["FbPractica", "Fb Practica"],
    sources: [
      "./fonts/FbPractica.woff2",
      "./fonts/FbPractica.woff",
      "./fonts/FbPractica.ttf",
      "./fonts/FbPractica.otf",
    ],
  },
  {
    name: "FbPracticaNarrow",
    family: "FbPracticaNarrow",
    aliases: ["FbPracticaNarrow", "Fb Practica Narrow", "FbPractica Narrow"],
    sources: [
      "./fonts/FbPracticaNarrow.woff2",
      "./fonts/FbPracticaNarrow.woff",
      "./fonts/FbPracticaNarrow.ttf",
      "./fonts/FbPracticaNarrow.otf",
    ],
  },
];
const TEXT_COLOR = "#231F20";
const BAR_COLORS = ["#2381BE", "#40A1D9"];
const FONT_WEIGHTS = {
  regular: 400,
  bold: 800,
};
const DEFAULT_FONT_SIZES = {
  title: 44,
  subtitle: 35,
  numbers: 24,
  labels: 26,
};
const LAYOUT_FONT_SIZES = {
  numbers: 24,
  labels: 18,
};
const DEFAULT_FONT_BOLD = {
  title: true,
  subtitle: false,
  numbers: true,
  labels: false,
};
const DEFAULT_LABEL_COLUMN_GAP = 8;
const LABEL_DIAGONAL_ANGLE = -Math.PI / 4;
const LABEL_DIAGONAL_SIN = Math.sin(Math.abs(LABEL_DIAGONAL_ANGLE));
const LABEL_DIAGONAL_COS = Math.cos(Math.abs(LABEL_DIAGONAL_ANGLE));
const LABEL_DIAGONAL_ANCHOR_WIDTH_RATIO = 0.55;
const LABEL_OVERLAP_MODES = new Set(["auto", "stacked", "diagonal", "single", "hide"]);
const LABEL_DENSITY_SETTINGS = {
  compact: {
    labelPaddingRatio: 0.04,
    labelPaddingMin: 2,
    labelPaddingScale: 0.3,
    minimumWidthRatio: 0.55,
    minimumWidthMin: 16,
    gapScale: 0.25,
    gapMin: 0,
    lineHeightRatio: 1,
    lineHeightMin: 11,
    laneGapRatio: 0.04,
    laneGapMin: 1,
  },
  balanced: {
    labelPaddingRatio: 0.1,
    labelPaddingMin: 4,
    labelPaddingScale: 0.7,
    minimumWidthRatio: 0.8,
    minimumWidthMin: 22,
    gapScale: 0.8,
    gapMin: 1,
    lineHeightRatio: 1.12,
    lineHeightMin: 12,
    laneGapRatio: 0.12,
    laneGapMin: 3,
  },
  spacious: {
    labelPaddingRatio: 0.18,
    labelPaddingMin: 8,
    labelPaddingScale: 1.3,
    minimumWidthRatio: 1.05,
    minimumWidthMin: 30,
    gapScale: 1.6,
    gapMin: 3,
    lineHeightRatio: 1.22,
    lineHeightMin: 14,
    laneGapRatio: 0.2,
    laneGapMin: 6,
  },
};
const LABEL_DENSITY_MODES = new Set(Object.keys(LABEL_DENSITY_SETTINGS));

const defaultParties = [
  { name: "הליכוד", value: 28.8333333333 },
  { name: "ביחד", value: 24.5 },
  { name: "ישר", value: 12.3333333333 },
  { name: "הרשימה המשותפת", value: 10.5 },
  { name: 'ש"ס', value: 9.6666666667 },
  { name: "הדמוקרטים", value: 9.1666666667 },
  { name: "עוצמה יהודית", value: 8.1666666667 },
  { name: "ישראל ביתנו", value: 7.5 },
  { name: "יהדות התורה", value: 7.3333333333 },
  { name: "הציונות הדתית", value: 2 },
  { name: "המילואימניקים", value: 0 },
  { name: "כחול לבן", value: 0 },
];

const state = {
  workbook: null,
  activeSheetName: "",
  parties: defaultParties.map((party) => ({ ...party })),
  meta: {
    title: "ממוצע ששת הסקרים האחרונים",
    subtitle: "שפורסמו בכלי התקשורת",
    date: "27.4",
  },
  options: {
    sort: true,
    transparent: true,
    width: 1200,
    height: 580,
    labelOverlap: "auto",
    labelDensity: "balanced",
    labelColumnGap: DEFAULT_LABEL_COLUMN_GAP,
    fontSizes: { ...DEFAULT_FONT_SIZES },
    fontBold: { ...DEFAULT_FONT_BOLD },
  },
};

const els = {
  sampleButton: document.querySelector("#sampleButton"),
  fileInput: document.querySelector("#fileInput"),
  exportButton: document.querySelector("#exportButton"),
  status: document.querySelector("#status"),
  sheetSelect: document.querySelector("#sheetSelect"),
  titleInput: document.querySelector("#titleInput"),
  subtitleInput: document.querySelector("#subtitleInput"),
  dateInput: document.querySelector("#dateInput"),
  widthInput: document.querySelector("#widthInput"),
  heightInput: document.querySelector("#heightInput"),
  titleFontSizeInput: document.querySelector("#titleFontSizeInput"),
  subtitleFontSizeInput: document.querySelector("#subtitleFontSizeInput"),
  numberFontSizeInput: document.querySelector("#numberFontSizeInput"),
  partyLabelFontSizeInput: document.querySelector("#partyLabelFontSizeInput"),
  labelOverlapSelect: document.querySelector("#labelOverlapSelect"),
  labelDensitySelect: document.querySelector("#labelDensitySelect"),
  labelColumnGapInput: document.querySelector("#labelColumnGapInput"),
  titleBoldInput: document.querySelector("#titleBoldInput"),
  subtitleBoldInput: document.querySelector("#subtitleBoldInput"),
  numberBoldInput: document.querySelector("#numberBoldInput"),
  partyLabelBoldInput: document.querySelector("#partyLabelBoldInput"),
  sortToggle: document.querySelector("#sortToggle"),
  transparentToggle: document.querySelector("#transparentToggle"),
  addRowButton: document.querySelector("#addRowButton"),
  partyTable: document.querySelector("#partyTable"),
  summaryText: document.querySelector("#summaryText"),
  chartCanvas: document.querySelector("#chartCanvas"),
};

let fontAvailabilityPromise = null;

class ZipReader {
  constructor(arrayBuffer) {
    this.bytes = new Uint8Array(arrayBuffer);
    this.view = new DataView(arrayBuffer);
    this.decoder = new TextDecoder("utf-8");
    this.entries = this.readEntries();
  }

  readEntries() {
    const eocdOffset = this.findEndOfCentralDirectory();
    const entryCount = this.view.getUint16(eocdOffset + 10, true);
    const centralDirectoryOffset = this.view.getUint32(eocdOffset + 16, true);
    const entries = new Map();
    let offset = centralDirectoryOffset;

    for (let index = 0; index < entryCount; index += 1) {
      if (this.view.getUint32(offset, true) !== 0x02014b50) {
        throw new Error("מבנה ZIP לא תקין בקובץ ה-XLSX.");
      }

      const compressionMethod = this.view.getUint16(offset + 10, true);
      const compressedSize = this.view.getUint32(offset + 20, true);
      const uncompressedSize = this.view.getUint32(offset + 24, true);
      const fileNameLength = this.view.getUint16(offset + 28, true);
      const extraLength = this.view.getUint16(offset + 30, true);
      const commentLength = this.view.getUint16(offset + 32, true);
      const localHeaderOffset = this.view.getUint32(offset + 42, true);
      const nameStart = offset + 46;
      const nameBytes = this.bytes.subarray(nameStart, nameStart + fileNameLength);
      const name = this.decoder.decode(nameBytes);

      entries.set(name, {
        name,
        compressionMethod,
        compressedSize,
        uncompressedSize,
        localHeaderOffset,
      });

      offset = nameStart + fileNameLength + extraLength + commentLength;
    }

    return entries;
  }

  findEndOfCentralDirectory() {
    const minOffset = Math.max(0, this.bytes.length - 0xffff - 22);
    for (let offset = this.bytes.length - 22; offset >= minOffset; offset -= 1) {
      if (this.view.getUint32(offset, true) === 0x06054b50) {
        return offset;
      }
    }
    throw new Error("לא נמצא מבנה ZIP תקין בקובץ.");
  }

  async text(path) {
    const data = await this.extract(path);
    return this.decoder.decode(data);
  }

  async extract(path) {
    const entry = this.entries.get(path);
    if (!entry) {
      throw new Error(`חסר קובץ פנימי ב-XLSX: ${path}`);
    }

    const offset = entry.localHeaderOffset;
    if (this.view.getUint32(offset, true) !== 0x04034b50) {
      throw new Error("מבנה ZIP לא תקין בקובץ ה-XLSX.");
    }

    const fileNameLength = this.view.getUint16(offset + 26, true);
    const extraLength = this.view.getUint16(offset + 28, true);
    const dataStart = offset + 30 + fileNameLength + extraLength;
    const compressed = this.bytes.subarray(dataStart, dataStart + entry.compressedSize);

    if (entry.compressionMethod === 0) {
      return compressed;
    }

    if (entry.compressionMethod !== 8) {
      throw new Error("סוג דחיסת ZIP לא נתמך בקובץ.");
    }

    const inflated = await inflateRaw(compressed);
    if (entry.uncompressedSize && inflated.byteLength !== entry.uncompressedSize) {
      return new Uint8Array(inflated);
    }
    return new Uint8Array(inflated);
  }
}

async function inflateRaw(bytes) {
  if (!("DecompressionStream" in window)) {
    throw new Error("הדפדפן הזה לא תומך בקריאת XLSX מקומית. מומלץ לפתוח ב-Chrome או Edge עדכני.");
  }

  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  const buffer = await new Response(stream).arrayBuffer();
  return buffer;
}

function parseXml(text) {
  const xml = new DOMParser().parseFromString(text, "application/xml");
  const parserError = xml.querySelector("parsererror");
  if (parserError) {
    throw new Error("לא ניתן לקרוא XML פנימי מתוך קובץ ה-XLSX.");
  }
  return xml;
}

async function parseXlsx(file) {
  const zip = new ZipReader(await file.arrayBuffer());
  const sharedStrings = await readSharedStrings(zip);
  const workbookXml = parseXml(await zip.text("xl/workbook.xml"));
  const relsXml = parseXml(await zip.text("xl/_rels/workbook.xml.rels"));
  const rels = readRelationships(relsXml);
  const sheetNodes = Array.from(workbookXml.getElementsByTagName("sheet"));
  const sheets = [];

  for (const sheetNode of sheetNodes) {
    const name = sheetNode.getAttribute("name") || "Sheet";
    const relId =
      sheetNode.getAttribute("r:id") ||
      sheetNode.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    const target = rels.get(relId);
    if (!target) {
      continue;
    }

    const path = resolveRelationshipPath("xl/workbook.xml", target);
    const rows = parseSheetXml(await zip.text(path), sharedStrings);
    sheets.push({ name, path, rows });
  }

  if (!sheets.length) {
    throw new Error("לא נמצאו גליונות בקובץ.");
  }

  return { fileName: file.name, sheets };
}

async function readSharedStrings(zip) {
  if (!zip.entries.has("xl/sharedStrings.xml")) {
    return [];
  }

  const xml = parseXml(await zip.text("xl/sharedStrings.xml"));
  return Array.from(xml.getElementsByTagName("si")).map((node) => node.textContent || "");
}

function readRelationships(xml) {
  const rels = new Map();
  for (const node of Array.from(xml.getElementsByTagName("Relationship"))) {
    rels.set(node.getAttribute("Id"), node.getAttribute("Target"));
  }
  return rels;
}

function resolveRelationshipPath(sourcePath, target) {
  if (target.startsWith("/")) {
    return normalizeZipPath(target.slice(1));
  }

  const sourceParts = sourcePath.split("/");
  sourceParts.pop();
  return normalizeZipPath([...sourceParts, target].join("/"));
}

function normalizeZipPath(path) {
  const parts = [];
  for (const part of path.split("/")) {
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      parts.pop();
    } else {
      parts.push(part);
    }
  }
  return parts.join("/");
}

function parseSheetXml(text, sharedStrings) {
  const xml = parseXml(text);
  const rows = [];

  for (const rowNode of Array.from(xml.getElementsByTagName("row"))) {
    const rowIndex = Number(rowNode.getAttribute("r")) || rows.length + 1;
    const row = [];

    for (const cellNode of Array.from(rowNode.getElementsByTagName("c"))) {
      const ref = cellNode.getAttribute("r") || "";
      const columnLetters = ref.match(/[A-Z]+/i)?.[0] || "A";
      const columnIndex = lettersToColumnIndex(columnLetters);
      row[columnIndex] = readCellValue(cellNode, sharedStrings);
    }

    rows[rowIndex - 1] = row;
  }

  return rows.map((row) => row || []);
}

function readCellValue(cellNode, sharedStrings) {
  const type = cellNode.getAttribute("t");
  const valueNode = cellNode.getElementsByTagName("v")[0];

  if (type === "s") {
    const index = Number(valueNode?.textContent || 0);
    return sharedStrings[index] || "";
  }

  if (type === "inlineStr") {
    return cellNode.getElementsByTagName("is")[0]?.textContent || "";
  }

  if (!valueNode) {
    return "";
  }

  const raw = valueNode.textContent || "";
  const numeric = parseNumber(raw);
  return Number.isFinite(numeric) ? numeric : raw;
}

function lettersToColumnIndex(letters) {
  let index = 0;
  for (const letter of letters.toUpperCase()) {
    index = index * 26 + letter.charCodeAt(0) - 64;
  }
  return index - 1;
}

function parseNumber(value) {
  if (typeof value === "number") {
    return value;
  }

  const normalized = String(value)
    .trim()
    .replace(/,/g, "")
    .replace(/[^\d.+-]/g, "");

  if (!normalized) {
    return NaN;
  }

  return Number(normalized);
}

function applyWorkbook(workbook, preferredSheetName = "") {
  state.workbook = workbook;
  renderSheetOptions(workbook.sheets, preferredSheetName);
  let sheet = getActiveSheet();
  let extraction = extractSheetData(sheet, workbook);

  if (!extraction.parties.length) {
    const fallback = workbook.sheets
      .map((candidate) => ({
        sheet: candidate,
        extraction: extractSheetData(candidate, workbook),
      }))
      .find((candidate) => candidate.extraction.parties.length);

    if (fallback) {
      sheet = fallback.sheet;
      extraction = fallback.extraction;
      state.activeSheetName = sheet.name;
      els.sheetSelect.value = sheet.name;
    }
  }

  state.parties = extraction.parties.length ? extraction.parties : defaultParties.map((party) => ({ ...party }));
  state.meta = {
    title: extraction.meta.title || state.meta.title,
    subtitle: extraction.meta.subtitle || state.meta.subtitle,
    date: extraction.meta.date || state.meta.date,
  };

  syncControls();
  renderPartyTable();
  renderChart();
  setStatus(`נטען: ${workbook.fileName}`);
}

function renderSheetOptions(sheets, preferredSheetName = "") {
  const preferred =
    sheets.find((sheet) => sheet.name === preferredSheetName) ||
    sheets.find((sheet) => /ממוצע|average|avg/i.test(sheet.name)) ||
    sheets[0];

  state.activeSheetName = preferred.name;
  els.sheetSelect.innerHTML = "";

  for (const sheet of sheets) {
    const option = document.createElement("option");
    option.value = sheet.name;
    option.textContent = sheet.name;
    option.selected = sheet.name === preferred.name;
    els.sheetSelect.append(option);
  }
}

function getActiveSheet() {
  return state.workbook?.sheets.find((sheet) => sheet.name === state.activeSheetName) || null;
}

function extractSheetData(sheet, workbook) {
  if (!sheet) {
    return { parties: [], meta: { ...state.meta } };
  }

  const header = findHeader(sheet.rows);
  const meta = extractMetadata(sheet.rows, header.index, workbook);
  if (header.index < 0) {
    return { parties: [], meta };
  }

  const headerRow = sheet.rows[header.index] || [];
  const valueMode = findValueMode(sheet.rows, header.index, header.nameColumn, headerRow);
  const parties = [];

  for (let rowIndex = header.index + 1; rowIndex < sheet.rows.length; rowIndex += 1) {
    const row = sheet.rows[rowIndex] || [];
    if (isBlankRow(row)) {
      continue;
    }

    const name = cleanText(row[header.nameColumn]);
    if (!name || isSummaryRow(name)) {
      continue;
    }

    const value =
      valueMode.type === "single"
        ? parseNumber(row[valueMode.column])
        : averageNumericCells(row, header.nameColumn, valueMode.columns);

    if (Number.isFinite(value) && value >= 0) {
      parties.push({ name, value });
    }
  }

  return { parties, meta };
}

function findHeader(rows) {
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const nameColumn = row.findIndex((cell) => /מפלגה|party/i.test(cleanText(cell)));
    if (nameColumn >= 0) {
      return { index: rowIndex, nameColumn };
    }
  }

  for (let rowIndex = 0; rowIndex < rows.length - 1; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const nextRow = rows[rowIndex + 1] || [];
    if (cleanText(row[0]) && nextRow.slice(1).some((cell) => Number.isFinite(parseNumber(cell)))) {
      return { index: rowIndex, nameColumn: 0 };
    }
  }

  return { index: -1, nameColumn: 0 };
}

function findValueMode(rows, headerIndex, nameColumn, headerRow) {
  const averageColumn = headerRow.findIndex((cell, index) => {
    if (index === nameColumn) {
      return false;
    }
    return /ממוצע|average|avg/i.test(cleanText(cell));
  });

  if (averageColumn >= 0) {
    return { type: "single", column: averageColumn, columns: [averageColumn] };
  }

  const columns = [];
  const maxColumns = Math.max(...rows.map((row) => row.length), 0);

  for (let column = 0; column < maxColumns; column += 1) {
    if (column === nameColumn) {
      continue;
    }

    let numericCount = 0;
    for (let rowIndex = headerIndex + 1; rowIndex < Math.min(rows.length, headerIndex + 10); rowIndex += 1) {
      const row = rows[rowIndex] || [];
      if (Number.isFinite(parseNumber(row[column]))) {
        numericCount += 1;
      }
    }

    if (numericCount >= 2) {
      columns.push(column);
    }
  }

  if (columns.length === 1) {
    return { type: "single", column: columns[0], columns };
  }

  return { type: "average", column: columns[0] ?? 1, columns };
}

function averageNumericCells(row, nameColumn, preferredColumns) {
  const values = [];
  const columns = preferredColumns.length ? preferredColumns : row.map((_, index) => index);

  for (const column of columns) {
    if (column === nameColumn) {
      continue;
    }
    const value = parseNumber(row[column]);
    if (Number.isFinite(value)) {
      values.push(value);
    }
  }

  if (!values.length) {
    return NaN;
  }

  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function extractMetadata(rows, headerIndex, workbook) {
  const firstRows = rows.slice(0, Math.max(headerIndex, 4));
  const firstColumnValues = firstRows.map((row) => cleanText(firstNonEmpty(row))).filter(Boolean);
  const [rawTitle = "", rawSubtitle = "", rawDate = ""] = firstColumnValues;
  const detectedPollCount = detectPollCount(workbook);

  return {
    title: normalizeTitle(rawTitle, detectedPollCount),
    subtitle: rawSubtitle || "שפורסמו בכלי התקשורת",
    date: rawDate ? formatValue(rawDate) : "",
  };
}

function detectPollCount(workbook) {
  if (!workbook) {
    return 0;
  }

  for (const sheet of workbook.sheets) {
    if (/ממוצע|average|avg/i.test(sheet.name)) {
      continue;
    }
    const header = findHeader(sheet.rows);
    if (header.index >= 0) {
      const mode = findValueMode(sheet.rows, header.index, header.nameColumn, sheet.rows[header.index] || []);
      if (mode.type === "average" && mode.columns.length) {
        return mode.columns.length;
      }
    }
  }

  return 0;
}

function normalizeTitle(rawTitle, pollCount) {
  let title = cleanText(rawTitle);
  if (title.startsWith("נתוני ")) {
    title = title.replace("נתוני", "ממוצע");
  }

  if (!title) {
    return pollCount ? `ממוצע ${pollCount} הסקרים האחרונים` : "ממוצע הסקרים האחרונים";
  }

  if (pollCount === 6 && /\b6\b/.test(title)) {
    return title.replace(/\b6\b/, "ששת");
  }

  return title;
}

function firstNonEmpty(row) {
  return (row || []).find((cell) => cleanText(cell));
}

function cleanText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function formatValue(value) {
  if (typeof value === "number") {
    return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(3))).replace(/\.0$/, "");
  }
  return cleanText(value);
}

function formatEditableNumber(value) {
  const numeric = parseNumber(value);
  if (!Number.isFinite(numeric)) {
    return "";
  }
  return Number.isInteger(numeric) ? String(numeric) : String(Number(numeric.toFixed(1)));
}

function isBlankRow(row) {
  return !(row || []).some((cell) => cleanText(cell));
}

function isSummaryRow(name) {
  return /קואליציה|אופוזיציה|סה.?כ|total|coalition|opposition/i.test(name);
}

function syncControls() {
  els.titleInput.value = state.meta.title;
  els.subtitleInput.value = state.meta.subtitle;
  els.dateInput.value = state.meta.date;
  els.widthInput.value = state.options.width;
  els.heightInput.value = state.options.height;
  els.titleFontSizeInput.value = state.options.fontSizes.title;
  els.subtitleFontSizeInput.value = state.options.fontSizes.subtitle;
  els.numberFontSizeInput.value = state.options.fontSizes.numbers;
  els.partyLabelFontSizeInput.value = state.options.fontSizes.labels;
  els.labelOverlapSelect.value = state.options.labelOverlap;
  els.labelDensitySelect.value = state.options.labelDensity;
  els.labelColumnGapInput.value = state.options.labelColumnGap;
  els.titleBoldInput.checked = state.options.fontBold.title;
  els.subtitleBoldInput.checked = state.options.fontBold.subtitle;
  els.numberBoldInput.checked = state.options.fontBold.numbers;
  els.partyLabelBoldInput.checked = state.options.fontBold.labels;
  els.sortToggle.checked = state.options.sort;
  els.transparentToggle.checked = state.options.transparent;
}

function renderPartyTable() {
  els.partyTable.innerHTML = "";
  const fragment = document.createDocumentFragment();

  state.parties.forEach((party, index) => {
    const row = document.createElement("tr");
    row.dataset.index = String(index);

    const nameCell = document.createElement("td");
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.value = party.name;
    nameInput.dir = "rtl";
    nameInput.dataset.field = "name";
    nameInput.setAttribute("aria-label", "שם מפלגה");
    nameCell.append(nameInput);

    const valueCell = document.createElement("td");
    const valueInput = document.createElement("input");
    valueInput.type = "number";
    valueInput.min = "0";
    valueInput.step = "0.1";
    valueInput.value = formatEditableNumber(party.value);
    valueInput.dir = "ltr";
    valueInput.dataset.field = "value";
    valueInput.setAttribute("aria-label", "מספר מנדטים");
    valueCell.append(valueInput);

    const actionCell = document.createElement("td");
    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "remove-row";
    removeButton.dataset.action = "remove";
    removeButton.title = "מחיקה";
    removeButton.setAttribute("aria-label", "מחיקת שורה");
    removeButton.textContent = "×";
    actionCell.append(removeButton);

    row.append(nameCell, valueCell, actionCell);
    fragment.append(row);
  });

  els.partyTable.append(fragment);
  updateSummary();
}

function renderChart() {
  state.meta.title = els.titleInput.value.trim();
  state.meta.subtitle = els.subtitleInput.value.trim();
  state.meta.date = els.dateInput.value.trim();
  state.options.sort = els.sortToggle.checked;
  state.options.transparent = els.transparentToggle.checked;
  state.options.width = clamp(Number(els.widthInput.value) || 1200, 420, 5000);
  state.options.height = clamp(Number(els.heightInput.value) || 620, 240, 1400);
  state.options.fontSizes = {
    title: readFontSize(els.titleFontSizeInput, DEFAULT_FONT_SIZES.title, 12, 96),
    subtitle: readFontSize(els.subtitleFontSizeInput, DEFAULT_FONT_SIZES.subtitle, 10, 72),
    numbers: readFontSize(els.numberFontSizeInput, DEFAULT_FONT_SIZES.numbers, 8, 72),
    labels: readFontSize(els.partyLabelFontSizeInput, DEFAULT_FONT_SIZES.labels, 8, 64),
  };
  state.options.labelOverlap = readSelectValue(els.labelOverlapSelect, LABEL_OVERLAP_MODES, "auto");
  state.options.labelDensity = readSelectValue(els.labelDensitySelect, LABEL_DENSITY_MODES, "balanced");
  state.options.labelColumnGap = readNumberInput(
    els.labelColumnGapInput,
    DEFAULT_LABEL_COLUMN_GAP,
    0,
    160,
  );
  state.options.fontBold = {
    title: els.titleBoldInput.checked,
    subtitle: els.subtitleBoldInput.checked,
    numbers: els.numberBoldInput.checked,
    labels: els.partyLabelBoldInput.checked,
  };

  const canvas = els.chartCanvas;
  if (canvas.width !== state.options.width) {
    canvas.width = state.options.width;
  }
  if (canvas.height !== state.options.height) {
    canvas.height = state.options.height;
  }

  let ctx = canvas.getContext("2d");
  let width = canvas.width;
  const height = canvas.height;
  let scale = Math.min(width / 550, height / 298);
  const data = getRenderableParties();

  let titleFontSize = state.options.fontSizes.title;
  let subtitleFontSize = state.options.fontSizes.subtitle;
  let valueFontSize = state.options.fontSizes.numbers;
  let labelFontSize = state.options.fontSizes.labels;
  let titleY = 24 * scale;
  let subtitleY = titleY + 22 * scale;
  let sidePadding = Math.max(16, 14 * scale);
  const labelSettings = getLabelDensitySettings(state.options.labelDensity);
  let labelMetrics = data.length ? measureLabelMetrics(ctx, data, labelFontSize, scale, labelSettings) : [];

  ctx.clearRect(0, 0, width, height);
  if (!state.options.transparent) {
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);
  }

  ctx.direction = "rtl";
  ctx.textAlign = "center";
  ctx.textBaseline = "alphabetic";
  ctx.fillStyle = TEXT_COLOR;

  drawFittedText(
    ctx,
    state.meta.title,
    width / 2,
    titleY,
    width - 34 * scale,
    titleFontSize,
    Math.max(8, titleFontSize * 0.55),
    getTextWeight("title"),
    TITLE_FONT_FAMILY,
  );
  const subtitle = state.meta.date ? `${state.meta.subtitle} (${state.meta.date})` : state.meta.subtitle;
  drawFittedText(
    ctx,
    subtitle,
    width / 2,
    subtitleY,
    width - 34 * scale,
    subtitleFontSize,
    Math.max(8, subtitleFontSize * 0.55),
    getTextWeight("subtitle"),
    TITLE_FONT_FAMILY,
  );

  if (!data.length) {
    drawFittedText(
      ctx,
      "אין נתונים להצגה",
      width / 2,
      height / 2,
      width - 40,
      18 * scale,
      12 * scale,
      getTextWeight("title"),
      GRAPH_FONT_FAMILY,
    );
    updateSummary();
    return;
  }

  const drawnLabelLineHeight = Math.max(labelSettings.lineHeightMin, labelFontSize * labelSettings.lineHeightRatio);
  const commonSegmentWidth = (width - sidePadding * 2) / data.length;
  const barWidth = clamp(commonSegmentWidth * 0.38, 10 * scale, 32 * scale);
  const barCenters = data.map((_, index) => sidePadding + commonSegmentWidth * index + commonSegmentWidth / 2);
  const labelGap = Math.max(labelSettings.gapMin, labelSettings.gapScale * scale);
  const labelLayout = resolveLabelLayout(
    barCenters,
    labelMetrics,
    0,
    width,
    labelGap,
    state.options.labelOverlap,
  );
  const labelLaneGap = Math.max(labelSettings.laneGapMin, labelFontSize * labelSettings.laneGapRatio);
  const labelLaneMetrics = labelLayout.diagonal
    ? getDiagonalLabelLaneMetrics(labelMetrics, labelLayout.visible)
    : getLabelLaneMetrics(
        labelMetrics,
        labelLayout.lanes,
        drawnLabelLineHeight,
        labelLaneGap,
        labelLayout.visible,
      );
  const labelTopGap = state.options.labelColumnGap;
  const labelBottomGap = Math.max(8, 3 * scale);
  const labelArea = Math.max(labelTopGap + labelLaneMetrics.totalHeight + labelBottomGap, 64);
  const barBase = height - labelArea;
  const valueBandBottom = subtitleY + 42 * scale;
  let chartTop = Math.max(valueBandBottom, 94 * scale);
  const minChartHeight = 86 * scale;
  if (barBase - chartTop < minChartHeight) {
    chartTop = Math.max(subtitleY + 30 * scale, barBase - minChartHeight);
  }
  const roundedValues = data.map((party) => Math.max(0, Math.round(party.value)));
  const maxValue = Math.max(1, ...roundedValues);
  const chartHeight = barBase - chartTop;
  const valueGap = Math.max(10, LAYOUT_FONT_SIZES.numbers * 0.48);

  data.forEach((party, index) => {
    const value = roundedValues[index];
    const centerX = barCenters[index];
    const x = centerX - barWidth / 2;
    const barHeight = value ? (chartHeight * value) / maxValue : 0;
    const y = barBase - barHeight;

    if (value > 0) {
      ctx.fillStyle = BAR_COLORS[index % BAR_COLORS.length];
      ctx.fillRect(Math.round(x), Math.round(y), Math.max(1, Math.round(barWidth)), Math.round(barHeight));
    }

    ctx.fillStyle = TEXT_COLOR;
    const valueY = value > 0 ? y - valueGap : barBase - 4 * scale;
    drawFittedText(
      ctx,
      String(value),
      centerX,
      valueY,
      Math.max(24 * scale, commonSegmentWidth),
      valueFontSize,
      10 * scale,
      getTextWeight("numbers"),
      GRAPH_FONT_FAMILY,
    );

    if (labelLayout.visible[index]) {
      const labelY = barBase + labelTopGap + labelLaneMetrics.offsets[labelLayout.lanes[index]];
      if (labelLayout.diagonal) {
        drawDiagonalPartyLabel(
          ctx,
          party.name,
          labelLayout.centers[index],
          labelY,
          labelFontSize,
          labelMetrics[index],
        );
      } else {
        drawPartyLabel(
          ctx,
          party.name,
          labelLayout.centers[index],
          labelY,
          labelFontSize,
          drawnLabelLineHeight,
        );
      }
    }
  });

  updateSummary();
}

function getRenderableParties() {
  const parties = state.parties
    .map((party) => ({
      name: cleanText(party.name),
      value: Math.max(0, parseNumber(party.value)),
    }))
    .filter((party) => party.name && Number.isFinite(party.value));

  if (state.options.sort) {
    parties.sort((a, b) => b.value - a.value);
  }

  return parties;
}

function drawFittedText(ctx, text, x, y, maxWidth, startSize, minSize, weight, fontFamily) {
  const clean = cleanText(text);
  if (!clean) {
    return;
  }

  let size = startSize;
  while (size > minSize) {
    ctx.font = `${weight} ${Math.round(size)}px ${fontFamily}`;
    if (ctx.measureText(clean).width <= maxWidth) {
      break;
    }
    size -= 1;
  }
  ctx.font = `${weight} ${Math.max(Math.round(size), Math.round(minSize))}px ${fontFamily}`;
  ctx.fillText(clean, x, y);
}

function measureLabelMetrics(ctx, data, fontSize, scale, settings) {
  const labelPadding = Math.max(
    settings.labelPaddingMin,
    fontSize * settings.labelPaddingRatio,
    scale * settings.labelPaddingScale,
  );
  const minimumSegmentWidth = Math.max(settings.minimumWidthMin, fontSize * settings.minimumWidthRatio);
  const lineHeight = Math.max(settings.lineHeightMin, fontSize * settings.lineHeightRatio);

  ctx.font = `${getTextWeight("labels")} ${Math.max(8, Math.round(fontSize))}px ${GRAPH_FONT_FAMILY}`;

  return data.map((party) => {
    const label = cleanText(party.name);
    const lines = getPartyLabelLines(label);
    const widestLine = Math.max(...lines.map((line) => ctx.measureText(line).width));
    const singleLineWidth = ctx.measureText(label).width;
    const diagonalBox = getDiagonalTextBox(singleLineWidth, lineHeight);
    return {
      lines,
      label,
      diagonalCenterOffset: diagonalBox.centerOffset,
      diagonalHeight: Math.ceil(diagonalBox.height + labelPadding),
      diagonalAnchorWidth: Math.max(4, Math.ceil(lineHeight * LABEL_DIAGONAL_ANCHOR_WIDTH_RATIO)),
      width: Math.max(minimumSegmentWidth, Math.ceil(widestLine + labelPadding)),
    };
  });
}

function getDiagonalTextBox(textWidth, textHeight) {
  return {
    centerOffset: (textHeight * LABEL_DIAGONAL_SIN - textWidth * LABEL_DIAGONAL_COS) / 2,
    height: textWidth * LABEL_DIAGONAL_SIN + textHeight * LABEL_DIAGONAL_COS,
    width: textWidth * LABEL_DIAGONAL_COS + textHeight * LABEL_DIAGONAL_SIN,
  };
}

function resolveLabelLayout(baseCenters, labelMetrics, minX, maxX, gap, mode) {
  if (mode === "single") {
    return resolveSingleRowLabelLayout(baseCenters, labelMetrics, minX, maxX);
  }

  if (mode === "stacked") {
    return resolveStackedLabelLayout(baseCenters, labelMetrics, minX, maxX, gap);
  }

  if (mode === "diagonal") {
    return resolveDiagonalLabelLayout(baseCenters, labelMetrics, minX, maxX);
  }

  if (mode === "hide") {
    return resolveHiddenLabelLayout(baseCenters, labelMetrics, minX, maxX, gap);
  }

  const centers = resolveLabelCenters(baseCenters, labelMetrics, minX, maxX, gap);
  if (!hasLabelOverlap(centers, labelMetrics, gap)) {
    return {
      centers,
      lanes: new Array(baseCenters.length).fill(0),
      visible: new Array(baseCenters.length).fill(true),
    };
  }

  return resolveStackedLabelLayout(baseCenters, labelMetrics, minX, maxX, gap);
}

function resolveSingleRowLabelLayout(baseCenters, labelMetrics, minX, maxX) {
  return {
    centers: baseCenters.map((center, index) =>
      clampLabelCenter(center, minX, maxX, labelMetrics[index].width / 2),
    ),
    lanes: new Array(baseCenters.length).fill(0),
    visible: new Array(baseCenters.length).fill(true),
  };
}

function resolveHiddenLabelLayout(baseCenters, labelMetrics, minX, maxX, gap) {
  const centers = resolveLabelCenters(baseCenters, labelMetrics, minX, maxX, gap);
  const visible = new Array(baseCenters.length).fill(true);
  let previousRight = Number.NEGATIVE_INFINITY;

  centers.forEach((center, index) => {
    const halfWidth = labelMetrics[index].width / 2;
    const left = center - halfWidth;
    const right = center + halfWidth;
    if (left < previousRight + gap) {
      visible[index] = false;
      return;
    }
    previousRight = right;
  });

  return {
    centers,
    lanes: new Array(baseCenters.length).fill(0),
    visible,
  };
}

function resolveDiagonalLabelLayout(baseCenters, labelMetrics, minX, maxX) {
  return {
    centers: baseCenters.map((center, index) =>
      clampLabelCenter(center, minX, maxX, labelMetrics[index].diagonalAnchorWidth / 2),
    ),
    lanes: new Array(baseCenters.length).fill(0),
    visible: new Array(baseCenters.length).fill(true),
    diagonal: true,
  };
}

function resolveLabelCenters(baseCenters, labelMetrics, minX, maxX, gap) {
  const centers = [...baseCenters];
  const widths = labelMetrics.map((metric) => metric.width);
  const count = centers.length;

  if (!count) {
    return centers;
  }

  for (let pass = 0; pass < 4; pass += 1) {
    for (let index = 1; index < count; index += 1) {
      const previousRight = centers[index - 1] + widths[index - 1] / 2;
      const currentLeft = centers[index] - widths[index] / 2;
      if (currentLeft < previousRight + gap) {
        centers[index] += previousRight + gap - currentLeft;
      }
    }

    const overflowRight = centers[count - 1] + widths[count - 1] / 2 - maxX;
    if (overflowRight > 0) {
      for (let index = 0; index < count; index += 1) {
        centers[index] -= overflowRight;
      }
    }

    for (let index = count - 2; index >= 0; index -= 1) {
      const nextLeft = centers[index + 1] - widths[index + 1] / 2;
      const currentRight = centers[index] + widths[index] / 2;
      if (currentRight > nextLeft - gap) {
        centers[index] -= currentRight - (nextLeft - gap);
      }
    }

    const overflowLeft = minX - (centers[0] - widths[0] / 2);
    if (overflowLeft > 0) {
      for (let index = 0; index < count; index += 1) {
        centers[index] += overflowLeft;
      }
    }
  }

  return centers;
}

function hasLabelOverlap(centers, labelMetrics, gap) {
  for (let index = 1; index < centers.length; index += 1) {
    const previousRight = centers[index - 1] + labelMetrics[index - 1].width / 2;
    const currentLeft = centers[index] - labelMetrics[index].width / 2;
    if (currentLeft < previousRight + gap) {
      return true;
    }
  }

  return false;
}

function resolveStackedLabelLayout(baseCenters, labelMetrics, minX, maxX, gap) {
  const centers = new Array(baseCenters.length);
  const lanes = new Array(baseCenters.length);
  const laneRightEdges = [];

  baseCenters.forEach((baseCenter, index) => {
    const width = labelMetrics[index].width;
    const halfWidth = width / 2;
    let bestLane = -1;
    let bestCenter = baseCenter;
    let bestDistance = Number.POSITIVE_INFINITY;

    laneRightEdges.forEach((rightEdge, laneIndex) => {
      const minimumCenter = rightEdge + gap + halfWidth;
      const center = clampLabelCenter(Math.max(baseCenter, minimumCenter), minX, maxX, halfWidth);
      const left = center - halfWidth;
      const right = center + halfWidth;

      if (left >= rightEdge + gap && right <= maxX) {
        const distance = Math.abs(center - baseCenter);
        if (distance < bestDistance) {
          bestLane = laneIndex;
          bestCenter = center;
          bestDistance = distance;
        }
      }
    });

    if (bestLane === -1) {
      bestLane = laneRightEdges.length;
      bestCenter = clampLabelCenter(baseCenter, minX, maxX, halfWidth);
    }

    centers[index] = bestCenter;
    lanes[index] = bestLane;
    laneRightEdges[bestLane] = bestCenter + halfWidth;
  });

  return {
    centers,
    lanes,
    visible: new Array(baseCenters.length).fill(true),
  };
}

function clampLabelCenter(center, minX, maxX, halfWidth) {
  const minCenter = minX + halfWidth;
  const maxCenter = maxX - halfWidth;

  if (minCenter > maxCenter) {
    return (minX + maxX) / 2;
  }

  return clamp(center, minCenter, maxCenter);
}

function getLabelLaneMetrics(labelMetrics, lanes, lineHeight, laneGap, visible) {
  const laneCount = Math.max(1, ...lanes.map((lane) => lane + 1));
  const heights = new Array(laneCount).fill(0);

  labelMetrics.forEach((metric, index) => {
    if (visible[index] === false) {
      return;
    }
    const lane = lanes[index] || 0;
    heights[lane] = Math.max(heights[lane], metric.lines.length * lineHeight);
  });

  const offsets = [];
  let offset = 0;
  heights.forEach((height, index) => {
    offsets[index] = offset;
    offset += height;
    if (index < heights.length - 1) {
      offset += laneGap;
    }
  });

  return {
    offsets,
    totalHeight: offset,
  };
}

function getDiagonalLabelLaneMetrics(labelMetrics, visible) {
  const totalHeight = labelMetrics.reduce((height, metric, index) => {
    if (visible[index] === false) {
      return height;
    }
    return Math.max(height, metric.diagonalHeight);
  }, 0);

  return {
    offsets: [0],
    totalHeight,
  };
}

function getPartyLabelLines(label) {
  const words = cleanText(label).split(/\s+/).filter(Boolean);
  return words.length ? words : [""];
}

function drawPartyLabel(ctx, label, centerX, firstLineY, fontSize, lineHeight) {
  const previousBaseline = ctx.textBaseline;
  ctx.textBaseline = "top";
  ctx.font = `${getTextWeight("labels")} ${Math.max(8, Math.round(fontSize))}px ${GRAPH_FONT_FAMILY}`;
  ctx.fillStyle = TEXT_COLOR;
  getPartyLabelLines(label).forEach((line, index) => {
    ctx.fillText(line, centerX, firstLineY + index * lineHeight);
  });
  ctx.textBaseline = previousBaseline;
}

function drawDiagonalPartyLabel(ctx, label, centerX, topY, fontSize, metric) {
  const clean = metric?.label || cleanText(label);
  if (!clean) {
    return;
  }

  const previousDirection = ctx.direction;
  const previousAlign = ctx.textAlign;
  const previousBaseline = ctx.textBaseline;

  ctx.save();
  ctx.translate(centerX - (metric?.diagonalCenterOffset ?? 0), topY);
  ctx.rotate(LABEL_DIAGONAL_ANGLE);
  ctx.direction = "rtl";
  ctx.textAlign = "right";
  ctx.textBaseline = "top";
  ctx.font = `${getTextWeight("labels")} ${Math.max(8, Math.round(fontSize))}px ${GRAPH_FONT_FAMILY}`;
  ctx.fillStyle = TEXT_COLOR;
  ctx.fillText(clean, 0, 0);
  ctx.restore();

  ctx.direction = previousDirection;
  ctx.textAlign = previousAlign;
  ctx.textBaseline = previousBaseline;
}

function measureText(ctx, text, fontSize, weight) {
  ctx.font = `${weight} ${Math.max(8, Math.round(fontSize))}px ${GRAPH_FONT_FAMILY}`;
  return ctx.measureText(text).width;
}

function updateSummary() {
  const data = getRenderableParties();
  const total = data.reduce((sum, party) => sum + Math.round(party.value), 0);
  els.summaryText.textContent = `${data.length} מפלגות, ${total} מנדטים מוצגים`;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function readNumberInput(input, fallback, min, max) {
  const value = Number(input.value);
  const next = Number.isFinite(value) ? Math.round(value) : fallback;
  const clamped = clamp(next, min, max);
  if (String(input.value) !== String(clamped)) {
    input.value = clamped;
  }
  return clamped;
}

function readFontSize(input, fallback, min, max) {
  return readNumberInput(input, fallback, min, max);
}

function readSelectValue(select, allowedValues, fallback) {
  if (allowedValues.has(select.value)) {
    return select.value;
  }
  select.value = fallback;
  return fallback;
}

function getLabelDensitySettings(mode) {
  return LABEL_DENSITY_SETTINGS[mode] || LABEL_DENSITY_SETTINGS.balanced;
}

function getTextWeight(key) {
  return state.options.fontBold[key] ? FONT_WEIGHTS.bold : FONT_WEIGHTS.regular;
}

function setStatus(message, type = "") {
  els.status.textContent = message;
  els.status.className = `status ${type}`.trim();
}

async function handleFile(file) {
  try {
    setStatus("קורא קובץ...");
    const workbook = await parseXlsx(file);
    applyWorkbook(workbook);
  } catch (error) {
    console.error(error);
    setStatus(error.message || "לא ניתן לקרוא את הקובץ.", "error");
  }
}

async function ensureFontsReady() {
  if (!fontAvailabilityPromise) {
    fontAvailabilityPromise = resolveFontAvailability();
  }
  return fontAvailabilityPromise;
}

async function resolveFontAvailability() {
  try {
    await loadBundledFonts();
    if (document.fonts) {
      await Promise.all([
        document.fonts.load("400 18px FbPractica"),
        document.fonts.load("800 24px FbPractica"),
        document.fonts.load("400 18px FbPracticaNarrow"),
        document.fonts.load("800 24px FbPracticaNarrow"),
        document.fonts.load("400 18px Heebo"),
        document.fonts.load("700 18px Heebo"),
        document.fonts.load("800 24px Heebo"),
      ]);
      await document.fonts.ready;
    }
  } catch (error) {
    console.warn("Font loading did not complete before rendering.", error);
  }

  const missing = REQUIRED_CANVAS_FONTS
    .filter((font) => !font.aliases.some((family) => isCanvasFontAvailable(family)))
    .map((font) => font.name);

  return { missing };
}

async function loadBundledFonts() {
  if (!("FontFace" in window) || !document.fonts) {
    return;
  }

  await Promise.all(REQUIRED_CANVAS_FONTS.map(loadBundledFont));
}

async function loadBundledFont(font) {
  for (const source of font.sources) {
    const loaded = await tryLoadFontFile(font.family, source);
    if (loaded) {
      return true;
    }
  }
  return false;
}

async function tryLoadFontFile(family, source) {
  try {
    const response = await fetch(source, { cache: "force-cache" });
    if (!response.ok) {
      return false;
    }
    const fontData = await response.arrayBuffer();
    const fontFace = new FontFace(family, fontData);
    await fontFace.load();
    document.fonts.add(fontFace);
    return true;
  } catch (error) {
    return false;
  }
}

function isCanvasFontAvailable(family) {
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return false;
  }

  const sample = "אבגדהוזחטיכלמנסעפצקרשת 0123456789";
  const size = 72;
  const testWidth = measureFontSample(ctx, `"${family}", Arial`, size, sample);
  const fallbackWidth = measureFontSample(ctx, "Arial", size, sample);
  const alternateTestWidth = measureFontSample(ctx, `"${family}", Georgia`, size, sample);
  const alternateFallbackWidth = measureFontSample(ctx, "Georgia", size, sample);

  return (
    Math.abs(testWidth - fallbackWidth) > 0.5 ||
    Math.abs(alternateTestWidth - alternateFallbackWidth) > 0.5
  );
}

function measureFontSample(ctx, fontFamily, size, sample) {
  ctx.font = `400 ${size}px ${fontFamily}`;
  return ctx.measureText(sample).width;
}

function getFontWarningMessage(missing) {
  return `חסרים פונטים לגרף: ${missing.join(", ")}. התקן אותם במחשב או הוסף קבצים מורשים לתיקיית fonts.`;
}

async function downloadPng() {
  const fontStatus = await ensureFontsReady();
  renderChart();
  if (fontStatus.missing.length) {
    setStatus(getFontWarningMessage(fontStatus.missing), "warning");
  }
  const fileDate = state.meta.date ? state.meta.date.replace(/[^\d.-]+/g, "-") : "chart";
  els.chartCanvas.toBlob((blob) => {
    if (!blob) {
      setStatus("לא ניתן לייצא PNG.", "error");
      return;
    }
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `election-polls-${fileDate}.png`;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }, "image/png");
}

els.fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files || [];
  if (file) {
    handleFile(file);
  }
});

els.sampleButton.addEventListener("click", async () => {
  try {
    setStatus("טוען דוגמה...");
    const response = await fetch(`./${encodeURIComponent(SAMPLE_FILE_NAME)}`);
    if (!response.ok) {
      throw new Error("לא ניתן לטעון את קובץ הדוגמה מהתיקייה.");
    }
    const blob = await response.blob();
    await handleFile(new File([blob], SAMPLE_FILE_NAME));
  } catch (error) {
    console.error(error);
    setStatus("בחר את קובץ ה-XLSX ידנית אם האפליקציה נפתחה ישירות מהדיסק.", "error");
  }
});

els.sheetSelect.addEventListener("change", () => {
  if (!state.workbook) {
    return;
  }
  state.activeSheetName = els.sheetSelect.value;
  const extraction = extractSheetData(getActiveSheet(), state.workbook);
  state.parties = extraction.parties;
  state.meta = extraction.meta;
  syncControls();
  renderPartyTable();
  renderChart();
});

for (const input of [
  els.titleInput,
  els.subtitleInput,
  els.dateInput,
  els.widthInput,
  els.heightInput,
  els.titleFontSizeInput,
  els.subtitleFontSizeInput,
  els.numberFontSizeInput,
  els.partyLabelFontSizeInput,
  els.labelColumnGapInput,
]) {
  input.addEventListener("input", renderChart);
}

for (const input of [
  els.sortToggle,
  els.transparentToggle,
  els.labelOverlapSelect,
  els.labelDensitySelect,
  els.titleBoldInput,
  els.subtitleBoldInput,
  els.numberBoldInput,
  els.partyLabelBoldInput,
]) {
  input.addEventListener("change", renderChart);
}

els.partyTable.addEventListener("input", (event) => {
  const input = event.target;
  const row = input.closest("tr");
  if (!row) {
    return;
  }

  const index = Number(row.dataset.index);
  const field = input.dataset.field;
  if (field === "name") {
    state.parties[index].name = input.value;
  }
  if (field === "value") {
    state.parties[index].value = parseNumber(input.value);
  }
  renderChart();
});

els.partyTable.addEventListener("click", (event) => {
  const button = event.target.closest("[data-action='remove']");
  if (!button) {
    return;
  }
  const row = button.closest("tr");
  const index = Number(row.dataset.index);
  state.parties.splice(index, 1);
  renderPartyTable();
  renderChart();
});

els.addRowButton.addEventListener("click", () => {
  state.parties.push({ name: "מפלגה חדשה", value: 0 });
  renderPartyTable();
  renderChart();
});

els.exportButton.addEventListener("click", downloadPng);

syncControls();
renderPartyTable();
renderChart();
ensureFontsReady().then((fontStatus) => {
  renderChart();
  if (fontStatus.missing.length) {
    setStatus(getFontWarningMessage(fontStatus.missing), "warning");
  }
});
