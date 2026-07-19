"use strict";

const SAMPLE_FILE_NAME = "סקר הסקרים - 27.4.xlsx";
const CHART_WIDTH = 550;
const TEXT_COLOR = "#231F20";
const GRID_COLOR = "#dedede";
const AXIS_COLOR = "#231F20";
const TITLE_FONT_FAMILY = '"FbPractica", "Fb Practica", "Heebo", Arial, "Noto Sans Hebrew", "Segoe UI", sans-serif';
const GRAPH_FONT_FAMILY =
  '"FbPracticaNarrow", "Fb Practica Narrow", "FbPractica Narrow", "FbPractica", "Fb Practica", "Heebo", Arial, "Noto Sans Hebrew", "Segoe UI", sans-serif';
const BAR_COLORS = ["#2381be", "#3195d1", "#45ace4", "#7fc9ef", "#b6e1f8"];
const POLL_COUNT = 5;
const BAR_END_LABEL_RESERVE = 45;
const DEFAULT_OPTIONS = {
  sort: true,
  showZeroParties: true,
  axisMax: 0,
  labelWidth: 150,
  barHeight: 8,
  barGap: 18,
  groupGap: 48,
  bottomPadding: 68,
  fontSizes: {
    title: 30,
    subtitle: 28,
    labels: 17,
    values: 20,
    publishers: 17,
    axis: 12,
  },
  fontBold: {
    title: true,
    subtitle: false,
    labels: false,
    values: false,
    publishers: false,
    axis: false,
  },
};

const fallbackState = {
  fileName: SAMPLE_FILE_NAME,
  meta: {
    title: "5 הסקרים האחרונים",
    subtitle: "שפורסמו בכלי התקשורת",
    date: "27.4",
  },
  pollNames: ["חדשות 12", "חדשות 13", "עכשיו 14", "מעריב/וואלה", "i24news"],
  parties: [
    { name: "הליכוד", values: [25, 26, 34, 28, 33] },
    { name: "ביחד", values: [26, 26, 20, 27, 24] },
    { name: "ישר", values: [15, 12, 9, 15, 8] },
    { name: "הרשימה המשותפת", values: [10, 11, 11, 10, 11] },
    { name: 'ש"ס', values: [9, 10, 11, 8, 11] },
    { name: "הדמוקרטים", values: [10, 8, 8, 9, 9] },
    { name: "עוצמה יהודית", values: [9, 10, 7, 8, 6] },
    { name: "ישראל ביתנו", values: [9, 6, 8, 8, 6] },
    { name: "יהדות התורה", values: [7, 7, 8, 7, 8] },
    { name: "הציונות הדתית", values: [0, 4, 4, 0, 4] },
    { name: "המילואימניקים", values: [0, 0, 0, 0, 0] },
    { name: "כחול לבן", values: [0, 0, 0, 0, 0] },
  ],
};

const state = {
  ...structuredClone(fallbackState),
  options: structuredClone(DEFAULT_OPTIONS),
};

const els = {
  advancedToggle: document.querySelector("#advancedToggle"),
  advancedSettings: document.querySelector("#advancedSettings"),
  sampleButton: document.querySelector("#sampleButton"),
  fileInput: document.querySelector("#fileInput"),
  exportButton: document.querySelector("#exportButton"),
  status: document.querySelector("#status"),
  summaryText: document.querySelector("#summaryText"),
  canvas: document.querySelector("#averageCanvas"),
  titleInput: document.querySelector("#titleInput"),
  subtitleInput: document.querySelector("#subtitleInput"),
  dateInput: document.querySelector("#dateInput"),
  sortToggle: document.querySelector("#sortToggle"),
  showZeroToggle: document.querySelector("#showZeroToggle"),
  axisMaxInput: document.querySelector("#axisMaxInput"),
  labelWidthInput: document.querySelector("#labelWidthInput"),
  barHeightInput: document.querySelector("#barHeightInput"),
  barGapInput: document.querySelector("#barGapInput"),
  groupGapInput: document.querySelector("#groupGapInput"),
  bottomPaddingInput: document.querySelector("#bottomPaddingInput"),
  titleFontSizeInput: document.querySelector("#titleFontSizeInput"),
  subtitleFontSizeInput: document.querySelector("#subtitleFontSizeInput"),
  partyLabelFontSizeInput: document.querySelector("#partyLabelFontSizeInput"),
  valueFontSizeInput: document.querySelector("#valueFontSizeInput"),
  publisherFontSizeInput: document.querySelector("#publisherFontSizeInput"),
  titleBoldInput: document.querySelector("#titleBoldInput"),
  subtitleBoldInput: document.querySelector("#subtitleBoldInput"),
  partyLabelBoldInput: document.querySelector("#partyLabelBoldInput"),
  valueBoldInput: document.querySelector("#valueBoldInput"),
  publisherBoldInput: document.querySelector("#publisherBoldInput"),
  axisBoldInput: document.querySelector("#axisBoldInput"),
  pollNameInputs: Array.from(document.querySelectorAll("[data-poll-name]")),
  pollHeaders: Array.from({ length: POLL_COUNT }, (_, index) => document.querySelector(`#pollHeader${index}`)),
  partyTable: document.querySelector("#partyTable"),
  addRowButton: document.querySelector("#addRowButton"),
};

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
    return new Uint8Array(inflated);
  }
}

async function inflateRaw(bytes) {
  if (!("DecompressionStream" in window)) {
    throw new Error("הדפדפן הזה לא תומך בקריאת XLSX מקומית. מומלץ לפתוח ב-Chrome או Edge עדכני.");
  }

  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Response(stream).arrayBuffer();
}

function parseXml(text) {
  const xml = new DOMParser().parseFromString(text, "application/xml");
  if (xml.querySelector("parsererror")) {
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
  const sheetNodes = elements(workbookXml, "sheet");
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
  return elements(xml, "si").map((node) => node.textContent || "");
}

function readRelationships(xml) {
  const rels = new Map();
  for (const node of elements(xml, "Relationship")) {
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

  for (const rowNode of elements(xml, "row")) {
    const rowIndex = Number(rowNode.getAttribute("r")) || rows.length + 1;
    const row = [];

    for (const cellNode of elements(rowNode, "c")) {
      const ref = cellNode.getAttribute("r") || "";
      const columnLetters = ref.match(/[A-Z]+/i)?.[0] || "A";
      row[lettersToColumnIndex(columnLetters)] = readCellValue(cellNode, sharedStrings);
    }

    rows[rowIndex - 1] = row;
  }

  return Array.from({ length: rows.length }, (_, index) => rows[index] || []);
}

function readCellValue(cellNode, sharedStrings) {
  const type = cellNode.getAttribute("t");
  const valueNode = elements(cellNode, "v")[0];

  if (type === "s") {
    const index = Number(valueNode?.textContent || 0);
    return sharedStrings[index] || "";
  }

  if (type === "inlineStr") {
    return elements(cellNode, "is")[0]?.textContent || "";
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

function elements(node, localName) {
  const direct = Array.from(node.getElementsByTagName(localName));
  if (direct.length) {
    return direct;
  }
  return Array.from(node.getElementsByTagNameNS("*", localName));
}

function extractAverageFiveData(workbook) {
  const candidates = workbook.sheets
    .map((sheet) => {
      const header = findHeader(sheet.rows);
      const numericColumns =
        header.index >= 0 ? getNumericColumns(sheet.rows, header.index, header.nameColumn) : [];
      return { sheet, header, numericColumns };
    })
    .filter((candidate) => candidate.header.index >= 0 && candidate.numericColumns.length >= 5)
    .sort((a, b) => b.numericColumns.length - a.numericColumns.length);
  const selected =
    candidates.find((candidate) => /כל הסקרים|all polls|polls/i.test(candidate.sheet.name)) || candidates[0];

  if (!selected) {
    throw new Error("לא נמצא גליון עם לפחות 5 עמודות סקרים.");
  }

  const { sheet, header, numericColumns } = selected;
  const selectedColumns = numericColumns.slice(-5);
  const headerRow = sheet.rows[header.index] || [];
  const pollNames = selectedColumns.map((column) =>
    normalizePublisherName(cleanText(headerRow[column]) || `סקר ${column + 1}`),
  );
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

    const values = selectedColumns.map((column) => Math.max(0, Math.round(parseNumber(row[column]) || 0)));
    parties.push({ name, values });
  }

  if (!parties.length) {
    throw new Error("לא נמצאו נתוני מפלגות להצגה.");
  }

  const meta = extractMetadata(sheet.rows, header.index);
  return {
    fileName: workbook.fileName,
    meta: {
      title: "5 הסקרים האחרונים",
      subtitle: meta.subtitle || fallbackState.meta.subtitle,
      date: meta.date || fallbackState.meta.date,
    },
    pollNames,
    parties,
  };
}

function findHeader(rows) {
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const nameColumn = row.findIndex((cell) => /מפלגה|party/i.test(cleanText(cell)));
    if (nameColumn >= 0) {
      return { index: rowIndex, nameColumn };
    }
  }

  return { index: -1, nameColumn: 0 };
}

function getNumericColumns(rows, headerIndex, nameColumn) {
  const columns = [];
  const maxColumns = Math.max(...rows.map((row) => row.length), 0);

  for (let column = 0; column < maxColumns; column += 1) {
    if (column === nameColumn) {
      continue;
    }

    let numericCount = 0;
    for (let rowIndex = headerIndex + 1; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const name = cleanText(row[nameColumn]);
      if (!name || isSummaryRow(name)) {
        continue;
      }
      if (Number.isFinite(parseNumber(row[column]))) {
        numericCount += 1;
      }
    }

    if (numericCount >= 2) {
      columns.push(column);
    }
  }

  return columns;
}

function extractMetadata(rows, headerIndex) {
  const firstRows = rows.slice(0, Math.max(headerIndex, 4));
  const firstColumnValues = firstRows.map((row) => cleanText(firstNonEmpty(row))).filter(Boolean);
  const [, rawSubtitle = "", rawDate = ""] = firstColumnValues;
  return {
    subtitle: rawSubtitle,
    date: rawDate ? formatValue(rawDate) : "",
  };
}

function firstNonEmpty(row) {
  return (row || []).find((cell) => cleanText(cell));
}

function average(values) {
  return values.reduce((sum, value) => sum + value, 0) / Math.max(1, values.length);
}

function normalizePublisherName(name) {
  const normalized = cleanText(name);
  if (/^ערוץ\s*12$/i.test(normalized)) {
    return "חדשות 12";
  }
  if (/^ערוץ\s*13$/i.test(normalized)) {
    return "חדשות 13";
  }
  if (/^ערוץ\s*14$/i.test(normalized)) {
    return "עכשיו 14";
  }
  return normalized;
}

function syncControls() {
  els.titleInput.value = state.meta.title;
  els.subtitleInput.value = state.meta.subtitle;
  els.dateInput.value = state.meta.date;
  els.sortToggle.checked = state.options.sort;
  els.showZeroToggle.checked = state.options.showZeroParties;
  els.axisMaxInput.value = state.options.axisMax;
  els.labelWidthInput.value = state.options.labelWidth;
  els.barHeightInput.value = state.options.barHeight;
  els.barGapInput.value = state.options.barGap;
  els.groupGapInput.value = state.options.groupGap;
  els.bottomPaddingInput.value = state.options.bottomPadding;
  els.titleFontSizeInput.value = state.options.fontSizes.title;
  els.subtitleFontSizeInput.value = state.options.fontSizes.subtitle;
  els.partyLabelFontSizeInput.value = state.options.fontSizes.labels;
  els.valueFontSizeInput.value = state.options.fontSizes.values;
  els.publisherFontSizeInput.value = state.options.fontSizes.publishers;
  els.titleBoldInput.checked = state.options.fontBold.title;
  els.subtitleBoldInput.checked = state.options.fontBold.subtitle;
  els.partyLabelBoldInput.checked = state.options.fontBold.labels;
  els.valueBoldInput.checked = state.options.fontBold.values;
  els.publisherBoldInput.checked = state.options.fontBold.publishers;
  els.axisBoldInput.checked = state.options.fontBold.axis;
  syncPollNameInputs();
}

function syncPollNameInputs() {
  els.pollNameInputs.forEach((input) => {
    const index = Number(input.dataset.pollName);
    input.value = state.pollNames[index] || `סקר ${index + 1}`;
  });
  updatePollHeaders();
}

function readControls() {
  state.meta.title = els.titleInput.value.trim();
  state.meta.subtitle = els.subtitleInput.value.trim();
  state.meta.date = els.dateInput.value.trim();
  state.options.sort = els.sortToggle.checked;
  state.options.showZeroParties = els.showZeroToggle.checked;
  state.options.axisMax = readNumberInput(els.axisMaxInput, DEFAULT_OPTIONS.axisMax, 0, 120);
  state.options.labelWidth = readNumberInput(els.labelWidthInput, DEFAULT_OPTIONS.labelWidth, 70, 190);
  state.options.barHeight = readNumberInput(els.barHeightInput, DEFAULT_OPTIONS.barHeight, 4, 20);
  state.options.barGap = readNumberInput(els.barGapInput, DEFAULT_OPTIONS.barGap, 0, 20);
  state.options.groupGap = readNumberInput(els.groupGapInput, DEFAULT_OPTIONS.groupGap, 6, 70);
  state.options.bottomPadding = readNumberInput(els.bottomPaddingInput, DEFAULT_OPTIONS.bottomPadding, 64, 120);
  state.options.fontSizes = {
    title: readNumberInput(els.titleFontSizeInput, DEFAULT_OPTIONS.fontSizes.title, 12, 44),
    subtitle: readNumberInput(els.subtitleFontSizeInput, DEFAULT_OPTIONS.fontSizes.subtitle, 10, 34),
    labels: readNumberInput(els.partyLabelFontSizeInput, DEFAULT_OPTIONS.fontSizes.labels, 8, 22),
    values: readNumberInput(els.valueFontSizeInput, DEFAULT_OPTIONS.fontSizes.values, 8, 22),
    publishers: readNumberInput(
      els.publisherFontSizeInput,
      DEFAULT_OPTIONS.fontSizes.publishers,
      8,
      22,
    ),
    axis: DEFAULT_OPTIONS.fontSizes.axis,
  };
  state.options.fontBold = {
    title: els.titleBoldInput.checked,
    subtitle: els.subtitleBoldInput.checked,
    labels: els.partyLabelBoldInput.checked,
    values: els.valueBoldInput.checked,
    publishers: els.publisherBoldInput.checked,
    axis: els.axisBoldInput.checked,
  };
  state.pollNames = els.pollNameInputs.map((input, index) => cleanText(input.value) || `סקר ${index + 1}`);
  updatePollHeaders();
}

function getRenderableParties() {
  const parties = state.parties
    .map((party) => ({
      name: cleanText(party.name),
      values: Array.from({ length: POLL_COUNT }, (_, index) =>
        Math.max(0, Math.round(parseNumber(party.values[index]) || 0)),
      ),
    }))
    .filter((party) => party.name && (state.options.showZeroParties || party.values.some((value) => value > 0)));

  if (state.options.sort) {
    return parties.slice().sort((a, b) => average(b.values) - average(a.values));
  }
  return parties;
}

function renderPartyTable() {
  updatePollHeaders();
  els.partyTable.innerHTML = "";
  const fragment = document.createDocumentFragment();

  state.parties.forEach((party, index) => {
    const row = document.createElement("tr");
    row.dataset.index = String(index);

    const nameCell = document.createElement("td");
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.value = party.name;
    nameInput.dataset.field = "name";
    nameCell.append(nameInput);
    row.append(nameCell);

    for (let pollIndex = 0; pollIndex < POLL_COUNT; pollIndex += 1) {
      const valueCell = document.createElement("td");
      valueCell.className = "poll-value";
      const valueInput = document.createElement("input");
      valueInput.type = "number";
      valueInput.min = "0";
      valueInput.max = "120";
      valueInput.step = "1";
      valueInput.value = formatEditableNumber(party.values[pollIndex] ?? 0);
      valueInput.dataset.field = "value";
      valueInput.dataset.pollIndex = String(pollIndex);
      valueCell.append(valueInput);
      row.append(valueCell);
    }

    const actionCell = document.createElement("td");
    const removeButton = document.createElement("button");
    removeButton.className = "remove-row";
    removeButton.type = "button";
    removeButton.dataset.action = "remove";
    removeButton.title = "מחיקה";
    removeButton.setAttribute("aria-label", "מחיקת שורה");
    removeButton.textContent = "×";
    actionCell.append(removeButton);
    row.append(actionCell);

    fragment.append(row);
  });

  els.partyTable.append(fragment);
}

function updatePollHeaders() {
  els.pollHeaders.forEach((header, index) => {
    header.textContent = state.pollNames[index] || String(index + 1);
  });
}

function setAdvancedSettingsOpen(isOpen) {
  els.advancedSettings.hidden = !isOpen;
  els.advancedToggle.setAttribute("aria-expanded", String(isOpen));
  els.advancedToggle.textContent = isOpen ? "סגור הגדרות מתקדמות" : "הגדרות מתקדמות";
  document.body.classList.toggle("advanced-settings-open", isOpen);
}

function renderChart() {
  readControls();
  const data = getRenderableParties();
  const maxValue = Math.max(1, ...data.flatMap((party) => party.values));
  const autoAxisMax = Math.max(10, Math.ceil((maxValue + 2) / 5) * 5);
  const axisMax = state.options.axisMax ? Math.max(maxValue, state.options.axisMax) : autoAxisMax;
  const topPadding = Math.max(
    86,
    state.options.fontSizes.title + state.options.fontSizes.subtitle + 52,
  );
  const labelWidth = state.options.labelWidth;
  const rightPadding = 10;
  const chartLeft = labelWidth;
  const chartRight = CHART_WIDTH - rightPadding - BAR_END_LABEL_RESERVE;
  const chartWidth = chartRight - chartLeft;
  const barHeight = state.options.barHeight;
  const barGap = state.options.barGap;
  const groupGap = state.options.groupGap;
  const groupHeight = POLL_COUNT * barHeight + (POLL_COUNT - 1) * barGap;
  const bottomPadding = state.options.bottomPadding;
  const chartHeight = data.length * groupHeight + Math.max(0, data.length - 1) * groupGap;
  const height = Math.ceil(topPadding + chartHeight + bottomPadding);

  const canvas = els.canvas;
  if (canvas.width !== CHART_WIDTH) {
    canvas.width = CHART_WIDTH;
  }
  if (canvas.height !== height) {
    canvas.height = height;
  }

  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, CHART_WIDTH, height);

  ctx.direction = "rtl";
  ctx.textAlign = "center";
  ctx.fillStyle = TEXT_COLOR;
  ctx.textBaseline = "alphabetic";
  ctx.font = `${getFontWeight("title")} ${state.options.fontSizes.title}px ${TITLE_FONT_FAMILY}`;
  ctx.fillText(state.meta.title, CHART_WIDTH / 2, 30);
  ctx.font = `${getFontWeight("subtitle")} ${state.options.fontSizes.subtitle}px ${TITLE_FONT_FAMILY}`;
  const subtitle = state.meta.date ? `${state.meta.subtitle} (${state.meta.date})` : state.meta.subtitle;
  ctx.fillText(subtitle, CHART_WIDTH / 2, 34 + state.options.fontSizes.subtitle + 8);

  const plotTop = topPadding;
  const plotBottom = topPadding + chartHeight;
  drawGrid(ctx, chartLeft, chartRight, plotTop, plotBottom, axisMax);

  data.forEach((party, partyIndex) => {
    const groupTop = plotTop + partyIndex * (groupHeight + groupGap);
    const groupCenter = groupTop + groupHeight / 2;

    ctx.fillStyle = TEXT_COLOR;
    ctx.font = `${getFontWeight("labels")} ${state.options.fontSizes.labels}px ${GRAPH_FONT_FAMILY}`;
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    drawFittedPartyName(
      ctx,
      party.name,
      chartLeft - 16,
      groupCenter,
      chartLeft - 22,
      state.options.fontSizes.labels,
      getFontWeight("labels"),
    );

    party.values.forEach((value, pollIndex) => {
      const y = groupTop + pollIndex * (barHeight + barGap);
      const width = (chartWidth * value) / axisMax;
      ctx.fillStyle = BAR_COLORS[pollIndex % BAR_COLORS.length];
      roundedRect(ctx, chartLeft, y, width, barHeight, 3);
      ctx.fill();

      drawBarEndLabel(
        ctx,
        value,
        state.pollNames[pollIndex] || `סקר ${pollIndex + 1}`,
        chartLeft + width + 7,
        y + barHeight / 2,
        CHART_WIDTH - rightPadding,
      );
    });
  });

  drawAxisLabels(
    ctx,
    chartLeft,
    chartRight,
    plotBottom,
    axisMax,
    state.options.fontSizes.axis,
    getFontWeight("axis"),
  );
  updateSummary(axisMax);
}

function drawGrid(ctx, chartLeft, chartRight, plotTop, plotBottom, axisMax) {
  ctx.save();
  ctx.strokeStyle = GRID_COLOR;
  ctx.lineWidth = 1;

  for (let value = 5; value <= axisMax; value += 5) {
    const x = chartLeft + ((chartRight - chartLeft) * value) / axisMax;
    ctx.beginPath();
    ctx.moveTo(x, plotTop);
    ctx.lineTo(x, plotBottom);
    ctx.stroke();
  }

  ctx.strokeStyle = AXIS_COLOR;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(chartLeft, plotTop);
  ctx.lineTo(chartLeft, plotBottom);
  ctx.stroke();
  ctx.restore();
}

function drawAxisLabels(ctx, chartLeft, chartRight, plotBottom, axisMax, fontSize, weight) {
  ctx.save();
  ctx.direction = "ltr";
  ctx.fillStyle = TEXT_COLOR;
  ctx.font = `${weight} ${fontSize}px ${GRAPH_FONT_FAMILY}`;
  ctx.textAlign = "center";
  ctx.textBaseline = "top";

  for (let value = 0; value <= axisMax; value += 5) {
    const x = chartLeft + ((chartRight - chartLeft) * value) / axisMax;
    ctx.fillText(String(value), x, plotBottom + 16);
  }

  ctx.direction = "rtl";
  ctx.fillText("מנדטים", (chartLeft + chartRight) / 2, plotBottom + 38);
  ctx.restore();
}

function roundedRect(ctx, x, y, width, height, radius) {
  const safeWidth = Math.max(0, width);
  const safeRadius = Math.min(radius, safeWidth / 2, height / 2);
  ctx.beginPath();
  ctx.moveTo(x + safeRadius, y);
  ctx.lineTo(x + safeWidth - safeRadius, y);
  ctx.quadraticCurveTo(x + safeWidth, y, x + safeWidth, y + safeRadius);
  ctx.lineTo(x + safeWidth, y + height - safeRadius);
  ctx.quadraticCurveTo(x + safeWidth, y + height, x + safeWidth - safeRadius, y + height);
  ctx.lineTo(x + safeRadius, y + height);
  ctx.quadraticCurveTo(x, y + height, x, y + height - safeRadius);
  ctx.lineTo(x, y + safeRadius);
  ctx.quadraticCurveTo(x, y, x + safeRadius, y);
  ctx.closePath();
}

function drawFittedPartyName(ctx, text, x, y, maxWidth, startSize, weight) {
  let size = startSize;
  while (size > 8) {
    ctx.font = `${weight} ${size}px ${GRAPH_FONT_FAMILY}`;
    if (ctx.measureText(text).width <= maxWidth) {
      break;
    }
    size -= 1;
  }
  ctx.fillText(text, x, y);
}

function drawBarEndLabel(ctx, value, publisher, x, y, maxX) {
  ctx.save();
  ctx.textAlign = "left";
  ctx.textBaseline = "middle";
  ctx.direction = "ltr";
  ctx.fillStyle = "#000000";
  ctx.font = `${getFontWeight("values")} ${state.options.fontSizes.values}px ${GRAPH_FONT_FAMILY}`;

  const valueText = String(value);
  ctx.fillText(valueText, x, y);
  const publisherX = x + ctx.measureText(valueText).width + 6;
  const publisherText = formatPublisherLabel(publisher);
  const maxPublisherWidth = Math.max(0, maxX - publisherX);
  let publisherSize = state.options.fontSizes.publishers;

  ctx.direction = "rtl";
  ctx.fillStyle = TEXT_COLOR;
  while (publisherSize > 8) {
    ctx.font = `${getFontWeight("publishers")} ${publisherSize}px ${GRAPH_FONT_FAMILY}`;
    if (ctx.measureText(publisherText).width <= maxPublisherWidth) {
      break;
    }
    publisherSize -= 1;
  }
  ctx.fillText(publisherText, publisherX, y, maxPublisherWidth);
  ctx.restore();
}

function formatPublisherLabel(publisher) {
  const text = cleanText(publisher).replace(/^\((.*)\)$/, "$1").trim();
  return `(${text})`;
}

function getFontWeight(key) {
  return state.options.fontBold[key] ? 800 : 400;
}

function updateSummary(axisMax) {
  els.summaryText.textContent = `${state.parties.length} מפלגות, ${state.pollNames.length} סקרים, ציר עד ${axisMax} מנדטים`;
}

async function handleFile(file) {
  try {
    setStatus("קורא קובץ...");
    const workbook = await parseXlsx(file);
    applyChartData(extractAverageFiveData(workbook));
    setStatus(`נטען: ${file.name}`);
  } catch (error) {
    console.error(error);
    setStatus(error.message || "לא ניתן לקרוא את הקובץ.", "error");
  }
}

function applyChartData(nextData) {
  state.fileName = nextData.fileName;
  state.meta = nextData.meta;
  state.pollNames = nextData.pollNames;
  state.parties = nextData.parties;
  syncControls();
  renderPartyTable();
  renderChart();
}

async function loadSampleWorkbook() {
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
    setStatus("הדוגמה המובנית מוצגת. כדי לקרוא XLSX יש לפתוח דרך שרת מקומי או לבחור קובץ ידנית.", "warning");
  }
}

function downloadPng() {
  renderChart();
  const fileDate = state.meta.date ? state.meta.date.replace(/[^\d.-]+/g, "-") : "chart";
  els.canvas.toBlob((blob) => {
    if (!blob) {
      setStatus("לא ניתן לייצא PNG.", "error");
      return;
    }
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `average-5-polls-${fileDate}.png`;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }, "image/png");
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

function readNumberInput(input, fallback, min, max) {
  const value = Number(input.value);
  const next = Number.isFinite(value) ? Math.round(value) : fallback;
  const clamped = Math.min(max, Math.max(min, next));
  if (String(input.value) !== String(clamped)) {
    input.value = clamped;
  }
  return clamped;
}

function cleanText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function formatEditableNumber(value) {
  const numeric = parseNumber(value);
  if (!Number.isFinite(numeric)) {
    return "";
  }
  return Number.isInteger(numeric) ? String(numeric) : String(Number(numeric.toFixed(1)));
}

function formatValue(value) {
  if (typeof value === "number") {
    return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(3))).replace(/\.0$/, "");
  }
  return cleanText(value);
}

function isBlankRow(row) {
  return !(row || []).some((cell) => cleanText(cell));
}

function isSummaryRow(name) {
  return /קואליציה|אופוזיציה|סה.?כ|total|coalition|opposition/i.test(name);
}

function setStatus(message, type = "") {
  els.status.textContent = message;
  els.status.className = `status ${type}`.trim();
}

els.advancedToggle.addEventListener("click", () => {
  setAdvancedSettingsOpen(els.advancedSettings.hidden);
});

els.sampleButton.addEventListener("click", loadSampleWorkbook);

els.fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files || [];
  if (file) {
    handleFile(file);
  }
});

els.exportButton.addEventListener("click", downloadPng);

for (const input of [
  els.titleInput,
  els.subtitleInput,
  els.dateInput,
  els.axisMaxInput,
  els.labelWidthInput,
  els.barHeightInput,
  els.barGapInput,
  els.groupGapInput,
  els.bottomPaddingInput,
  els.titleFontSizeInput,
  els.subtitleFontSizeInput,
  els.partyLabelFontSizeInput,
  els.valueFontSizeInput,
  els.publisherFontSizeInput,
]) {
  input.addEventListener("input", renderChart);
}

for (const input of [
  els.sortToggle,
  els.showZeroToggle,
  els.titleBoldInput,
  els.subtitleBoldInput,
  els.partyLabelBoldInput,
  els.valueBoldInput,
  els.publisherBoldInput,
  els.axisBoldInput,
]) {
  input.addEventListener("change", renderChart);
}

for (const input of els.pollNameInputs) {
  input.addEventListener("input", () => {
    renderChart();
  });
}

els.partyTable.addEventListener("input", (event) => {
  const input = event.target;
  const row = input.closest("tr");
  if (!row) {
    return;
  }

  const index = Number(row.dataset.index);
  const field = input.dataset.field;
  if (!state.parties[index]) {
    return;
  }

  if (field === "name") {
    state.parties[index].name = input.value;
  }
  if (field === "value") {
    const pollIndex = Number(input.dataset.pollIndex);
    state.parties[index].values[pollIndex] = Math.max(0, Math.round(parseNumber(input.value) || 0));
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
  state.parties.push({
    name: "מפלגה חדשה",
    values: Array.from({ length: POLL_COUNT }, () => 0),
  });
  renderPartyTable();
  renderChart();
});

setAdvancedSettingsOpen(false);
syncControls();
renderPartyTable();
renderChart();
loadSampleWorkbook();
