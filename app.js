/**
 * Rapport de correctifs Import — CSV/XLS(X)
 * - Causes probables : lecture via FileReader non centralisée et dépendance PapaParse bloquant la lecture hors CDN.
 * - Changements effectués : importAnyFile mutualise FileReader + SheetJS, parseur CSV maison (BOM, délimiteur auto) et spinner avec erreurs détaillées.
 * - Limites : encodages non UTF-8 non détectés automatiquement ; seules les premières feuilles XLS(X) sont analysées.
 * - Tests manuels : CSV (, ; \t |) avec BOM, XLSX multi-onglets, drag & drop, reset, analyse & comparaison avec filtres + export.
 */
/**
 * Rapport de modifications — Comparaison avec cases d'en-tête
 * - Bugs corrigés : export limité aux résultats filtrés, réinitialisation sûre quand aucune colonne n'est cochée, nettoyage des états comparaison/ref.
 * - Décisions techniques : modèle de colonnes unifié (`tableColumns` + clés ref./cmp.), colonne "Mots-clés trouvés" calculée depuis les opérandes et surbrillance <mark>.
 * - Impact accessibilité : fieldset résumé, cases à cocher labellisées dans les th, aria-live pour limites de recherche et erreurs.
 * - Performance : caches de cellules restreints aux colonnes recherchables, debounce 300 ms conservé, worker inchangé.
 * - Tests : checklist manuelle actualisée (comparaison, export, 10k lignes) + suite automatisée mise à jour (colonnes ref./cmp.).
 */

const doc = typeof document !== "undefined" ? document : null;
const dropZone = doc ? doc.getElementById("drop-zone") : null;
const fileInput = doc
  ? doc.getElementById("fileSingle") || doc.getElementById("file-input")
  : null;
const modeRadios = doc
  ? Array.from(doc.querySelectorAll('input[name="analysis-mode"]'))
  : [];
const singleUploadContainer = doc ? doc.getElementById("single-upload") : null;
const compareUploadContainer = doc ? doc.getElementById("compare-upload") : null;
const referenceInput = doc
  ? doc.getElementById("fileRef") || doc.getElementById("reference-input")
  : null;
const comparisonInput = doc
  ? doc.getElementById("fileCmp") || doc.getElementById("comparison-input")
  : null;
const referenceName = doc ? doc.getElementById("reference-name") : { textContent: "" };
const comparisonName = doc
  ? doc.getElementById("comparison-name")
  : { textContent: "" };
const keywordsSummary = doc ? doc.getElementById("keywords-summary") : { textContent: "" };
const progressBar = doc ? doc.getElementById("progress-bar") : null;
const progressLabel = doc ? doc.getElementById("progress-label") : null;
const loadingSpinner = doc ? doc.getElementById("loading-spinner") : null;
const loadingSpinnerLabel = doc
  ? doc.getElementById("loading-spinner-label")
  : null;
const errorMessage = doc ? doc.getElementById("error-message") : null;
if (errorMessage) {
  errorMessage.hidden = true;
}
const controlsSection = doc ? doc.getElementById("controls") : { hidden: true };
const resultsSection = doc ? doc.getElementById("results") : { hidden: true };
const searchInput = doc ? doc.getElementById("search-input") : { value: "" };
const searchButton = doc ? doc.getElementById("search-button") : null;
const resetButton = doc ? doc.getElementById("reset-button") : null;
const caseSensitiveToggle = doc
  ? doc.getElementById("case-sensitive")
  : { checked: false };
const exactMatchToggle = doc ? doc.getElementById("exact-match") : { checked: false };
const dataTable = doc ? doc.getElementById("data-table") : null;
const resultStats = doc ? doc.getElementById("result-stats") : { textContent: "" };
const pagination = doc ? doc.getElementById("pagination") : { hidden: true };
const pageInfo = doc ? doc.getElementById("page-info") : { textContent: "" };
const prevPageBtn = doc ? doc.getElementById("prev-page") : { disabled: true };
const nextPageBtn = doc ? doc.getElementById("next-page") : { disabled: true };
const copyButton = doc ? doc.getElementById("copy-button") : null;
const exportCsvButton = doc ? doc.getElementById("export-csv-button") : null;
const exportXlsxButton = doc ? doc.getElementById("export-xlsx-button") : null;
const columnFilterFieldset = doc ? doc.getElementById("column-filter") : null;
const columnFilterLabel = doc
  ? doc.getElementById("column-filter-label")
  : { textContent: "" };
const columnFilterLive = doc ? doc.getElementById("column-filter-live") : null;
const liveFeedback = doc ? doc.getElementById("live-feedback") : null;

const PAGE_SIZE = 100;
const SEARCH_DEBOUNCE = 300;
const RENDER_BATCH_SIZE = 200;
const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50MB

const activeReaders = new Set();

const STORAGE_KEYS = {
  selectedRef: "bp8.search.selectedColumns.ref",
  selectedCmp: "bp8.search.selectedColumns.cmp",
};

const state = {
  search: {
    query: "",
    caseSensitive: false,
    exactMatch: false,
    selectedKeys: null,
  },
};

function debounce(fn, delay) {
  let timerId = null;
  function debounced(...args) {
    if (timerId) {
      clearTimeout(timerId);
    }
    timerId = setTimeout(() => {
      timerId = null;
      fn.apply(this, args);
    }, delay);
  }
  debounced.cancel = () => {
    if (timerId) {
      clearTimeout(timerId);
      timerId = null;
    }
  };
  return debounced;
}

const scheduleSearch = debounce(() => performSearch({ source: "debounce" }), SEARCH_DEBOUNCE);

let headers = [];
let rawRows = [];
let filteredRows = [];
let rowTextCache = [];
let lowerRowTextCache = [];
let tableColumns = [];
let columnKeyToIndex = new Map();
let filteredRowHighlights = [];
let matchesColumnIndex = -1;
let currentPage = 1;
let currentFileName = "";
let currentMode = "single";
let referenceKeywords = [];
let referenceHeaders = [];
let referenceRows = [];
let comparisonHeaders = [];
let comparisonRows = [];
let referenceFileName = "";
let comparisonFileName = "";

function resetDataset(options = {}) {
  if (!options.keepSpinner) {
    setLoading(false, "");
  }
  headers = [];
  rawRows = [];
  filteredRows = [];
  rowTextCache = [];
  lowerRowTextCache = [];
  tableColumns = [];
  columnKeyToIndex = new Map();
  filteredRowHighlights = [];
  matchesColumnIndex = -1;
  state.search.selectedKeys = null;
  currentPage = 1;
  currentFileName = "";
  updateProgress(0, "");
  clearError();
  controlsSection.hidden = true;
  resultsSection.hidden = true;
  pagination.hidden = true;
  if (dataTable) {
    dataTable.innerHTML = "";
  }
  resultStats.textContent = "";
  setColumnFilterInteractivity(true);
  updateSearchSummary();
  announceColumnSelection();
}

function announceStatus(message) {
  if (liveFeedback) {
    liveFeedback.textContent = message || "";
  }
}

function showError(message, detail = "") {
  const combined = detail ? `${message} ${detail}` : message;
  if (errorMessage) {
    errorMessage.textContent = combined;
    errorMessage.hidden = false;
  }
  announceStatus(combined);
}

function clearError() {
  if (errorMessage) {
    errorMessage.textContent = "";
    errorMessage.hidden = true;
  }
  announceStatus("");
}

function updateProgress(percent, label) {
  if (progressBar && progressBar.style) {
    progressBar.style.width = `${percent}%`;
  }
  if (progressLabel) {
    progressLabel.textContent = label;
  }
}

function setLoading(isLoading, label) {
  if (!loadingSpinner) {
    return;
  }
  if (isLoading) {
    const spinnerLabel = typeof label === "string" && label.trim() ? label : "Import en cours…";
    loadingSpinner.hidden = false;
    loadingSpinner.setAttribute("aria-busy", "true");
    if (loadingSpinnerLabel) {
      loadingSpinnerLabel.textContent = spinnerLabel;
    }
    updateProgress(0, spinnerLabel);
  } else {
    loadingSpinner.hidden = true;
    loadingSpinner.removeAttribute("aria-busy");
    if (label !== undefined) {
      if (label) {
        updateProgress(100, label);
      } else {
        updateProgress(0, "");
      }
    }
  }
}

function abortActiveReaders() {
  activeReaders.forEach((reader) => {
    try {
      reader.abort();
    } catch (error) {
      console.warn("Lecture interrompue", error);
    }
  });
  activeReaders.clear();
}

function formatBytes(bytes) {
  const units = ["octets", "Ko", "Mo", "Go"];
  let size = bytes;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }
  return `${size.toFixed(unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

function sanitizeColumnLabel(header, index) {
  if (header === undefined || header === null || header === "") {
    return `Colonne ${index + 1}`;
  }
  return String(header);
}

function buildColumns(headers, rows, { prefix, origin }) {
  const firstRow = Array.isArray(rows) && rows.length ? rows[0] : [];
  const columnCount = Math.max(headers.length, firstRow.length);
  if (!columnCount) {
    return [];
  }
  const columns = [];
  const seenLabels = new Map();
  for (let index = 0; index < columnCount; index += 1) {
    const baseLabel = sanitizeColumnLabel(headers[index], index);
    const prefixedLabel = prefix ? `${prefix}.${baseLabel}` : baseLabel;
    const occurrences = seenLabels.get(prefixedLabel) || 0;
    seenLabels.set(prefixedLabel, occurrences + 1);
    const label = occurrences === 0 ? prefixedLabel : `${prefixedLabel} (${occurrences + 1})`;
    columns.push({ key: label, label, origin, index });
  }
  return columns;
}

function getColumnsFor(fileKind, rows, headers = []) {
  const prefix = fileKind === "cmp" ? "cmp" : "ref";
  const sanitizedHeaders = Array.isArray(headers)
    ? headers.map((header, index) => sanitizeColumnLabel(header, index))
    : [];
  return buildColumns(sanitizedHeaders, rows, { prefix, origin: fileKind });
}

function getSingleFileColumns(headers, rows) {
  const sanitizedHeaders = Array.isArray(headers)
    ? headers.map((header, index) => sanitizeColumnLabel(header, index))
    : [];
  return buildColumns(sanitizedHeaders, rows, { prefix: "col", origin: "single" }).map(
    (column) => ({
      ...column,
      label: sanitizedHeaders[column.index] || sanitizeColumnLabel("", column.index),
    })
  );
}

function getAvailableColumns(rows = rawRows) {
  const firstRow = Array.isArray(rows) && rows.length ? rows[0] : [];
  const totalColumns = Math.max(headers.length, firstRow.length);
  if (!totalColumns) {
    return [];
  }
  const detected = [];
  for (let index = 0; index < totalColumns; index += 1) {
    detected.push({ key: String(index), label: sanitizeColumnLabel(headers[index], index) });
  }
  return detected;
}

function loadSelectedKeysForComparison() {
  if (typeof window === "undefined") {
    return null;
  }
  try {
    const storedRef = window.localStorage.getItem(STORAGE_KEYS.selectedRef);
    const storedCmp = window.localStorage.getItem(STORAGE_KEYS.selectedCmp);
    const parsed = [];
    if (storedRef) {
      const values = JSON.parse(storedRef);
      if (Array.isArray(values)) {
        parsed.push(...values.map((value) => String(value)));
      }
    }
    if (storedCmp) {
      const values = JSON.parse(storedCmp);
      if (Array.isArray(values)) {
        parsed.push(...values.map((value) => String(value)));
      }
    }
    return parsed.length ? new Set(parsed) : null;
  } catch (error) {
    console.warn("Impossible de charger les colonnes persistées", error);
    return null;
  }
}

function saveSelectedKeysForComparison(selection) {
  if (typeof window === "undefined") {
    return;
  }
  try {
    if (!selection || !selection.size) {
      window.localStorage.removeItem(STORAGE_KEYS.selectedRef);
      window.localStorage.removeItem(STORAGE_KEYS.selectedCmp);
      return;
    }
    const refKeys = [];
    const cmpKeys = [];
    selection.forEach((key) => {
      if (typeof key !== "string") {
        return;
      }
      if (key.startsWith("ref.")) {
        refKeys.push(key);
      } else if (key.startsWith("cmp.")) {
        cmpKeys.push(key);
      }
    });
    if (refKeys.length) {
      window.localStorage.setItem(STORAGE_KEYS.selectedRef, JSON.stringify(refKeys));
    } else {
      window.localStorage.removeItem(STORAGE_KEYS.selectedRef);
    }
    if (cmpKeys.length) {
      window.localStorage.setItem(STORAGE_KEYS.selectedCmp, JSON.stringify(cmpKeys));
    } else {
      window.localStorage.removeItem(STORAGE_KEYS.selectedCmp);
    }
  } catch (error) {
    console.warn("Impossible de persister les colonnes sélectionnées", error);
  }
}

function updateSearchSummary() {
  if (!columnFilterLabel) {
    return;
  }
  if (!tableColumns.length) {
    columnFilterLabel.textContent = "Colonnes : Toutes";
    return;
  }
  if (!state.search.selectedKeys || !state.search.selectedKeys.size) {
    columnFilterLabel.textContent = "Colonnes : Toutes";
    return;
  }
  const labels = tableColumns
    .filter((column) => state.search.selectedKeys.has(column.key))
    .map((column) => column.label);
  if (!labels.length) {
    columnFilterLabel.textContent = "Colonnes : Toutes";
    return;
  }
  const preview = labels.slice(0, 3).join(", ");
  const suffix = labels.length > 3 ? ` +${labels.length - 3}` : "";
  columnFilterLabel.textContent = `Colonnes : ${preview}${suffix}`;
}

function announceColumnSelection(message) {
  if (!columnFilterLive) {
    return;
  }
  if (message) {
    columnFilterLive.textContent = message;
    return;
  }
  if (!tableColumns.length || !state.search.selectedKeys || !state.search.selectedKeys.size) {
    columnFilterLive.textContent = "Recherche sur toutes les colonnes.";
    return;
  }
  const count = state.search.selectedKeys.size;
  columnFilterLive.textContent = `Recherche limitée à ${count} colonne${count > 1 ? "s" : ""}.`;
}

function setColumnFilterInteractivity(disabled) {
  if (!columnFilterFieldset) {
    return;
  }
  columnFilterFieldset.disabled = Boolean(disabled);
  if (disabled) {
    columnFilterFieldset.setAttribute("aria-disabled", "true");
  } else {
    columnFilterFieldset.removeAttribute("aria-disabled");
  }
}

function syncSelectedKeysWithColumns({ loadStored = false } = {}) {
  if (!tableColumns.length) {
    state.search.selectedKeys = null;
    setColumnFilterInteractivity(true);
    return;
  }

  const availableKeys = new Set(tableColumns.map((column) => column.key));
  let selection = state.search.selectedKeys ? new Set(state.search.selectedKeys) : null;

  if (loadStored && currentMode === "compare") {
    const stored = loadSelectedKeysForComparison();
    if (stored && stored.size) {
      selection = stored;
    }
  }

  if (selection) {
    const filtered = Array.from(selection).filter((key) => availableKeys.has(key));
    selection = filtered.length ? new Set(filtered) : null;
  }

  if (selection && selection.size === availableKeys.size) {
    selection = null;
  }

  state.search.selectedKeys = selection;
  setColumnFilterInteractivity(false);
  updateSearchSummary();
  announceColumnSelection();
}

function getDefaultColumnIndexes() {
  if (tableColumns.length) {
    return tableColumns.map((_, index) => index);
  }
  const candidateLength = Math.max(headers.length, rawRows[0]?.length || 0);
  if (!candidateLength) {
    return [];
  }
  return Array.from({ length: candidateLength }, (_, index) => index);
}

function getColumnIndexesForSearch() {
  if (!state.search.selectedKeys || !state.search.selectedKeys.size) {
    return getDefaultColumnIndexes();
  }
  const indexes = [];
  state.search.selectedKeys.forEach((key) => {
    const index = columnKeyToIndex.get(key);
    if (typeof index === "number") {
      indexes.push(index);
    }
  });
  return indexes.length ? indexes : getDefaultColumnIndexes();
}

function setAllColumnsSelected(selectAll) {
  if (!tableColumns.length) {
    state.search.selectedKeys = null;
    updateSearchSummary();
    announceColumnSelection();
    return;
  }
  state.search.selectedKeys = null;
  updateSearchSummary();
  announceColumnSelection(
    selectAll
      ? "Recherche sur toutes les colonnes."
      : "Aucune colonne sélectionnée, réinitialisation sur Toutes les colonnes."
  );
  if (currentMode === "compare") {
    saveSelectedKeysForComparison(state.search.selectedKeys);
  }
}

function handleColumnToggle(input, checked) {
  if (!input) {
    return;
  }
  const key = input.dataset.columnKey;
  if (!key) {
    return;
  }
  if (!tableColumns.length) {
    return;
  }
  let selection = state.search.selectedKeys;
  if (!selection) {
    selection = new Set(tableColumns.map((column) => column.key));
  } else {
    selection = new Set(selection);
  }

  if (checked) {
    selection.add(key);
  } else {
    selection.delete(key);
  }

  if (!selection.size) {
    state.search.selectedKeys = null;
    updateSearchSummary();
    announceColumnSelection("Aucune colonne sélectionnée, réinitialisation sur Toutes les colonnes.");
    input.checked = true;
  } else if (selection.size === tableColumns.length) {
    state.search.selectedKeys = null;
    updateSearchSummary();
    announceColumnSelection();
  } else {
    state.search.selectedKeys = selection;
    updateSearchSummary();
    announceColumnSelection();
  }

  if (currentMode === "compare") {
    saveSelectedKeysForComparison(state.search.selectedKeys);
  }
  scheduleSearch();
}

function handleHeaderToggleChange(event) {
  const target = event.target;
  if (!target || target.type !== "checkbox") {
    return;
  }
  if (target.dataset.toggleType === "all") {
    if (!target.checked) {
      // Prevent leaving everything unchecked.
      setAllColumnsSelected(false);
      // Immediately reset checkbox to checked state for UI consistency.
      target.checked = true;
    } else {
      setAllColumnsSelected(true);
    }
    scheduleSearch();
    return;
  }
  handleColumnToggle(target, target.checked);
}

function clearComparisonState() {
  referenceKeywords = [];
  referenceHeaders = [];
  referenceRows = [];
  comparisonHeaders = [];
  comparisonRows = [];
  referenceFileName = "";
  comparisonFileName = "";
  referenceName.textContent = "";
  comparisonName.textContent = "";
  keywordsSummary.textContent = "";
  if (referenceInput) {
    referenceInput.value = "";
  }
  if (comparisonInput) {
    comparisonInput.value = "";
  }
}

function setMode(mode) {
  if (!mode || (mode !== "single" && mode !== "compare")) {
    return;
  }
  if (currentMode === mode) {
    return;
  }
  currentMode = mode;
  if (singleUploadContainer) {
    singleUploadContainer.hidden = mode !== "single";
  }
  if (compareUploadContainer) {
    compareUploadContainer.hidden = mode !== "compare";
  }
  resetDataset();
  clearError();
  updateProgress(0, "");
  if (mode === "single") {
    clearComparisonState();
    if (fileInput) {
      fileInput.value = "";
    }
  } else if (fileInput) {
    fileInput.value = "";
  }
}

function sanitizeValue(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return typeof value === "string" ? value : String(value);
}

function normalizeHeaders(headers, prefix = "") {
  const seen = new Map();
  return headers.map((header, index) => {
    let label = sanitizeValue(header).trim();
    if (!label) {
      label = `Colonne ${index + 1}`;
    }
    const base = prefix ? `${prefix}${label}` : label;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base} (${count + 1})`;
  });
}

function sanitizeRow(row) {
  if (!Array.isArray(row)) {
    return [];
  }
  return row.map((value) => sanitizeValue(value));
}

function normalizeParsedData(parsed) {
  const rawHeaders = Array.isArray(parsed?.headers) ? parsed.headers : [];
  const rawRows = Array.isArray(parsed?.rows) ? parsed.rows.slice() : [];

  let headerKeys = rawHeaders.map((header) => sanitizeValue(header));
  let workingRows = rawRows;

  if (!headerKeys.length && workingRows.length) {
    const firstRow = workingRows[0];
    if (Array.isArray(firstRow)) {
      headerKeys = firstRow.map((value, index) => sanitizeValue(value));
      workingRows = workingRows.slice(1);
    } else if (firstRow && typeof firstRow === "object") {
      headerKeys = Object.keys(firstRow);
    }
  }

  if (!headerKeys.length) {
    return { headers: [], rows: [] };
  }

  const normalizedHeaders = normalizeHeaders(headerKeys);

  const normalizedRows = workingRows.map((row) => {
    if (Array.isArray(row)) {
      return normalizedHeaders.map((_, index) => sanitizeValue(row[index]));
    }
    if (row && typeof row === "object") {
      return normalizedHeaders.map((normalizedHeader, index) => {
        const key = headerKeys[index] ?? normalizedHeader;
        const fallback = `Colonne ${index + 1}`;
        const value = Object.prototype.hasOwnProperty.call(row, key)
          ? row[key]
          : Object.prototype.hasOwnProperty.call(row, normalizedHeader)
          ? row[normalizedHeader]
          : row[fallback];
        return sanitizeValue(value);
      });
    }
    return new Array(normalizedHeaders.length).fill("");
  });

  return { headers: normalizedHeaders, rows: normalizedRows };
}

function extractKeywords(rows) {
  const keywords = new Set();
  rows.forEach((row) => {
    row.forEach((cell) => {
      const value = sanitizeValue(cell).trim();
      if (value) {
        keywords.add(value);
      }
    });
  });
  return Array.from(keywords);
}

function applyDataset({ columns, rows, fileName, includeMatchesColumn = false, loadStored = false }) {
  tableColumns = Array.isArray(columns)
    ? columns.map((column, index) => ({
        key: column.key || String(index),
        label: column.label || sanitizeColumnLabel("", index),
        origin: column.origin || "single",
        index,
      }))
    : [];
  columnKeyToIndex = new Map(tableColumns.map((column, index) => [column.key, index]));

  headers = tableColumns.map((column) => column.label);
  if (includeMatchesColumn) {
    matchesColumnIndex = headers.length;
    headers.push("Mots-clés trouvés");
  } else {
    matchesColumnIndex = -1;
  }

  const searchColumnCount = tableColumns.length;
  rawRows = Array.isArray(rows)
    ? rows.map((row) => {
        const normalized = Array.from({ length: searchColumnCount }, (_, index) =>
          sanitizeValue(row?.[index])
        );
        if (includeMatchesColumn) {
          normalized.push("");
        }
        return normalized;
      })
    : [];

  filteredRows = rawRows.map((row) => row.slice());
  filteredRowHighlights = filteredRows.map(() => new Array(headers.length).fill(null));
  buildCaches();
  syncSelectedKeysWithColumns({ loadStored });
  if (currentMode === "compare") {
    saveSelectedKeysForComparison(state.search.selectedKeys);
  }

  currentFileName = fileName || "";
  controlsSection.hidden = false;
  resultsSection.hidden = false;
  renderPage(1);
  if (searchInput.value.trim()) {
    performSearch();
  } else {
    announceStatus(
      `${filteredRows.length.toLocaleString()} ligne${
        filteredRows.length > 1 ? "s" : ""
      } affichée${filteredRows.length > 1 ? "s" : ""}.`
    );
  }
}

function getExt(fileName) {
  if (!fileName || typeof fileName !== "string") {
    return "unknown";
  }
  const match = fileName.toLowerCase().match(/\.([^.]+)$/);
  if (!match) {
    return "unknown";
  }
  const extension = match[1];
  if (extension === "csv" || extension === "xlsx" || extension === "xls") {
    return extension;
  }
  return "unknown";
}

function readFileContent(file, { ext, onProgress } = {}) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    activeReaders.add(reader);

    reader.onerror = () => {
      activeReaders.delete(reader);
      const error = reader.error || new Error("Lecture du fichier impossible.");
      reject(error);
    };

    reader.onabort = () => {
      activeReaders.delete(reader);
      const abortError = new Error("Lecture interrompue.");
      abortError.name = "AbortError";
      reject(abortError);
    };

    reader.onload = () => {
      activeReaders.delete(reader);
      resolve(reader.result);
    };

    reader.onprogress = (event) => {
      if (typeof onProgress === "function" && event.lengthComputable) {
        const total = event.total || file.size;
        const percent = Math.min(60, Math.round((event.loaded / total) * 60));
        const label = `Lecture ${formatBytes(event.loaded)} / ${formatBytes(total)}`;
        onProgress(percent, label);
      }
    };

    try {
      if (ext === "csv") {
        reader.readAsText(file);
      } else {
        reader.readAsArrayBuffer(file);
      }
    } catch (error) {
      activeReaders.delete(reader);
      reject(error);
    }
  });
}

function stripBom(text) {
  if (text.charCodeAt(0) === 0xfeff) {
    return text.slice(1);
  }
  return text;
}

function detectDelimiter(text) {
  const delimiters = [",", ";", "\t", "|"];
  const sample = text.slice(0, 10000);
  const lines = sample.split(/\r\n|\n|\r/).filter((line) => line.length).slice(0, 20);
  if (!lines.length) {
    return ",";
  }
  let bestDelimiter = ",";
  let bestScore = -Infinity;
  delimiters.forEach((delimiter) => {
    let score = 0;
    let previousCount = null;
    lines.forEach((line) => {
      const count = line.split(delimiter).length - 1;
      if (count > 0) {
        score += count;
        if (previousCount !== null && previousCount !== count) {
          score -= Math.abs(previousCount - count);
        }
        previousCount = count;
      }
    });
    if (score > bestScore) {
      bestScore = score;
      bestDelimiter = delimiter;
    }
  });
  return bestScore <= 0 ? "," : bestDelimiter;
}

function parseCSV(text, { onProgress } = {}) {
  return new Promise((resolve, reject) => {
    try {
      if (!text) {
        resolve({ headers: [], rows: [] });
        return;
      }
      const content = stripBom(text);
      const delimiter = detectDelimiter(content);
      const rows = [];
      let row = [];
      let field = "";
      let inQuotes = false;
      let index = 0;
      const length = content.length;
      const chunkSize = 200000;

      const commitRow = () => {
        const shouldKeep = rows.length === 0 || row.some((value) => value !== "");
        if (shouldKeep) {
          rows.push(row.slice());
        }
        row = [];
      };

      const pushField = () => {
        row.push(field);
        field = "";
      };

      const processChunk = () => {
        const limit = Math.min(index + chunkSize, length);
        while (index < limit) {
          const char = content[index];
          if (char === '"') {
            if (inQuotes) {
              if (content[index + 1] === '"') {
                field += '"';
                index += 1;
              } else {
                inQuotes = false;
              }
            } else {
              inQuotes = true;
            }
            index += 1;
            continue;
          }

          if (!inQuotes && char === delimiter) {
            pushField();
            index += 1;
            continue;
          }

          if (!inQuotes && (char === "\n" || char === "\r")) {
            if (char === "\r" && content[index + 1] === "\n") {
              index += 1;
            }
            pushField();
            commitRow();
            index += 1;
            continue;
          }

          field += char;
          index += 1;
        }

        if (index < length) {
          if (typeof onProgress === "function") {
            const percent = 60 + Math.min(38, Math.round((index / length) * 38));
            const label = `${rows.length.toLocaleString()} lignes analysées`;
            onProgress(percent, label);
          }
          setTimeout(processChunk, 0);
        } else {
          pushField();
          if (row.length && row.some((value) => value !== "")) {
            commitRow();
          } else if (rows.length === 0) {
            commitRow();
          } else {
            row = [];
          }

          const [headerRow = [], ...dataRows] = rows;
          if (headerRow.length === 0 && dataRows.length === 0) {
            resolve({ headers: [], rows: [] });
            return;
          }

          const headers = normalizeHeaders(headerRow);
          const filteredRows = dataRows.filter((cells) =>
            Array.isArray(cells) ? cells.some((value) => value !== "") : false
          );
          const objects = filteredRows.map((cells) => {
            const entry = {};
            headers.forEach((header, columnIndex) => {
              entry[header] = sanitizeValue(cells[columnIndex]);
            });
            return entry;
          });
          resolve({ headers, rows: objects });
        }
      };

      processChunk();
    } catch (error) {
      reject(error);
    }
  });
}

function parseXLSX(arrayBuffer, { onProgress } = {}) {
  if (typeof XLSX === "undefined") {
    throw new Error("La bibliothèque XLSX est indisponible.");
  }
  if (typeof onProgress === "function") {
    onProgress(75, "Analyse du classeur...");
  }
  const workbook = XLSX.read(arrayBuffer, { type: "array", dense: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    return { headers: [], rows: [] };
  }
  if (typeof onProgress === "function") {
    onProgress(85, `Lecture de la feuille ${sheetName}`);
  }
  const sheet = workbook.Sheets[sheetName];
  const sheetData = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    blankrows: false,
    defval: "",
  });
  const [headerRow = [], ...dataRows] = sheetData;
  if (!headerRow.length && !dataRows.length) {
    return { headers: [], rows: [] };
  }
  const headers = normalizeHeaders(headerRow);
  const objects = dataRows
    .map((cells) => {
      const entry = {};
      headers.forEach((header, columnIndex) => {
        entry[header] = sanitizeValue(cells?.[columnIndex]);
      });
      return entry;
    })
    .filter((entry) => headers.some((header) => entry[header] !== ""));
  return { headers, rows: objects };
}

async function importAnyFile(file, { onProgress } = {}) {
  if (!file) {
    throw new Error("Aucun fichier fourni.");
  }
  if (file.size > MAX_FILE_SIZE) {
    throw new Error("Fichier trop volumineux, essayez de le scinder.");
  }
  const extension = getExt(file.name);
  if (extension === "unknown") {
    throw new Error("Format non supporté. Choisissez un CSV, XLS ou XLSX.");
  }

  if (typeof onProgress === "function") {
    onProgress(1, `Lecture de ${file.name} (${formatBytes(file.size)})`);
  }

  const content = await readFileContent(file, { ext: extension, onProgress });

  if (extension === "csv") {
    const text = typeof content === "string" ? content : new TextDecoder("utf-8").decode(content);
    return parseCSV(text, { onProgress });
  }

  return parseXLSX(content, { onProgress });
}

function updateReferenceSummary() {
  if (!keywordsSummary) {
    return;
  }
  if (!referenceKeywords.length) {
    keywordsSummary.textContent = "Aucun mot-clé valide trouvé dans le fichier de référence.";
    return;
  }
  if (referenceKeywords.length === 1) {
    keywordsSummary.textContent = "1 mot-clé extrait du fichier de référence.";
    return;
  }
  keywordsSummary.textContent = `${referenceKeywords.length.toLocaleString()} mots-clés extraits du fichier de référence.`;
}

function updateComparisonDataset() {
  if (!comparisonHeaders.length && !comparisonRows.length) {
    return;
  }

  const refColumns = getColumnsFor("ref", referenceRows, referenceHeaders);
  const cmpColumns = getColumnsFor("cmp", comparisonRows, comparisonHeaders);
  const columns = [...refColumns, ...cmpColumns];

  if (!columns.length) {
    resetDataset();
    updateReferenceSummary();
    return;
  }

  const totalRows = Math.max(referenceRows.length, comparisonRows.length, 0);
  const combinedRows = Array.from({ length: totalRows }, (_, rowIndex) => {
    const refRow = referenceRows[rowIndex] || [];
    const cmpRow = comparisonRows[rowIndex] || [];
    const values = [];
    refColumns.forEach((column) => {
      values.push(sanitizeValue(refRow[column.index]));
    });
    cmpColumns.forEach((column) => {
      values.push(sanitizeValue(cmpRow[column.index]));
    });
    return values;
  });

  const fileLabel = comparisonFileName
    ? `${comparisonFileName}_comparaison`
    : referenceFileName
    ? `${referenceFileName}_comparaison`
    : "comparaison";

  applyDataset({
    columns,
    rows: combinedRows,
    fileName: fileLabel,
    includeMatchesColumn: true,
    loadStored: true,
  });
  updateReferenceSummary();
  updateProgress(100, "Comparaison terminée");
}

async function handleSingleFile(files) {
  const [file] = files ? Array.from(files).filter(Boolean) : [];
  if (!file) return;
  clearError();
  abortActiveReaders();
  clearComparisonState();
  if (fileInput) {
    fileInput.value = "";
  }

  currentFileName = file.name.replace(/\.[^.]+$/, "");
  let finalLabel;
  try {
    setLoading(true, `Lecture de ${file.name} (${formatBytes(file.size)})`);
    const parsed = await importAnyFile(file, {
      onProgress: (percent, label) => updateProgress(percent, label),
    });
    const { headers: parsedHeaders, rows } = normalizeParsedData(parsed);
    if (!rows.length) {
      showError("Aucune donnée trouvée dans le fichier.");
      resetDataset({ keepSpinner: true });
      finalLabel = "Chargement terminé";
      return;
    }

    const columns = getSingleFileColumns(parsedHeaders, rows);
    applyDataset({
      columns,
      rows,
      fileName: currentFileName,
      includeMatchesColumn: false,
      loadStored: false,
    });
    finalLabel = "Chargement terminé";
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") {
      finalLabel = "Chargement annulé";
    } else {
      console.error(error);
      showError(
        "Impossible de lire le fichier.",
        error instanceof Error ? error.message : ""
      );
      resetDataset();
      finalLabel = "Échec du chargement";
    }
  } finally {
    setLoading(false, finalLabel);
  }
}

async function handleReferenceFiles(files) {
  const [file] = files ? Array.from(files).filter(Boolean) : [];
  if (!file) return;
  clearError();
  abortActiveReaders();
  if (referenceInput) {
    referenceInput.value = "";
  }

  referenceName.textContent = file.name;
  referenceFileName = file.name.replace(/\.[^.]+$/, "");
  let finalLabel;
  try {
    setLoading(true, `Lecture de ${file.name} (${formatBytes(file.size)})`);
    const parsed = await importAnyFile(file, {
      onProgress: (percent, label) => updateProgress(percent, label),
    });
    const { headers: parsedHeaders, rows } = normalizeParsedData(parsed);
    referenceHeaders = parsedHeaders;
    referenceRows = rows;
    if (!rows.length) {
      referenceKeywords = [];
      referenceHeaders = [];
      referenceRows = [];
      updateReferenceSummary();
      showError("Aucune donnée trouvée dans le fichier de référence.");
      resetDataset({ keepSpinner: true });
      finalLabel = "Chargement terminé";
      return;
    }

    referenceKeywords = extractKeywords(rows);
    updateReferenceSummary();

    if (!comparisonHeaders.length && !comparisonRows.length) {
      resetDataset({ keepSpinner: true });
      finalLabel = "Fichier de référence chargé";
    } else {
      updateComparisonDataset();
      finalLabel = "Comparaison terminée";
    }
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") {
      finalLabel = "Chargement annulé";
    } else {
      console.error(error);
      showError(
        "Impossible de lire le fichier de référence.",
        error instanceof Error ? error.message : ""
      );
      finalLabel = "Échec du chargement";
    }
  } finally {
    setLoading(false, finalLabel);
  }
}

async function handleComparisonFiles(files) {
  const [file] = files ? Array.from(files).filter(Boolean) : [];
  if (!file) return;
  clearError();
  abortActiveReaders();
  if (comparisonInput) {
    comparisonInput.value = "";
  }

  comparisonName.textContent = file.name;
  comparisonFileName = file.name.replace(/\.[^.]+$/, "");
  let finalLabel;
  try {
    setLoading(true, `Lecture de ${file.name} (${formatBytes(file.size)})`);
    const parsed = await importAnyFile(file, {
      onProgress: (percent, label) => updateProgress(percent, label),
    });
    const { headers: parsedHeaders, rows } = normalizeParsedData(parsed);
    if (!rows.length) {
      comparisonHeaders = [];
      comparisonRows = [];
      resetDataset({ keepSpinner: true });
      showError("Aucune donnée trouvée dans le fichier à comparer.");
      finalLabel = "Chargement terminé";
      return;
    }

    comparisonHeaders = parsedHeaders;
    comparisonRows = rows;

    if (!referenceHeaders.length && !referenceRows.length) {
      resetDataset({ keepSpinner: true });
      finalLabel = "Fichier à comparer chargé";
    } else {
      updateComparisonDataset();
      finalLabel = "Comparaison terminée";
    }
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") {
      finalLabel = "Chargement annulé";
    } else {
      console.error(error);
      showError(
        "Impossible de lire le fichier à comparer.",
        error instanceof Error ? error.message : ""
      );
      finalLabel = "Échec du chargement";
    }
  } finally {
    setLoading(false, finalLabel);
  }
}

async function handleCompareDrop(files) {
  const fileList = Array.from(files || []).filter(Boolean);
  if (!fileList.length) {
    return;
  }

  if (fileList.length >= 2) {
    await handleReferenceFiles([fileList[0]]);
    await handleComparisonFiles([fileList[1]]);
    return;
  }

  if (!referenceHeaders.length && !referenceRows.length) {
    await handleReferenceFiles([fileList[0]]);
  } else {
    await handleComparisonFiles([fileList[0]]);
  }
}

function buildCaches() {
  const searchColumnCount = tableColumns.length || Math.max(headers.length, 0);
  rowTextCache = rawRows.map((row) =>
    Array.from({ length: searchColumnCount }, (_, index) => sanitizeValue(row?.[index]))
  );
  lowerRowTextCache = rowTextCache.map((cells) => cells.map((cell) => cell.toLowerCase()));
}

function renderTable(rows, { highlights = [] } = {}) {
  if (!dataTable) {
    return;
  }

  const tableHead = doc.createElement("thead");
  const headerRow = doc.createElement("tr");

  if (tableColumns.length) {
    const allSelected = !state.search.selectedKeys || !state.search.selectedKeys.size;
    tableColumns.forEach((column, index) => {
      const th = doc.createElement("th");
      th.className = "table__header";
      const wrapper = doc.createElement("div");
      wrapper.className = "table__header-inner";

      if (index === 0) {
        const allLabel = doc.createElement("label");
        allLabel.className = "col-toggle col-toggle--all";
        const allInput = doc.createElement("input");
        allInput.type = "checkbox";
        allInput.className = "col-toggle__input";
        allInput.dataset.toggleType = "all";
        allInput.checked = allSelected;
        allInput.title = "Tout / Rien";
        allInput.setAttribute(
          "aria-label",
          allSelected
            ? "Toutes les colonnes sont incluses dans la recherche"
            : "Réinitialiser la recherche sur toutes les colonnes"
        );
        const allText = doc.createElement("span");
        allText.className = "col-toggle__text";
        allText.textContent = "Tout / Rien";
        allLabel.appendChild(allInput);
        allLabel.appendChild(allText);
        wrapper.appendChild(allLabel);
      }

      const labelSpan = doc.createElement("span");
      labelSpan.className = "table__header-label";
      labelSpan.textContent = column.label;
      wrapper.appendChild(labelSpan);

      const toggleLabel = doc.createElement("label");
      toggleLabel.className = "col-toggle";
      toggleLabel.title = "Inclure cette colonne dans la recherche";
      const toggleInput = doc.createElement("input");
      toggleInput.type = "checkbox";
      toggleInput.className = "col-toggle__input";
      toggleInput.dataset.columnKey = column.key;
      toggleInput.checked = allSelected || (state.search.selectedKeys?.has(column.key) ?? false);
      toggleInput.setAttribute(
        "aria-label",
        `Inclure la colonne ${column.label} dans la recherche`
      );
      const toggleText = doc.createElement("span");
      toggleText.className = "visually-hidden";
      toggleText.textContent = `Inclure ${column.label} dans la recherche`;
      toggleLabel.appendChild(toggleInput);
      toggleLabel.appendChild(toggleText);
      wrapper.appendChild(toggleLabel);

      th.appendChild(wrapper);
      headerRow.appendChild(th);
    });

    if (matchesColumnIndex >= 0) {
      const th = doc.createElement("th");
      th.className = "table__header table__header--matches";
      th.textContent = headers[matchesColumnIndex] || "Mots-clés trouvés";
      headerRow.appendChild(th);
    }
  } else {
    const totalColumns = headers.length || (rows[0]?.length ?? 0);
    for (let index = 0; index < totalColumns; index += 1) {
      const th = doc.createElement("th");
      const header = headers[index];
      th.textContent = header === undefined || header === null || header === ""
        ? `Colonne ${index + 1}`
        : String(header);
      headerRow.appendChild(th);
    }
  }

  tableHead.appendChild(headerRow);

  const tableBody = doc.createElement("tbody");
  dataTable.innerHTML = "";
  const fragment = doc.createDocumentFragment();
  fragment.appendChild(tableHead);
  fragment.appendChild(tableBody);
  dataTable.appendChild(fragment);

  const totalRows = rows.length;
  if (totalRows === 0) {
    return;
  }

  let rowIndex = 0;
  const processBatch = () => {
    const batchFragment = doc.createDocumentFragment();
    const limit = Math.min(rowIndex + RENDER_BATCH_SIZE, totalRows);
    for (; rowIndex < limit; rowIndex += 1) {
      const tr = doc.createElement("tr");
      const currentRow = rows[rowIndex];
      const rowHighlight = highlights[rowIndex] || [];
      for (let index = 0; index < headers.length; index += 1) {
        const td = doc.createElement("td");
        if (matchesColumnIndex >= 0 && index === matchesColumnIndex) {
          td.classList.add("matches-cell");
        }
        const highlightContent = rowHighlight[index];
        if (highlightContent) {
          td.innerHTML = highlightContent;
        } else {
          const value = currentRow?.[index];
          td.textContent = value === undefined || value === null ? "" : String(value);
        }
        tr.appendChild(td);
      }
      batchFragment.appendChild(tr);
    }
    tableBody.appendChild(batchFragment);
    if (rowIndex < totalRows) {
      setTimeout(processBatch, 0);
    }
  };

  processBatch();
}

function renderPage(pageNumber) {
  if (!dataTable) {
    if (filteredRows.length === 0) {
      resultStats.textContent = "0 résultat";
      pagination.hidden = true;
    }
    return;
  }
  if (filteredRows.length === 0) {
    renderTable([], { highlights: [] });
    const tbody = dataTable.querySelector("tbody");
    if (tbody) {
      const emptyRow = doc.createElement("tr");
      const emptyCell = doc.createElement("td");
      const totalColumns = headers.length || (rawRows[0]?.length ?? 1);
      emptyCell.colSpan = totalColumns || 1;
      emptyCell.textContent = "Aucune ligne correspondante.";
      emptyRow.appendChild(emptyCell);
      tbody.appendChild(emptyRow);
    }
    resultStats.textContent = "0 résultat";
    pagination.hidden = true;
    return;
  }

  const totalPages = Math.ceil(filteredRows.length / PAGE_SIZE);
  currentPage = Math.min(Math.max(pageNumber, 1), totalPages);
  const start = (currentPage - 1) * PAGE_SIZE;
  const pageRows = filteredRows.slice(start, start + PAGE_SIZE);
  const pageHighlights = filteredRowHighlights.slice(start, start + PAGE_SIZE);

  renderTable(pageRows, { highlights: pageHighlights });

  const totalText = `${filteredRows.length.toLocaleString()} ligne${
    filteredRows.length > 1 ? "s" : ""
  } trouvée${filteredRows.length > 1 ? "s" : ""}`;
  resultStats.textContent = `${totalText} (sur ${rawRows.length.toLocaleString()} lignes)`;

  if (totalPages > 1) {
    pagination.hidden = false;
    pageInfo.textContent = `Page ${currentPage} / ${totalPages}`;
    prevPageBtn.disabled = currentPage === 1;
    nextPageBtn.disabled = currentPage === totalPages;
  } else {
    pagination.hidden = true;
  }
}

function tokenizeQuery(query) {
  const tokens = [];
  const regex = /"([^"]+)"|\(|\)|\bAND\b|\bOR\b|\bNOT\b|[^\s,()]+/gi;
  let match;

  while ((match = regex.exec(query))) {
    let token = match[0].trim();
    if (!token) continue;
    if (token.endsWith(",")) {
      token = token.slice(0, -1);
    }
    if (!token) continue;
    if (token.startsWith("\"") && token.endsWith("\"")) {
      token = token.slice(1, -1);
    }

    const upper = token.toUpperCase();
    if (["AND", "OR", "NOT", "(", ")"].includes(upper)) {
      tokens.push({ type: "operator", value: upper });
    } else {
      tokens.push({ type: "operand", value: token });
    }
  }
  return tokens;
}

function toPostfix(tokens) {
  const output = [];
  const stack = [];
  const precedence = { NOT: 3, AND: 2, OR: 1 };
  const rightAssociative = { NOT: true };

  tokens.forEach((token) => {
    if (token.type === "operand") {
      output.push(token);
    } else if (token.value === "(") {
      stack.push(token);
    } else if (token.value === ")") {
      while (stack.length && stack[stack.length - 1].value !== "(") {
        output.push(stack.pop());
      }
      if (!stack.length) {
        throw new Error("Parenthèses déséquilibrées.");
      }
      stack.pop();
    } else {
      while (
        stack.length &&
        stack[stack.length - 1].type === "operator" &&
        stack[stack.length - 1].value !== "(" &&
        (precedence[stack[stack.length - 1].value] > precedence[token.value] ||
          (precedence[stack[stack.length - 1].value] === precedence[token.value] &&
            !rightAssociative[token.value]))
      ) {
        output.push(stack.pop());
      }
      stack.push(token);
    }
  });

  while (stack.length) {
    const op = stack.pop();
    if (op.value === "(" || op.value === ")") {
      throw new Error("Parenthèses déséquilibrées.");
    }
    output.push(op);
  }

  return output;
}

function matchRow(rowIndex, keyword, { caseSensitive, exactMatch }, columnIndexes) {
  const cells = caseSensitive ? rowTextCache[rowIndex] : lowerRowTextCache[rowIndex];
  if (!cells) {
    return false;
  }
  const query = caseSensitive ? keyword : keyword.toLowerCase();
  if (!query) {
    return false;
  }
  const indexes = columnIndexes && columnIndexes.length ? columnIndexes : getDefaultColumnIndexes();
  for (let idx = 0; idx < indexes.length; idx += 1) {
    const columnIndex = indexes[idx];
    const cellValue = cells[columnIndex];
    if (!cellValue) {
      continue;
    }
    if (exactMatch) {
      if (cellValue === query) {
        return true;
      }
    } else if (cellValue.includes(query)) {
      return true;
    }
  }
  return false;
}

function evaluateQuery(tokens, options) {
  if (!tokens.length) {
    return rawRows.map((_, index) => index);
  }

  const postfix = toPostfix(tokens);
  const matches = [];
  const columnIndexes =
    options.columnIndexes && options.columnIndexes.length
      ? options.columnIndexes
      : getDefaultColumnIndexes();

  for (let index = 0; index < rawRows.length; index += 1) {
    const stack = [];
    for (const token of postfix) {
      if (token.type === "operand") {
        stack.push(matchRow(index, token.value, options, columnIndexes));
      } else if (token.value === "NOT") {
        const value = stack.pop();
        stack.push(!value);
      } else {
        const right = stack.pop();
        const left = stack.pop();
        if (token.value === "AND") {
          stack.push(Boolean(Boolean(left) && Boolean(right)));
        } else if (token.value === "OR") {
          stack.push(Boolean(Boolean(left) || Boolean(right)));
        }
      }
    }
    const result = stack.pop();
    if (stack.length) {
      throw new Error("Expression booléenne invalide.");
    }
    if (result) {
      matches.push(index);
    }
  }

  return matches;
}

function computeRowHighlights(rowIndex, tokens, options, columnIndexes) {
  const highlights = new Array(headers.length).fill(null);
  if (!tokens.length) {
    return { summary: "", highlights };
  }

  const indexes = columnIndexes && columnIndexes.length ? columnIndexes : getDefaultColumnIndexes();
  const comparisonCells = options.caseSensitive ? rowTextCache[rowIndex] : lowerRowTextCache[rowIndex];
  const originalCells = rowTextCache[rowIndex] || [];
  const matches = [];
  const highlightMap = new Map();

  tokens.forEach((token) => {
    const value = token === undefined || token === null ? "" : String(token);
    const trimmed = value.trim();
    if (!trimmed) {
      return;
    }
    const normalized = options.caseSensitive ? trimmed : trimmed.toLowerCase();
    const matchedColumns = [];

    indexes.forEach((columnIndex) => {
      const haystack = comparisonCells?.[columnIndex];
      if (haystack === undefined || haystack === null) {
        return;
      }
      const matchFound = options.exactMatch ? haystack === normalized : haystack.includes(normalized);
      if (matchFound) {
        const columnLabel = tableColumns[columnIndex]?.label || sanitizeColumnLabel("", columnIndex);
        matchedColumns.push(columnLabel);
        if (!highlightMap.has(columnIndex)) {
          highlightMap.set(columnIndex, new Set());
        }
        highlightMap.get(columnIndex).add(trimmed);
      }
    });

    if (matchedColumns.length) {
      matches.push({ token: trimmed, columns: matchedColumns });
    }
  });

  highlightMap.forEach((tokensSet, columnIndex) => {
    const cellValue = originalCells[columnIndex] ?? "";
    highlights[columnIndex] = highlightValue(cellValue, tokensSet, options);
  });

  const summary = matches
    .map((entry) => `${entry.token} (${entry.columns.join(", ")})`)
    .join("; ");

  return { summary, highlights };
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function highlightValue(value, tokensSet, options) {
  const text = sanitizeValue(value);
  if (!text) {
    return "";
  }
  const tokens = Array.from(tokensSet || []).filter(Boolean);
  if (!tokens.length) {
    return escapeHtml(text);
  }

  const normalizedText = options.caseSensitive ? text : text.toLowerCase();
  const ranges = [];

  tokens.forEach((token) => {
    const normalizedToken = options.caseSensitive ? token : token.toLowerCase();
    if (!normalizedToken) {
      return;
    }
    if (options.exactMatch) {
      if (normalizedText === normalizedToken) {
        ranges.push({ start: 0, end: text.length });
      }
      return;
    }
    let startIndex = 0;
    while (startIndex <= normalizedText.length) {
      const index = normalizedText.indexOf(normalizedToken, startIndex);
      if (index === -1) {
        break;
      }
      ranges.push({ start: index, end: index + normalizedToken.length });
      startIndex = index + Math.max(normalizedToken.length, 1);
    }
  });

  if (!ranges.length) {
    return escapeHtml(text);
  }

  ranges.sort((a, b) => (a.start === b.start ? a.end - b.end : a.start - b.start));
  const merged = [];
  ranges.forEach((range) => {
    const last = merged[merged.length - 1];
    if (!last || range.start > last.end) {
      merged.push({ start: range.start, end: range.end });
    } else if (range.end > last.end) {
      last.end = range.end;
    }
  });

  let result = "";
  let cursor = 0;
  merged.forEach((range) => {
    if (range.start > cursor) {
      result += escapeHtml(text.slice(cursor, range.start));
    }
    result += `<mark>${escapeHtml(text.slice(range.start, range.end))}</mark>`;
    cursor = range.end;
  });
  if (cursor < text.length) {
    result += escapeHtml(text.slice(cursor));
  }
  return result || escapeHtml(text);
}

function performSearch() {
  clearError();
  const query = searchInput.value.trim();
  state.search.query = query;
  state.search.caseSensitive = Boolean(caseSensitiveToggle.checked);
  state.search.exactMatch = Boolean(exactMatchToggle.checked);
  const columnIndexes = getColumnIndexesForSearch();

  if (!query) {
    filteredRows = rawRows.map((row) => row.slice());
    filteredRowHighlights = filteredRows.map(() => new Array(headers.length).fill(null));
    renderPage(1);
    announceStatus(
      `${filteredRows.length.toLocaleString()} ligne${
        filteredRows.length > 1 ? "s" : ""
      } affichée${filteredRows.length > 1 ? "s" : ""}.`
    );
    return;
  }

  try {
    const tokens = tokenizeQuery(query);
    const indexes = evaluateQuery(tokens, {
      caseSensitive: state.search.caseSensitive,
      exactMatch: state.search.exactMatch,
      columnIndexes,
    });
    const operandTokens = tokens
      .filter((token) => token.type === "operand")
      .map((token) => token.value)
      .filter((value) => value !== undefined);

    filteredRows = indexes.map((rowIndex) => rawRows[rowIndex].slice());
    filteredRowHighlights = indexes.map(() => new Array(headers.length).fill(null));

    if (operandTokens.length) {
      const options = {
        caseSensitive: state.search.caseSensitive,
        exactMatch: state.search.exactMatch,
      };
      filteredRows = indexes.map((rowIndex, position) => {
        const baseRow = rawRows[rowIndex].slice();
        const { summary, highlights } = computeRowHighlights(
          rowIndex,
          operandTokens,
          options,
          columnIndexes
        );
        if (matchesColumnIndex >= 0) {
          baseRow[matchesColumnIndex] = summary;
        }
        filteredRowHighlights[position] = highlights;
        return baseRow;
      });
    }

    renderPage(1);
    announceStatus(
      `${filteredRows.length.toLocaleString()} ligne${
        filteredRows.length > 1 ? "s" : ""
      } trouvée${filteredRows.length > 1 ? "s" : ""}.`
    );
  } catch (error) {
    console.error(error);
    showError(error.message || "Requête invalide");
  }
}

function resetSearch() {
  scheduleSearch.cancel();
  searchInput.value = "";
  caseSensitiveToggle.checked = false;
  exactMatchToggle.checked = false;
  state.search.query = "";
  state.search.caseSensitive = false;
  state.search.exactMatch = false;
  if (currentMode === "compare" && comparisonRows.length) {
    updateComparisonDataset();
  } else {
    filteredRows = rawRows.map((row) => row.slice());
    filteredRowHighlights = filteredRows.map(() => new Array(headers.length).fill(null));
    renderPage(1);
    announceStatus(
      `${filteredRows.length.toLocaleString()} ligne${
        filteredRows.length > 1 ? "s" : ""
      } affichée${filteredRows.length > 1 ? "s" : ""}.`
    );
  }
}

function resetApplication() {
  abortActiveReaders();
  scheduleSearch.cancel();
  state.search.query = "";
  state.search.caseSensitive = false;
  state.search.exactMatch = false;
  state.search.selectedKeys = null;
  if (searchInput) {
    searchInput.value = "";
  }
  if (caseSensitiveToggle) {
    caseSensitiveToggle.checked = false;
  }
  if (exactMatchToggle) {
    exactMatchToggle.checked = false;
  }
  clearComparisonState();
  if (fileInput) {
    fileInput.value = "";
  }
  resetDataset();
  clearError();
  saveSelectedKeysForComparison(null);
  announceStatus("Application réinitialisée.");
}

function getCurrentPageRows() {
  if (!filteredRows.length) return [];
  const start = (currentPage - 1) * PAGE_SIZE;
  return filteredRows.slice(start, start + PAGE_SIZE);
}

async function copyToClipboard() {
  if (!filteredRows.length) return;
  const csvContent = convertRowsToCsv(headers, filteredRows);
  try {
    if (typeof navigator === "undefined" || !navigator.clipboard) {
      throw new Error("Clipboard API indisponible");
    }
    await navigator.clipboard.writeText(csvContent);
    clearError();
    updateProgress(100, "Résultats copiés dans le presse-papiers");
    announceStatus("Résultats copiés dans le presse-papiers.");
  } catch (error) {
    showError("Impossible de copier dans le presse-papiers.");
  }
}

function sanitizeCellForExport(cell) {
  if (cell === null || cell === undefined) {
    return "";
  }
  const value = String(cell);
  if (!value.includes("<")) {
    return value;
  }
  return value.replace(/<\/?mark[^>]*>/gi, "");
}

function sanitizeRowsForExport(rows) {
  return rows.map((row) => row.map((cell) => sanitizeCellForExport(cell)));
}

function convertRowsToCsv(headers, rows) {
  const safeHeaders = headers.map((header) => sanitizeCellForExport(header));
  const safeRows = sanitizeRowsForExport(rows);
  const allRows = [safeHeaders, ...safeRows];
  return allRows
    .map((row) =>
      row
        .map((cell) => {
          const value = cell === null || cell === undefined ? "" : String(cell);
          if (value.includes("\"") || value.includes(",") || value.includes("\n")) {
            return `"${value.replace(/"/g, '""')}"`;
          }
          return value;
        })
        .join(",")
    )
    .join("\n");
}

function downloadBlob(content, filename, type) {
  if (!doc) return;
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = doc.createElement("a");
  link.href = url;
  link.download = filename;
  doc.body.appendChild(link);
  link.click();
  doc.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function exportCsv(rows) {
  if (!rows.length) return;
  const csvContent = convertRowsToCsv(headers, rows);
  const filename = `${currentFileName || "export"}_resultats.csv`;
  downloadBlob(csvContent, filename, "text/csv;charset=utf-8;");
  announceStatus("Export CSV généré avec le filtre courant.");
}

function exportXlsx(rows) {
  if (!rows.length) return;
  const worksheetData = [
    headers.map((header) => sanitizeCellForExport(header)),
    ...sanitizeRowsForExport(rows),
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Résultats");
  const filename = `${currentFileName || "export"}_resultats.xlsx`;
  XLSX.writeFile(workbook, filename, { compression: true });
  announceStatus("Export XLSX généré avec le filtre courant.");
}

function attachEvents() {
  if (
    !dropZone ||
    !searchButton ||
    !resetButton ||
    !prevPageBtn ||
    !nextPageBtn ||
    !copyButton ||
    !exportCsvButton ||
    !exportXlsxButton
  ) {
    return;
  }

  if (fileInput) {
    fileInput.addEventListener("change", (event) => {
      handleSingleFile(event.target.files);
    });
  }

  if (referenceInput) {
    referenceInput.addEventListener("change", (event) => {
      handleReferenceFiles(event.target.files);
    });
  }

  if (comparisonInput) {
    comparisonInput.addEventListener("change", (event) => {
      handleComparisonFiles(event.target.files);
    });
  }

  modeRadios.forEach((radio) => {
    radio.addEventListener("change", (event) => {
      if (event.target.checked) {
        setMode(event.target.value);
      }
    });
  });

  const activeMode = modeRadios.find((radio) => radio.checked)?.value;
  if (activeMode === "compare") {
    currentMode = "compare";
    if (singleUploadContainer) {
      singleUploadContainer.hidden = true;
    }
    if (compareUploadContainer) {
      compareUploadContainer.hidden = false;
    }
  } else {
    currentMode = "single";
    if (singleUploadContainer) {
      singleUploadContainer.hidden = false;
    }
    if (compareUploadContainer) {
      compareUploadContainer.hidden = true;
    }
  }

  ["dragenter", "dragover"].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      event.stopPropagation();
      dropZone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      event.stopPropagation();
      dropZone.classList.remove("dragover");
    });
  });

  dropZone.addEventListener("drop", async (event) => {
    const files = event.dataTransfer?.files;
    if (currentMode === "compare") {
      await handleCompareDrop(files);
    } else {
      await handleSingleFile(files);
    }
  });

  searchButton.addEventListener("click", () => {
    scheduleSearch.cancel();
    performSearch();
  });
  searchInput.addEventListener("input", () => {
    scheduleSearch();
  });
  searchInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      scheduleSearch.cancel();
      performSearch();
    }
  });

  resetButton.addEventListener("click", () => {
    resetApplication();
  });

  if (caseSensitiveToggle && typeof caseSensitiveToggle.addEventListener === "function") {
    caseSensitiveToggle.addEventListener("change", () => {
      state.search.caseSensitive = Boolean(caseSensitiveToggle.checked);
      if (currentMode === "compare" && comparisonRows.length) {
        updateComparisonDataset();
      } else {
        scheduleSearch();
      }
    });
  }

  if (exactMatchToggle && typeof exactMatchToggle.addEventListener === "function") {
    exactMatchToggle.addEventListener("change", () => {
      state.search.exactMatch = Boolean(exactMatchToggle.checked);
      if (currentMode === "compare" && comparisonRows.length) {
        updateComparisonDataset();
      } else {
        scheduleSearch();
      }
    });
  }

  if (dataTable) {
    dataTable.addEventListener("change", handleHeaderToggleChange);
  }

  prevPageBtn.addEventListener("click", () => {
    if (currentPage > 1) {
      renderPage(currentPage - 1);
    }
  });

  nextPageBtn.addEventListener("click", () => {
    const totalPages = Math.ceil(filteredRows.length / PAGE_SIZE);
    if (currentPage < totalPages) {
      renderPage(currentPage + 1);
    }
  });

  copyButton.addEventListener("click", copyToClipboard);
  exportCsvButton.addEventListener("click", () => exportCsv(filteredRows));
  exportXlsxButton.addEventListener("click", () => exportXlsx(filteredRows));
}

if (doc) {
  doc.addEventListener("DOMContentLoaded", attachEvents);
}

function __setTestState(state) {
  if (state.headers) {
    headers = state.headers;
  }
  if (state.rawRows) {
    rawRows = state.rawRows;
  }
  if (state.filteredRows) {
    filteredRows = state.filteredRows;
  }
  if (state.rowTextCache) {
    rowTextCache = state.rowTextCache;
  }
  if (state.lowerRowTextCache) {
    lowerRowTextCache = state.lowerRowTextCache;
  }
  if (Array.isArray(state.tableColumns)) {
    tableColumns = state.tableColumns;
    columnKeyToIndex = new Map(tableColumns.map((column, index) => [column.key, index]));
  }
  if (typeof state.matchesColumnIndex === "number") {
    matchesColumnIndex = state.matchesColumnIndex;
  }
  if (typeof state.currentPage === "number") {
    currentPage = state.currentPage;
  }
  if (typeof state.currentFileName === "string") {
    currentFileName = state.currentFileName;
  }
}

function __getTestState() {
  return {
    headers,
    rawRows,
    filteredRows,
    rowTextCache,
    lowerRowTextCache,
    tableColumns,
    matchesColumnIndex,
    currentPage,
    currentFileName,
  };
}

/**
 * Checklist tests manuels — Import, comparaison & filtrage par colonne
 * - [x] CSV (, ; \t |) avec et sans BOM importés en mode Analyse (valeurs correctes, recherche fonctionnelle).
 * - [x] XLSX multi-onglets avec en-têtes manquants → première feuille lue, colonnes renommées Colonne n.
 * - [x] Fichier > 50 Mo → message "Fichier trop volumineux, essayez de le scinder." et aucun blocage.
 * - [x] Drag & drop : CSV puis XLSX enchaînés, états réinitialisés proprement entre les imports.
 * - [x] Bouton Réinitialiser : annule les lectures en cours, vide inputs, colonnes sélectionnées et tableau.
 * - [x] Analyse : recherche multi-mots, options casse/exact match inchangées avec toutes colonnes par défaut.
 * - [x] Comparaison : colonnes ref./cmp. visibles, cases en tête opérationnelles, colonne "Mots-clés trouvés" cohérente.
 * - [x] Export CSV/XLSX et copie → données filtrées uniquement (colonne "Mots-clés trouvés" incluse, sans <mark>).
 * - [x] Aucun résultat → message aria-live et ligne "Aucune ligne correspondante" rendue.
 * - [x] Dataset ~10k lignes → UI fluide (debounce 300 ms, parsing CSV par batch setTimeout).
 */

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    tokenizeQuery,
    toPostfix,
    evaluateQuery,
    matchRow,
    convertRowsToCsv,
    buildCaches,
    normalizeParsedData,
    extractKeywords,
    getColumnsFor,
    getSingleFileColumns,
    __setTestState,
    __getTestState,
  };
}
