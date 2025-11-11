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
const fileInput = doc ? doc.getElementById("file-input") : null;
const modeRadios = doc
  ? Array.from(doc.querySelectorAll('input[name="analysis-mode"]'))
  : [];
const singleUploadContainer = doc ? doc.getElementById("single-upload") : null;
const compareUploadContainer = doc ? doc.getElementById("compare-upload") : null;
const referenceInput = doc ? doc.getElementById("reference-input") : null;
const comparisonInput = doc ? doc.getElementById("comparison-input") : null;
const referenceName = doc ? doc.getElementById("reference-name") : { textContent: "" };
const comparisonName = doc
  ? doc.getElementById("comparison-name")
  : { textContent: "" };
const keywordsSummary = doc ? doc.getElementById("keywords-summary") : { textContent: "" };
const progressBar = doc ? doc.getElementById("progress-bar") : null;
const progressLabel = doc ? doc.getElementById("progress-label") : null;
const errorMessage = doc ? doc.getElementById("error-message") : null;
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
const MAX_FILE_SIZE = 200 * 1024 * 1024; // 200MB

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

function resetDataset() {
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

function showError(message) {
  if (errorMessage) {
    errorMessage.textContent = message;
  }
  announceStatus(message);
}

function clearError() {
  if (errorMessage) {
    errorMessage.textContent = "";
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

function sanitizeRow(row) {
  if (!Array.isArray(row)) {
    return [];
  }
  return row.map((value) => sanitizeValue(value));
}

function normalizeParsedData(parsed) {
  const rawHeaders = Array.isArray(parsed.headers) ? parsed.headers : [];
  const rawRows = Array.isArray(parsed.rows) ? parsed.rows : [];
  const sanitizedRows = rawRows.map((row) => sanitizeRow(row));
  let sanitizedHeaders = rawHeaders.map((header) => sanitizeValue(header));

  if (!sanitizedHeaders.length && sanitizedRows.length) {
    sanitizedHeaders = sanitizedRows.shift() || [];
  }

  return { headers: sanitizedHeaders, rows: sanitizedRows };
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

function getFileExtension(file) {
  if (file.size > MAX_FILE_SIZE) {
    throw new Error(
      `Le fichier est trop volumineux (${formatBytes(file.size)}). Limite : ${formatBytes(
        MAX_FILE_SIZE
      )}.`
    );
  }
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (!extension || !["csv", "xlsx", "xls"].includes(extension)) {
    throw new Error("Format non supporté. Seuls les fichiers CSV ou XLSX sont acceptés.");
  }
  return extension;
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
  clearComparisonState();
  if (fileInput) {
    fileInput.value = "";
  }

  let extension;
  try {
    extension = getFileExtension(file);
  } catch (validationError) {
    showError(validationError.message);
    return;
  }

  currentFileName = file.name.replace(/\.[^.]+$/, "");
  updateProgress(0, "Préparation du fichier...");

  try {
    const parsed =
      extension === "csv" ? await parseCsv(file) : await parseXlsx(file);
    const { headers: parsedHeaders, rows } = normalizeParsedData(parsed);
    if (!rows.length) {
      showError("Aucune donnée trouvée dans le fichier.");
      resetDataset();
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
  } catch (error) {
    console.error(error);
    showError(
      "Impossible de lire le fichier. Vérifiez son encodage ou son intégrité et réessayez."
    );
    resetDataset();
  } finally {
    updateProgress(100, "Chargement terminé");
  }
}

async function handleReferenceFiles(files) {
  const [file] = files ? Array.from(files).filter(Boolean) : [];
  if (!file) return;
  clearError();
  if (referenceInput) {
    referenceInput.value = "";
  }

  let extension;
  try {
    extension = getFileExtension(file);
  } catch (validationError) {
    showError(validationError.message);
    return;
  }

  referenceName.textContent = file.name;
  referenceFileName = file.name.replace(/\.[^.]+$/, "");
  updateProgress(0, "Préparation du fichier de référence...");

  try {
    const parsed =
      extension === "csv" ? await parseCsv(file) : await parseXlsx(file);
    const { headers: parsedHeaders, rows } = normalizeParsedData(parsed);
    referenceHeaders = parsedHeaders;
    referenceRows = rows;
    if (!rows.length) {
      referenceKeywords = [];
      referenceHeaders = [];
      referenceRows = [];
      updateReferenceSummary();
      showError("Aucune donnée trouvée dans le fichier de référence.");
      resetDataset();
      return;
    }

    referenceKeywords = extractKeywords(rows);
    updateReferenceSummary();

    if (!comparisonHeaders.length && !comparisonRows.length) {
      resetDataset();
      updateProgress(100, "Fichier de référence chargé");
    } else {
      updateComparisonDataset();
    }
  } catch (error) {
    console.error(error);
    showError(
      "Impossible de lire le fichier de référence. Vérifiez son encodage ou son intégrité et réessayez."
    );
  }
}

async function handleComparisonFiles(files) {
  const [file] = files ? Array.from(files).filter(Boolean) : [];
  if (!file) return;
  clearError();
  if (comparisonInput) {
    comparisonInput.value = "";
  }

  let extension;
  try {
    extension = getFileExtension(file);
  } catch (validationError) {
    showError(validationError.message);
    return;
  }

  comparisonName.textContent = file.name;
  comparisonFileName = file.name.replace(/\.[^.]+$/, "");
  updateProgress(0, "Préparation du fichier à comparer...");

  try {
    const parsed =
      extension === "csv" ? await parseCsv(file) : await parseXlsx(file);
    const { headers: parsedHeaders, rows } = normalizeParsedData(parsed);
    if (!rows.length) {
      comparisonHeaders = [];
      comparisonRows = [];
      resetDataset();
      showError("Aucune donnée trouvée dans le fichier à comparer.");
      return;
    }

    comparisonHeaders = parsedHeaders;
    comparisonRows = rows;

    if (!referenceHeaders.length && !referenceRows.length) {
      resetDataset();
      updateProgress(100, "Fichier à comparer chargé");
    } else {
      updateComparisonDataset();
    }
  } catch (error) {
    console.error(error);
    showError(
      "Impossible de lire le fichier à comparer. Vérifiez son encodage ou son intégrité et réessayez."
    );
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

function parseCsv(file) {
  return new Promise((resolve, reject) => {
    const rows = [];
    let headerRow = null;
    let totalRows = 0;

    Papa.parse(file, {
      worker: true,
      skipEmptyLines: "greedy",
      chunkSize: 1024 * 1024,
      step: (results, parser) => {
        const { data, errors, meta } = results;
        if (errors.length) {
          parser.abort();
          reject(new Error(errors.map((err) => err.message).join("; ")));
          return;
        }

        if (!headerRow) {
          headerRow = data;
        } else {
          rows.push(data);
        }

        totalRows += 1;
        const percent = Math.min(99, Math.round((meta.cursor / file.size) * 100));
        updateProgress(percent, `${totalRows.toLocaleString()} lignes lues`);
      },
      complete: () => {
        resolve({ headers: headerRow, rows });
      },
      error: (error) => {
        reject(error);
      },
    });
  });
}

async function parseXlsx(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", dense: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("Le fichier ne contient pas de feuille exploitable.");
  }
  const sheet = workbook.Sheets[sheetName];
  const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const [headerRow, ...rows] = sheetData;
  return { headers: headerRow, rows };
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
  rows.forEach((row, rowIndex) => {
    const tr = doc.createElement("tr");
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
        const value = row?.[index];
        td.textContent = value === undefined || value === null ? "" : String(value);
      }
      tr.appendChild(td);
    }
    tableBody.appendChild(tr);
  });

  dataTable.innerHTML = "";
  dataTable.appendChild(tableHead);
  dataTable.appendChild(tableBody);
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
    resetSearch();
    clearError();
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
 * Checklist tests manuels — Comparaison & filtrage par colonne
 * - [x] Import ref + comparaison : colonnes ref./cmp. visibles avec cases en tête + contrôle "Tout / Rien".
 * - [x] Aucune case décochée → recherche équivalente à l'ancienne version (toutes les colonnes inspectées).
 * - [x] Colonne décochée → résultats actualisés, résumé "Colonnes" mis à jour et aria-live annonce la limite.
 * - [x] Sélection multi-colonnes → union logique des colonnes cochées, interactions casse/exact match respectées.
 * - [x] Colonne "Mots-clés trouvés" : liste des tokens détectés + colonnes associées, export/copier sans balises HTML.
 * - [x] Surbrillance <mark> dans les cellules correspondant aux mots-clés (respect options casse / correspondance exacte).
 * - [x] Aucun résultat → message aria-live + ligne "Aucune ligne correspondante" dans le tableau.
 * - [x] Dataset ~10k lignes → UI fluide (debounce 300 ms, worker conservé).
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
