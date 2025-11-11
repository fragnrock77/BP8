/**
 * Rapport de modifications — Filtrage par colonne
 * - Bugs corrigés : cache de recherche concaténé (faux positifs inter-colonnes) remplacé par un cache par cellule et annonces d'erreurs via aria-live pour éviter les silences en cas d'échec.
 * - Changements UI : ajout d'un filtre multi-colonnes persistant avec bouton accessible et menu déroulant.
 * - Implémentation : selectedColumns dans l'état global, getAvailableColumns, adaptation de matchRow/performSearch.
 * - Accessibilité : fieldset/legend, aria-live dédié, focus visible dans le menu, annonces vocales des limites de recherche.
 * - Performance : debounce 300 ms sur recherche/filtre, réduction des colonnes inspectées, worker conservé.
 * - Tests manuels : checklist en fin de fichier (cas "Toutes", une colonne, multi, export, 10k lignes, options casse/exacte).
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
const columnFilterToggle = doc ? doc.getElementById("column-filter-toggle") : null;
const columnFilterMenu = doc ? doc.getElementById("column-filter-menu") : null;
const columnFilterOptionsContainer = doc
  ? doc.getElementById("column-filter-options")
  : null;
const columnFilterAll = doc ? doc.getElementById("column-filter-all") : null;
const columnFilterLabel = doc
  ? doc.getElementById("column-filter-label")
  : { textContent: "" };
const columnFilterLive = doc ? doc.getElementById("column-filter-live") : null;
const liveFeedback = doc ? doc.getElementById("live-feedback") : null;

const PAGE_SIZE = 100;
const SEARCH_DEBOUNCE = 300;
const MAX_FILE_SIZE = 200 * 1024 * 1024; // 200MB

const STORAGE_KEYS = {
  selectedColumns: "bp8.search.selectedColumns",
};

function loadSelectedColumns() {
  if (typeof window === "undefined") {
    return null;
  }
  try {
    const stored = window.localStorage.getItem(STORAGE_KEYS.selectedColumns);
    if (!stored) {
      return null;
    }
    const parsed = JSON.parse(stored);
    if (!Array.isArray(parsed) || !parsed.length) {
      return null;
    }
    return parsed.map((key) => String(key));
  } catch (error) {
    console.warn("Impossible de charger les colonnes persistées", error);
    return null;
  }
}

function saveSelectedColumns(columns) {
  if (typeof window === "undefined") {
    return;
  }
  try {
    if (!columns || !columns.length) {
      window.localStorage.removeItem(STORAGE_KEYS.selectedColumns);
      return;
    }
    window.localStorage.setItem(STORAGE_KEYS.selectedColumns, JSON.stringify(columns));
  } catch (error) {
    console.warn("Impossible de persister les colonnes sélectionnées", error);
  }
}

const state = {
  search: {
    query: "",
    caseSensitive: false,
    exactMatch: false,
    selectedColumns: loadSelectedColumns(),
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
let availableColumns = [];
let currentPage = 1;
let currentFileName = "";
let currentMode = "single";
let referenceKeywords = [];
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
  availableColumns = [];
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
  rebuildColumnFilterOptions();
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

function syncSelectedColumnsWithAvailable() {
  if (!state.search.selectedColumns || !state.search.selectedColumns.length) {
    state.search.selectedColumns = null;
    return;
  }
  const availableKeys = new Set(availableColumns.map((column) => column.key));
  const filtered = state.search.selectedColumns.filter((key) => availableKeys.has(key));
  state.search.selectedColumns = filtered.length ? filtered : null;
}

function closeColumnFilterMenu() {
  if (!columnFilterMenu || !columnFilterToggle) {
    return;
  }
  columnFilterMenu.hidden = true;
  columnFilterToggle.setAttribute("aria-expanded", "false");
}

function openColumnFilterMenu() {
  if (
    !columnFilterMenu ||
    !columnFilterToggle ||
    (columnFilterFieldset && columnFilterFieldset.disabled)
  ) {
    return;
  }
  columnFilterMenu.hidden = false;
  columnFilterToggle.setAttribute("aria-expanded", "true");
  const firstInput = columnFilterMenu.querySelector('input[type="checkbox"]');
  if (firstInput) {
    firstInput.focus({ preventScroll: true });
  }
}

function toggleColumnFilterMenu(force) {
  if (!columnFilterMenu || !columnFilterToggle) {
    return;
  }
  const isOpen = !columnFilterMenu.hidden;
  if (force === true) {
    if (!isOpen) {
      openColumnFilterMenu();
    }
    return;
  }
  if (force === false) {
    if (isOpen) {
      closeColumnFilterMenu();
    }
    return;
  }
  if (isOpen) {
    closeColumnFilterMenu();
  } else {
    openColumnFilterMenu();
  }
}

function setColumnFilterDisabled(disabled) {
  if (!columnFilterFieldset) {
    return;
  }
  columnFilterFieldset.disabled = Boolean(disabled);
  if (columnFilterToggle) {
    columnFilterToggle.setAttribute("aria-disabled", String(Boolean(disabled)));
  }
  if (disabled) {
    closeColumnFilterMenu();
  }
}

function updateColumnFilterLabel() {
  if (!columnFilterLabel) {
    return;
  }
  if (!availableColumns.length) {
    columnFilterLabel.textContent = "Colonnes : Toutes";
    return;
  }
  if (!state.search.selectedColumns || !state.search.selectedColumns.length) {
    columnFilterLabel.textContent = "Colonnes : Toutes";
    return;
  }
  const selectedLabels = availableColumns
    .filter((column) => state.search.selectedColumns.includes(column.key))
    .map((column) => column.label);
  if (!selectedLabels.length) {
    columnFilterLabel.textContent = "Colonnes : Toutes";
    return;
  }
  const preview = selectedLabels.slice(0, 3).join(", ");
  const remaining = selectedLabels.length > 3 ? ` +${selectedLabels.length - 3}` : "";
  columnFilterLabel.textContent = `Colonnes : ${preview}${remaining}`;
}

function announceColumnSelection() {
  if (!columnFilterLive) {
    return;
  }
  if (!availableColumns.length) {
    columnFilterLive.textContent = "Recherche sur toutes les colonnes.";
    return;
  }
  if (!state.search.selectedColumns || !state.search.selectedColumns.length) {
    columnFilterLive.textContent = "Recherche sur toutes les colonnes.";
    return;
  }
  const count = state.search.selectedColumns.length;
  columnFilterLive.textContent = `Recherche limitée à ${count} colonne${count > 1 ? "s" : ""}.`;
}

function rebuildColumnFilterOptions() {
  if (!doc || !columnFilterOptionsContainer || !columnFilterAll) {
    return;
  }
  closeColumnFilterMenu();
  columnFilterOptionsContainer.innerHTML = "";
  if (!availableColumns.length) {
    columnFilterAll.checked = true;
    setColumnFilterDisabled(true);
    updateColumnFilterLabel();
    announceColumnSelection();
    return;
  }

  setColumnFilterDisabled(false);
  const fragment = doc.createDocumentFragment();
  availableColumns.forEach((column) => {
    const optionLabel = doc.createElement("label");
    optionLabel.className = "checkbox column-filter__option";
    const input = doc.createElement("input");
    input.type = "checkbox";
    input.dataset.columnKey = column.key;
    input.id = `column-filter-${column.key}`;
    input.checked = Boolean(
      state.search.selectedColumns && state.search.selectedColumns.includes(column.key)
    );
    optionLabel.appendChild(input);
    optionLabel.appendChild(doc.createTextNode(column.label));
    fragment.appendChild(optionLabel);
  });

  columnFilterOptionsContainer.appendChild(fragment);
  const hasSelection = Boolean(state.search.selectedColumns && state.search.selectedColumns.length);
  columnFilterAll.checked = !hasSelection;
  updateColumnFilterLabel();
  announceColumnSelection();
}

function updateAvailableColumns() {
  availableColumns = getAvailableColumns(rawRows);
  syncSelectedColumnsWithAvailable();
  if (!state.search.selectedColumns || !state.search.selectedColumns.length) {
    saveSelectedColumns(null);
  } else {
    saveSelectedColumns(state.search.selectedColumns);
  }
  rebuildColumnFilterOptions();
}

function getDefaultColumnIndexes() {
  if (availableColumns.length) {
    return availableColumns.map((column) => Number(column.key));
  }
  const candidateLength = Math.max(headers.length, rawRows[0]?.length || 0);
  if (!candidateLength) {
    return [];
  }
  return Array.from({ length: candidateLength }, (_, index) => index);
}

function getColumnIndexesForSearch() {
  if (!state.search.selectedColumns || !state.search.selectedColumns.length) {
    return getDefaultColumnIndexes();
  }
  const availableMap = new Map(
    availableColumns.map((column) => [column.key, Number(column.key)])
  );
  const indexes = state.search.selectedColumns
    .map((key) => availableMap.get(key))
    .filter((value) => typeof value === "number" && Number.isFinite(value));
  return indexes.length ? indexes : getDefaultColumnIndexes();
}

function handleColumnFilterAllChange() {
  if (!columnFilterAll) {
    return;
  }
  if (!columnFilterAll.checked) {
    if (!state.search.selectedColumns || !state.search.selectedColumns.length) {
      columnFilterAll.checked = true;
    }
    return;
  }
  state.search.selectedColumns = null;
  saveSelectedColumns(null);
  if (columnFilterOptionsContainer) {
    const inputs = columnFilterOptionsContainer.querySelectorAll('input[type="checkbox"]');
    inputs.forEach((input) => {
      input.checked = false;
    });
  }
  updateColumnFilterLabel();
  announceColumnSelection();
  scheduleSearch();
}

function handleColumnFilterOptionChange(event) {
  const target = event.target;
  if (!target || target.tagName !== "INPUT" || target.type !== "checkbox") {
    return;
  }
  const { columnKey } = target.dataset;
  if (!columnKey) {
    return;
  }
  const selection = new Set(state.search.selectedColumns || []);
  if (target.checked) {
    selection.add(columnKey);
  } else {
    selection.delete(columnKey);
  }
  if (!selection.size) {
    state.search.selectedColumns = null;
    columnFilterAll.checked = true;
    saveSelectedColumns(null);
  } else {
    const ordered = availableColumns
      .map((column) => column.key)
      .filter((key) => selection.has(key));
    state.search.selectedColumns = ordered;
    columnFilterAll.checked = false;
    saveSelectedColumns(state.search.selectedColumns);
  }
  updateColumnFilterLabel();
  announceColumnSelection();
  scheduleSearch();
}

function handleDocumentClick(event) {
  if (!columnFilterFieldset || !columnFilterMenu || columnFilterMenu.hidden) {
    return;
  }
  if (columnFilterFieldset.contains(event.target)) {
    return;
  }
  closeColumnFilterMenu();
}

function handleDocumentKeydown(event) {
  if (event.key !== "Escape") {
    return;
  }
  if (!columnFilterMenu || columnFilterMenu.hidden) {
    return;
  }
  event.preventDefault();
  closeColumnFilterMenu();
  if (columnFilterToggle) {
    columnFilterToggle.focus({ preventScroll: true });
  }
}

function clearComparisonState() {
  referenceKeywords = [];
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

function applyDataset(newHeaders, newRows, fileName) {
  headers = Array.isArray(newHeaders) ? newHeaders.slice() : [];
  rawRows = Array.isArray(newRows) ? newRows.map((row) => row.slice()) : [];
  filteredRows = [...rawRows];
  buildCaches();
  updateAvailableColumns();
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
  if (!comparisonRows.length) {
    return;
  }

  const options = {
    caseSensitive: Boolean(caseSensitiveToggle.checked),
    exactMatch: Boolean(exactMatchToggle.checked),
  };
  state.search.caseSensitive = options.caseSensitive;
  state.search.exactMatch = options.exactMatch;

  const keywordCache = referenceKeywords.map((keyword) => ({
    original: keyword,
    normalized: options.caseSensitive ? keyword : keyword.toLowerCase(),
  }));

  const rowsWithMatches = comparisonRows.map((row) => {
    const haystack = options.caseSensitive
      ? row
      : row.map((value) => value.toLowerCase());
    const matches = [];
    keywordCache.forEach((keyword) => {
      if (!keyword.normalized) {
        return;
      }
      const found = haystack.some((cell) =>
        options.exactMatch ? cell === keyword.normalized : cell.includes(keyword.normalized)
      );
      if (found) {
        matches.push(keyword.original);
      }
    });
    const outputRow = row.slice();
    outputRow.push(matches.join(", "));
    return outputRow;
  });

  const headersWithMatches = [...comparisonHeaders, "Mots-clés trouvés"];
  const fileLabel = comparisonFileName ? `${comparisonFileName}_comparaison` : "comparaison";
  applyDataset(headersWithMatches, rowsWithMatches, fileLabel);
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

    applyDataset(parsedHeaders, rows, currentFileName);
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
    const { rows } = normalizeParsedData(parsed);
    if (!rows.length) {
      referenceKeywords = [];
      updateReferenceSummary();
      showError("Aucune donnée trouvée dans le fichier de référence.");
      resetDataset();
      return;
    }

    referenceKeywords = extractKeywords(rows);
    updateReferenceSummary();

    if (!comparisonRows.length) {
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
      comparisonRows = [];
      comparisonHeaders = parsedHeaders;
      resetDataset();
      showError("Aucune donnée trouvée dans le fichier à comparer.");
      return;
    }

    comparisonHeaders = parsedHeaders;
    comparisonRows = rows;

    if (!referenceKeywords.length) {
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

  if (!referenceKeywords.length) {
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
  rowTextCache = rawRows.map((row) => row.map((value) => sanitizeValue(value)));
  lowerRowTextCache = rowTextCache.map((cells) => cells.map((cell) => cell.toLowerCase()));
}

function renderTable(rows) {
  if (!dataTable) {
    return;
  }
  const columnCount = rows.reduce((max, row) => Math.max(max, row.length), headers.length);
  const effectiveColumnCount = columnCount || headers.length || (rows[0]?.length ?? 0);
  const tableHead = doc.createElement("thead");
  const headerRow = doc.createElement("tr");
  const totalColumns = Math.max(effectiveColumnCount, headers.length);
  for (let index = 0; index < totalColumns; index += 1) {
    const th = doc.createElement("th");
    const header = headers[index];
    th.textContent = header === undefined || header === null || header === ""
      ? `Colonne ${index + 1}`
      : String(header);
    headerRow.appendChild(th);
  }
  tableHead.appendChild(headerRow);

  const tableBody = doc.createElement("tbody");
  rows.forEach((row) => {
    const tr = doc.createElement("tr");
    for (let index = 0; index < totalColumns; index += 1) {
      const td = doc.createElement("td");
      const value = row[index];
      td.textContent = value === undefined || value === null ? "" : String(value);
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
    renderTable([]);
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

  renderTable(pageRows);

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

function performSearch() {
  clearError();
  const query = searchInput.value.trim();
  state.search.query = query;
  state.search.caseSensitive = Boolean(caseSensitiveToggle.checked);
  state.search.exactMatch = Boolean(exactMatchToggle.checked);
  const columnIndexes = getColumnIndexesForSearch();

  if (!query) {
    filteredRows = [...rawRows];
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
    filteredRows = indexes.map((i) => rawRows[i]);
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
    filteredRows = [...rawRows];
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

function convertRowsToCsv(headers, rows) {
  const allRows = [headers, ...rows];
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
  const worksheetData = [headers, ...rows];
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

  if (columnFilterToggle) {
    columnFilterToggle.addEventListener("click", () => {
      toggleColumnFilterMenu();
    });
    columnFilterToggle.addEventListener("keydown", (event) => {
      if (event.key === " " || event.key === "Enter" || event.key === "ArrowDown") {
        event.preventDefault();
        toggleColumnFilterMenu(true);
      }
      if (event.key === "Escape") {
        toggleColumnFilterMenu(false);
      }
    });
  }

  if (columnFilterAll) {
    columnFilterAll.addEventListener("change", handleColumnFilterAllChange);
  }

  if (columnFilterOptionsContainer) {
    columnFilterOptionsContainer.addEventListener("change", handleColumnFilterOptionChange);
  }

  if (doc) {
    doc.addEventListener("click", handleDocumentClick);
    doc.addEventListener("keydown", handleDocumentKeydown);
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
  availableColumns = Array.isArray(state.availableColumns) ? state.availableColumns : [];
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
    availableColumns,
    currentPage,
    currentFileName,
  };
}

/**
 * Checklist tests manuels — Filtrage par colonne
 * - [x] Aucune colonne cochée (ou "Toutes") : recherche identique au comportement initial.
 * - [x] Sélection d'une colonne unique : les résultats ne correspondent qu'aux valeurs de cette colonne.
 * - [x] Sélection multi-colonnes : union logique, un résultat suffit sur l'une des colonnes cochées.
 * - [x] Bascule "Toutes les colonnes" : activée par défaut, se décoche dès qu'une colonne est cochée puis se réactive lorsque tout est décoché.
 * - [x] Options "Sensible à la casse" et "Correspondance exacte" combinées avec le filtrage par colonne.
 * - [x] Exports CSV/XLSX et copie utilisent uniquement les lignes filtrées courantes.
 * - [x] Message aria-live explicite lorsqu'aucun résultat n'est trouvé et lors des limitations à N colonnes.
 * - [x] Dataset volumineux (~10k lignes) : navigation fluide grâce au debounce 300 ms et au worker existant.
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
    getAvailableColumns,
    __setTestState,
    __getTestState,
  };
}
