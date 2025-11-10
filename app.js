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

const PAGE_SIZE = 100;
const MAX_FILE_SIZE = 200 * 1024 * 1024; // 200MB

let headers = [];
let rawRows = [];
let filteredRows = [];
let rowTextCache = [];
let lowerRowTextCache = [];
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
}

function showError(message) {
  if (errorMessage) {
    errorMessage.textContent = message;
  }
}

function clearError() {
  if (errorMessage) {
    errorMessage.textContent = "";
  }
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
  currentFileName = fileName || "";
  controlsSection.hidden = false;
  resultsSection.hidden = false;
  renderPage(1);
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
  if (searchInput.value.trim()) {
    performSearch();
  }
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
  rowTextCache = rawRows.map((row) =>
    row
      .map((value) => (value === null || value === undefined ? "" : String(value)))
      .join(" \u2022 ")
  );
  lowerRowTextCache = rowTextCache.map((text) => text.toLowerCase());
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

function matchRow(rowIndex, keyword, { caseSensitive, exactMatch }) {
  const source = caseSensitive ? rowTextCache[rowIndex] : lowerRowTextCache[rowIndex];
  const query = caseSensitive ? keyword : keyword.toLowerCase();
  if (!query) return false;

  if (exactMatch) {
    const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const regex = new RegExp(`(^|\\b)${escaped}(\\b|$)`);
    return regex.test(source);
  }

  return source.includes(query);
}

function evaluateQuery(tokens, options) {
  if (!tokens.length) {
    return rawRows.map((_, index) => index);
  }

  const postfix = toPostfix(tokens);
  const matches = [];

  for (let index = 0; index < rawRows.length; index += 1) {
    const stack = [];
    for (const token of postfix) {
      if (token.type === "operand") {
        stack.push(matchRow(index, token.value, options));
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
  const options = {
    caseSensitive: caseSensitiveToggle.checked,
    exactMatch: exactMatchToggle.checked,
  };

  if (!query) {
    filteredRows = [...rawRows];
    renderPage(1);
    return;
  }

  try {
    const tokens = tokenizeQuery(query);
    const indexes = evaluateQuery(tokens, options);
    filteredRows = indexes.map((i) => rawRows[i]);
    renderPage(1);
  } catch (error) {
    console.error(error);
    showError(error.message || "Requête invalide");
  }
}

function resetSearch() {
  searchInput.value = "";
  caseSensitiveToggle.checked = false;
  exactMatchToggle.checked = false;
  if (currentMode === "compare" && comparisonRows.length) {
    updateComparisonDataset();
  } else {
    filteredRows = [...rawRows];
    renderPage(1);
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
    showError("");
    updateProgress(100, "Résultats copiés dans le presse-papiers");
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
}

function exportXlsx(rows) {
  if (!rows.length) return;
  const worksheetData = [headers, ...rows];
  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Résultats");
  const filename = `${currentFileName || "export"}_resultats.xlsx`;
  XLSX.writeFile(workbook, filename, { compression: true });
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

  searchButton.addEventListener("click", performSearch);
  searchInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      performSearch();
    }
  });

  resetButton.addEventListener("click", () => {
    resetSearch();
    clearError();
  });

  if (caseSensitiveToggle && typeof caseSensitiveToggle.addEventListener === "function") {
    caseSensitiveToggle.addEventListener("change", () => {
      if (currentMode === "compare" && comparisonRows.length) {
        updateComparisonDataset();
      }
    });
  }

  if (exactMatchToggle && typeof exactMatchToggle.addEventListener === "function") {
    exactMatchToggle.addEventListener("change", () => {
      if (currentMode === "compare" && comparisonRows.length) {
        updateComparisonDataset();
      }
    });
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
    currentPage,
    currentFileName,
  };
}

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
    __setTestState,
    __getTestState,
  };
}
