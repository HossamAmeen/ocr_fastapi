const form = document.getElementById("soe-form");
const submitBtn = document.getElementById("submit-btn");
const statusEl = document.getElementById("status");
const resultsCard = document.getElementById("results-card");
const summaryEl = document.getElementById("summary");
const pdfSummaryBody = document.querySelector("#pdf-summary-table tbody");
const rowsTableBody = document.querySelector("#rows-table tbody");
const downloadLink = document.getElementById("download-link");
const pdfFilesInput = document.getElementById("pdf-files");
const pdfFolderInput = document.getElementById("pdf-folder");
const pdfFileList = document.getElementById("pdf-file-list");
const tableNamesInput = document.getElementById("table-names");
const tableNamesLabel = document.getElementById("table-names-label");
const tableNamesHint = document.getElementById("table-names-hint");
const modeTableInput = document.getElementById("mode-table");
const modeSummaryInput = document.getElementById("mode-summary");

const excelInput = document.getElementById("excel-file");
const excelFileNameEl = document.getElementById("excel-file-name");

let selectedPdfFiles = [];

function setStatus(message, type = "info") {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  statusEl.classList.remove("hidden");
}

function clearStatus() {
  statusEl.classList.add("hidden");
}

function displayName(file) {
  return file.webkitRelativePath || file.name;
}

function showSelectedFile(element, file) {
  if (!file) {
    element.textContent = "";
    element.classList.add("hidden");
    return;
  }
  element.textContent = file.name;
  element.classList.remove("hidden");
}

if (excelInput) {
  excelInput.addEventListener("change", () => {
    showSelectedFile(excelFileNameEl, excelInput.files[0]);
  });
}

function collectPdfFiles(fileList) {
  return Array.from(fileList || [])
    .filter((file) => file.name.toLowerCase().endsWith(".pdf"))
    .sort((left, right) =>
      displayName(left).localeCompare(displayName(right), undefined, {
        numeric: true,
        sensitivity: "base",
      })
    );
}

function renderSelectedFiles() {
  if (!selectedPdfFiles.length) {
    pdfFileList.classList.add("hidden");
    pdfFileList.innerHTML = "";
    return;
  }

  pdfFileList.innerHTML = selectedPdfFiles
    .map(
      (file, index) => `
        <li class="file-list-item">
          <span class="file-list-name">${displayName(file)}</span>
          <button
            type="button"
            class="file-list-remove"
            data-index="${index}"
            aria-label="Remove ${displayName(file)}"
            title="Remove file"
          >&times;</button>
        </li>
      `
    )
    .join("");
  pdfFileList.classList.remove("hidden");
}

function removeSelectedPdfFile(index) {
  if (index < 0 || index >= selectedPdfFiles.length) {
    return;
  }
  selectedPdfFiles.splice(index, 1);
  if (!selectedPdfFiles.length) {
    pdfFilesInput.value = "";
    pdfFolderInput.value = "";
  }
  renderSelectedFiles();
}

function setSelectedPdfFiles(files, sourceInput) {
  selectedPdfFiles = collectPdfFiles(files);

  if (sourceInput === pdfFilesInput) {
    pdfFolderInput.value = "";
  } else {
    pdfFilesInput.value = "";
  }

  renderSelectedFiles();
}

function formatSource(source) {
  if (source === "operational_time_summary") {
    return "Operational Time Summary";
  }
  if (source === "time_log") {
    return "Time Log";
  }
  if (source === "job_time_log") {
    return "Job Time Log";
  }
  if (source === "paragraph_summary") {
    return "Summary (paragraph)";
  }
  return source || "-";
}

function formatWellOrDate(summary) {
  if (summary.report_date) {
    return summary.report_date;
  }
  return summary.well_name || "-";
}

function formatSkipReason(summary) {
  if (!summary.skipped) {
    return summary.row_count;
  }
  if (summary.skip_reason === "rig_mismatch") {
    return "Skipped (rig mismatch)";
  }
  if (summary.skip_reason === "no_matching_table") {
    return "Skipped (no matching table)";
  }
  if (summary.skip_reason === "empty_table") {
    return "Skipped (empty table)";
  }
  return "Skipped";
}

function formatPeriod(summary) {
  if (summary.report_period_from) {
    return `${summary.report_period_from}${summary.report_period_to ? ` to ${summary.report_period_to}` : ""}`;
  }
  return "-";
}

function parseTableNames(rawValue) {
  return rawValue
    .split(/[\n,]+/)
    .map((name) => name.trim())
    .filter(Boolean);
}

function isSummaryMode() {
  return Boolean(modeSummaryInput && modeSummaryInput.checked);
}

function updateModeUi() {
  if (!tableNamesLabel || !tableNamesHint) {
    return;
  }
  if (isSummaryMode()) {
    tableNamesLabel.textContent = "Paragraph title (used as the row's title)";
    tableNamesHint.textContent =
      "The exact paragraph title as it appears in the PDF. Its content is captured as one Excel row.";
    if (
      tableNamesInput &&
      /Time Log|Job Time Log|Operational Time Summary/.test(tableNamesInput.value)
    ) {
      tableNamesInput.value = "";
      tableNamesInput.placeholder = "24 Hrs Summary";
    }
  } else {
    tableNamesLabel.textContent = "Table names to extract";
    tableNamesHint.textContent =
      "One table name per line (or comma-separated). Only matching PDF tables are extracted.";
    if (tableNamesInput && !tableNamesInput.value.trim()) {
      tableNamesInput.placeholder = "Time Log\nJob Time Log\nOperational Time Summary";
    }
  }
}

if (modeTableInput && modeSummaryInput) {
  modeTableInput.addEventListener("change", updateModeUi);
  modeSummaryInput.addEventListener("change", updateModeUi);
  updateModeUi();
}

function renderResults(data) {
  summaryEl.innerHTML = `
    <div class="summary-item">
      <span>PDFs</span>
      <strong>${data.pdf_count}</strong>
    </div>
    <div class="summary-item">
      <span>Rows appended</span>
      <strong>${data.row_count}</strong>
    </div>
    ${data.rig_filter ? `
    <div class="summary-item">
      <span>Rig filter (from Excel)</span>
      <strong>${data.rig_filter}</strong>
    </div>` : ""}
    <div class="summary-item">
      <span>Output file</span>
      <strong>${data.filename}</strong>
    </div>
  `;

  pdfSummaryBody.innerHTML = data.pdf_summaries
    .map(
      (summary) => `
        <tr>
          <td>${summary.filename}</td>
          <td>${formatSource(summary.source)}</td>
          <td>${formatWellOrDate(summary)}</td>
          <td>${summary.rig || "-"}</td>
          <td>${formatPeriod(summary)}</td>
          <td class="num">${formatSkipReason(summary)}</td>
        </tr>
      `
    )
    .join("");

  rowsTableBody.innerHTML = data.rows
    .map(
      (row) => `
        <tr>
          <td class="num">${row.date}</td>
          <td class="num">${row.time}</td>
          <td class="desc">${row.event.replace(/\n/g, "<br>")}</td>
        </tr>
      `
    )
    .join("");

  downloadLink.href = data.download_url;
  downloadLink.download = data.filename;
  resultsCard.classList.remove("hidden");
}

pdfFilesInput.addEventListener("change", () => {
  setSelectedPdfFiles(pdfFilesInput.files, pdfFilesInput);
});

pdfFolderInput.addEventListener("change", () => {
  const pdfFiles = collectPdfFiles(pdfFolderInput.files);
  if (pdfFolderInput.files?.length && !pdfFiles.length) {
    selectedPdfFiles = [];
    pdfFolderInput.value = "";
    renderSelectedFiles();
    setStatus("The selected folder does not contain any PDF files.", "error");
    return;
  }
  setSelectedPdfFiles(pdfFolderInput.files, pdfFolderInput);
  clearStatus();
});

if (pdfFileList) {
  pdfFileList.addEventListener("click", (event) => {
    const button = event.target.closest(".file-list-remove");
    if (!button) {
      return;
    }
    removeSelectedPdfFile(Number(button.dataset.index));
  });
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  clearStatus();

  const excelFile = excelInput ? excelInput.files[0] : null;

  if (!selectedPdfFiles.length || !excelFile) {
    setStatus("Please select at least one PDF (or a folder of PDFs) and an Excel template.", "error");
    return;
  }

  const formData = new FormData();
  selectedPdfFiles.forEach((file) => {
    formData.append("pdfs", file);
    formData.append("pdf_names", displayName(file));
  });
  parseTableNames(tableNamesInput.value).forEach((name) => {
    formData.append("table_names", name);
  });
  formData.append("is_summary", isSummaryMode() ? "true" : "false");
  formData.append("extraction_mode", isSummaryMode() ? "summary" : "table");
  formData.append("excel", excelFile);

  submitBtn.disabled = true;
  submitBtn.textContent = "Processing...";
  setStatus(`Extracting ${selectedPdfFiles.length} PDF(s) and writing Excel...`, "info");

  try {
    const response = await fetch("/api/soe/generate", {
      method: "POST",
      body: formData,
    });

    const payload = await response.json();

    if (!response.ok) {
      const detail = payload.detail;
      const message = Array.isArray(detail)
        ? detail.map((entry) => entry.msg || JSON.stringify(entry)).join(", ")
        : detail || "Request failed.";
      throw new Error(message);
    }

    renderResults(payload);
    setStatus("SOE workbook generated successfully.", "info");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = "Generate SOE Workbook";
  }
});
