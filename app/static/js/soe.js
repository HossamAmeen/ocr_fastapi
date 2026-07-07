const form = document.getElementById("soe-form");
const submitBtn = document.getElementById("submit-btn");
const statusEl = document.getElementById("status");
const resultsCard = document.getElementById("results-card");
const summaryEl = document.getElementById("summary");
const pdfSummaryBody = document.querySelector("#pdf-summary-table tbody");
const rowsTableBody = document.querySelector("#rows-table tbody");
const downloadLink = document.getElementById("download-link");
const pdfFilesInput = document.getElementById("pdf-files");
const pdfFileList = document.getElementById("pdf-file-list");

function setStatus(message, type = "info") {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  statusEl.classList.remove("hidden");
}

function clearStatus() {
  statusEl.classList.add("hidden");
}

function renderSelectedFiles() {
  const files = Array.from(pdfFilesInput.files || []);
  if (!files.length) {
    pdfFileList.classList.add("hidden");
    pdfFileList.innerHTML = "";
    return;
  }

  pdfFileList.innerHTML = files.map((file) => `<li>${file.name}</li>`).join("");
  pdfFileList.classList.remove("hidden");
}

function formatSource(source) {
  if (source === "operational_time_summary") {
    return "Operational Time Summary";
  }
  if (source === "time_log") {
    return "Time Log";
  }
  return source || "-";
}

function formatWellOrDate(summary) {
  if (summary.report_date) {
    return summary.report_date;
  }
  return summary.well_name || "-";
}

function formatPeriod(summary) {
  if (summary.report_period_from) {
    return `${summary.report_period_from}${summary.report_period_to ? ` to ${summary.report_period_to}` : ""}`;
  }
  return "-";
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
          <td>${formatPeriod(summary)}</td>
          <td class="num">${summary.skipped ? "Skipped" : summary.row_count}</td>
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

pdfFilesInput.addEventListener("change", renderSelectedFiles);

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  clearStatus();

  const pdfFiles = Array.from(pdfFilesInput.files || []);
  const excelFile = document.getElementById("excel-file").files[0];

  if (!pdfFiles.length || !excelFile) {
    setStatus("Please select at least one PDF and an Excel template.", "error");
    return;
  }

  const formData = new FormData();
  pdfFiles.forEach((file) => formData.append("pdfs", file));
  formData.append("excel", excelFile);

  submitBtn.disabled = true;
  submitBtn.textContent = "Processing...";
  setStatus(`Extracting ${pdfFiles.length} PDF(s) and writing Excel...`, "info");

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
