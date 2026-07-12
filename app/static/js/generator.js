const form = document.getElementById("generator-form");
const submitBtn = document.getElementById("submit-btn");
const statusEl = document.getElementById("status");
const resultsCard = document.getElementById("results-card");
const summaryEl = document.getElementById("summary");
const downloadLink = document.getElementById("download-link");

const excelInput = document.getElementById("excel-file");
const proformaInput = document.getElementById("proforma-pdf");
const jobOrderInput = document.getElementById("job-order-pdf");
const soeFilesInput = document.getElementById("soe-pdf-files");
const soeFolderInput = document.getElementById("soe-pdf-folder");
const jobOrderSourceInput = document.getElementById("job-order-source");

const excelFileNameEl = document.getElementById("excel-file-name");
const proformaFileNameEl = document.getElementById("proforma-file-name");
const jobOrderFileNameEl = document.getElementById("job-order-file-name");
const soeFileList = document.getElementById("soe-file-list");

let selectedSoeFiles = [];

function setStatus(message, type = "info") {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  statusEl.classList.remove("hidden");
}

function clearStatus() {
  statusEl.classList.add("hidden");
}

function formatMoney(value) {
  return Number(value).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
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

function renderSoeFiles() {
  if (!selectedSoeFiles.length) {
    soeFileList.classList.add("hidden");
    soeFileList.innerHTML = "";
    return;
  }

  soeFileList.innerHTML = selectedSoeFiles
    .map((file) => `<li>${displayName(file)}</li>`)
    .join("");
  soeFileList.classList.remove("hidden");
}

function setSelectedSoeFiles(files, sourceInput) {
  selectedSoeFiles = collectPdfFiles(files);
  if (sourceInput === soeFilesInput) {
    soeFolderInput.value = "";
  } else {
    soeFilesInput.value = "";
  }
  renderSoeFiles();
}

function formatSource(source) {
  if (source === "operational_time_summary") {
    return "Operational Time Summary";
  }
  if (source === "time_log") {
    return "Time Log";
  }
  if (source === "1c") {
    return "1-C procedure";
  }
  if (source === "running") {
    return "Running completion";
  }
  if (source === "auto") {
    return "Auto-detect";
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

function hideResultBlocks() {
  document.getElementById("proforma-results").classList.add("hidden");
  document.getElementById("soe-results").classList.add("hidden");
  document.getElementById("job-order-results").classList.add("hidden");
}

function renderResults(data) {
  hideResultBlocks();

  summaryEl.innerHTML = `
    <div class="summary-item">
      <span>Sections</span>
      <strong>${data.processed_sections.join(", ")}</strong>
    </div>
    <div class="summary-item">
      <span>Output file</span>
      <strong>${data.filename}</strong>
    </div>
  `;

  if (data.proforma) {
    const block = document.getElementById("proforma-results");
    const tbody = document.querySelector("#proforma-table tbody");
    summaryEl.innerHTML += `
      <div class="summary-item">
        <span>Proforma items</span>
        <strong>${data.proforma.item_count}</strong>
      </div>
      <div class="summary-item">
        <span>Proforma total</span>
        <strong>$${formatMoney(data.proforma.gross_total)}</strong>
      </div>
    `;
    tbody.innerHTML = data.proforma.items
      .map(
        (item) => `
          <tr>
            <td class="num">${item.sno}</td>
            <td class="desc">${item.description.replace(/\n/g, "<br>")}</td>
            <td class="num">${formatMoney(item.per_day_rate)}</td>
            <td class="num">${item.days}</td>
            <td class="num">${formatMoney(item.total)}</td>
          </tr>
        `
      )
      .join("");
    block.classList.remove("hidden");
  }

  if (data.soe) {
    const block = document.getElementById("soe-results");
    summaryEl.innerHTML += `
      <div class="summary-item">
        <span>SOE PDFs</span>
        <strong>${data.soe.pdf_count}</strong>
      </div>
      <div class="summary-item">
        <span>SOE rows</span>
        <strong>${data.soe.row_count}</strong>
      </div>
    `;
    document.querySelector("#soe-summary-table tbody").innerHTML = data.soe.pdf_summaries
      .map(
        (summary) => `
          <tr>
            <td>${summary.filename}</td>
            <td>${formatSource(summary.source)}</td>
            <td>${formatWellOrDate(summary)}</td>
            <td>${summary.rig || "-"}</td>
            <td>${formatPeriod(summary)}</td>
            <td class="num">${
              summary.skipped
                ? summary.skip_reason === "rig_mismatch"
                  ? "Skipped (rig)"
                  : "Skipped"
                : summary.row_count
            }</td>
          </tr>
        `
      )
      .join("");
    document.querySelector("#soe-rows-table tbody").innerHTML = data.soe.rows
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
    block.classList.remove("hidden");
  }

  if (data.job_order) {
    const block = document.getElementById("job-order-results");
    summaryEl.innerHTML += `
      <div class="summary-item">
        <span>Job Order rows</span>
        <strong>${data.job_order.line_count}</strong>
      </div>
    `;
    document.querySelector("#job-order-table tbody").innerHTML = data.job_order.lines
      .map(
        (line) => `
          <tr>
            <td class="num">${line.line_no || ""}</td>
            <td class="desc">${line.text.replace(/\n/g, "<br>")}</td>
          </tr>
        `
      )
      .join("");
    block.classList.remove("hidden");
  }

  downloadLink.href = data.download_url;
  downloadLink.download = data.filename;
  resultsCard.classList.remove("hidden");
}

excelInput.addEventListener("change", () => {
  showSelectedFile(excelFileNameEl, excelInput.files[0]);
});

proformaInput.addEventListener("change", () => {
  showSelectedFile(proformaFileNameEl, proformaInput.files[0]);
});

jobOrderInput.addEventListener("change", () => {
  showSelectedFile(jobOrderFileNameEl, jobOrderInput.files[0]);
});

soeFilesInput.addEventListener("change", () => {
  setSelectedSoeFiles(soeFilesInput.files, soeFilesInput);
});

soeFolderInput.addEventListener("change", () => {
  const pdfFiles = collectPdfFiles(soeFolderInput.files);
  if (soeFolderInput.files?.length && !pdfFiles.length) {
    selectedSoeFiles = [];
    soeFolderInput.value = "";
    renderSoeFiles();
    setStatus("The selected folder does not contain any PDF files.", "error");
    return;
  }
  setSelectedSoeFiles(soeFolderInput.files, soeFolderInput);
  clearStatus();
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  clearStatus();

  const excelFile = excelInput.files[0];
  const proformaFile = proformaInput.files[0];
  const jobOrderFile = jobOrderInput.files[0];

  if (!excelFile) {
    setStatus("Excel template is required.", "error");
    return;
  }

  if (!proformaFile && !jobOrderFile && !selectedSoeFiles.length) {
    setStatus("Select at least one PDF section: Proforma, SOE, or Job Order.", "error");
    return;
  }

  const formData = new FormData();
  formData.append("excel", excelFile);

  if (proformaFile) {
    formData.append("proforma_pdf", proformaFile);
  }
  if (jobOrderFile) {
    formData.append("job_order_pdf", jobOrderFile);
    formData.append("job_order_source", jobOrderSourceInput.value);
  }
  selectedSoeFiles.forEach((file) => {
    formData.append("soe_pdfs", file);
    formData.append("soe_pdf_names", displayName(file));
  });

  submitBtn.disabled = true;
  submitBtn.textContent = "Processing...";
  setStatus("Extracting PDFs and writing workbook...", "info");

  try {
    const response = await fetch("/api/generate", {
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
    setStatus("Workbook generated successfully.", "info");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = "Generate Workbook";
  }
});
