const form = document.getElementById("job-order-form");
const submitBtn = document.getElementById("submit-btn");
const statusEl = document.getElementById("status");
const resultsCard = document.getElementById("results-card");
const summaryEl = document.getElementById("summary");
const linesTableBody = document.querySelector("#lines-table tbody");
const downloadLink = document.getElementById("download-link");
const sourceSelect = document.getElementById("source");
const templateHint = document.getElementById("template-hint");
const customMarkers = document.getElementById("custom-markers");
const startMarkerInput = document.getElementById("start-marker");
const endMarkerInput = document.getElementById("end-marker");

const TEMPLATE_HINTS = {
  auto:
    "Auto-detect scans the PDF and picks the first matching built-in template.",
  "1c":
    'Start: "1-C Upper Completion - Running Procedure". End: "1-D Additional Information" or "The final report requested...".',
  running:
    'Start: "Running Completion" (line start). End: "32 Perform final tests of TR-SCSSSV" or "4 SECURE WELL AND RELEASE RIG".',
  completion_procedure:
    'Start: "Completion Procedure". End: "CT OPERATION FOR MILLING GLASS DISC" (or end of PDF).',
  custom:
    "Enter the exact start phrase from your PDF below. Optionally add an end phrase.",
};

function setStatus(message, type = "info") {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  statusEl.classList.remove("hidden");
}

function clearStatus() {
  statusEl.classList.add("hidden");
}

function updateTemplateUi() {
  const source = sourceSelect.value;
  const isCustom = source === "custom";
  templateHint.textContent = TEMPLATE_HINTS[source] || "";
  customMarkers.classList.toggle("hidden", !isCustom);
  startMarkerInput.required = isCustom;
}

function formatSource(source) {
  if (source === "1c") {
    return "1-C procedure";
  }
  if (source === "running") {
    return "Running completion";
  }
  if (source === "completion_procedure") {
    return "Completion Procedure (SWI)";
  }
  if (source === "custom") {
    return "Custom start/end";
  }
  if (source === "auto") {
    return "Auto-detect";
  }
  return source || "-";
}

function renderResults(data) {
  summaryEl.innerHTML = `
    <div class="summary-item">
      <span>Section</span>
      <strong>${data.section_title || "-"}</strong>
    </div>
    <div class="summary-item">
      <span>Template used</span>
      <strong>${formatSource(data.source)}</strong>
    </div>
    <div class="summary-item">
      <span>Rows appended</span>
      <strong>${data.line_count}</strong>
    </div>
    <div class="summary-item">
      <span>Output file</span>
      <strong>${data.filename}</strong>
    </div>
  `;

  linesTableBody.innerHTML = data.lines
    .map(
      (line) => `
        <tr>
          <td class="num">${line.line_no || ""}</td>
          <td class="desc">${line.text.replace(/\n/g, "<br>")}</td>
        </tr>
      `
    )
    .join("");

  downloadLink.href = data.download_url;
  downloadLink.download = data.filename;
  resultsCard.classList.remove("hidden");
}

sourceSelect.addEventListener("change", updateTemplateUi);
updateTemplateUi();

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  clearStatus();

  const pdfFile = document.getElementById("pdf-file").files[0];
  const excelFile = document.getElementById("excel-file").files[0];
  const source = sourceSelect.value;
  const startMarker = startMarkerInput.value.trim();
  const endMarker = endMarkerInput.value.trim();

  if (!pdfFile || !excelFile) {
    setStatus("Please select a PDF and an Excel template.", "error");
    return;
  }

  if (source === "custom" && !startMarker) {
    setStatus("Start text is required for the custom template.", "error");
    updateTemplateUi();
    return;
  }

  const formData = new FormData();
  formData.append("source", source);
  formData.append("start_marker", startMarker);
  formData.append("end_marker", endMarker);
  formData.append("pdf", pdfFile);
  formData.append("excel", excelFile);

  submitBtn.disabled = true;
  submitBtn.textContent = "Processing...";
  setStatus("Extracting procedure and writing Excel...", "info");

  try {
    const response = await fetch("/api/job-order/generate", {
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
    setStatus("Job Order workbook generated successfully.", "info");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = "Generate Job Order";
  }
});
