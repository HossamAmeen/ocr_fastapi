const form = document.getElementById("job-over-form");
const submitBtn = document.getElementById("submit-btn");
const statusEl = document.getElementById("status");
const resultsCard = document.getElementById("results-card");
const summaryEl = document.getElementById("summary");
const linesTableBody = document.querySelector("#lines-table tbody");
const downloadLink = document.getElementById("download-link");

function setStatus(message, type = "info") {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  statusEl.classList.remove("hidden");
}

function clearStatus() {
  statusEl.classList.add("hidden");
}

function formatSource(source) {
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

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  clearStatus();

  const pdfFile = document.getElementById("pdf-file").files[0];
  const excelFile = document.getElementById("excel-file").files[0];
  const source = document.getElementById("source").value;

  if (!pdfFile || !excelFile) {
    setStatus("Please select a PDF and an Excel template.", "error");
    return;
  }

  const formData = new FormData();
  formData.append("pdf", pdfFile);
  formData.append("excel", excelFile);
  formData.append("source", source);

  submitBtn.disabled = true;
  submitBtn.textContent = "Processing...";
  setStatus("Extracting procedure and writing Excel...", "info");

  try {
    const response = await fetch("/api/job-over/generate", {
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
