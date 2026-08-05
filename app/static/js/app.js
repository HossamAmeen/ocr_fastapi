const form = document.getElementById("proforma-form");
const submitBtn = document.getElementById("submit-btn");
const statusEl = document.getElementById("status");
const resultsCard = document.getElementById("results-card");
const summaryEl = document.getElementById("summary");
const tableBody = document.querySelector("#items-table tbody");
const downloadLink = document.getElementById("download-link");

const pdfInput = document.getElementById("pdf-file");
const excelInput = document.getElementById("excel-file");
const pdfFileNameEl = document.getElementById("pdf-file-name");
const excelFileNameEl = document.getElementById("excel-file-name");

function formatMoney(value) {
  return Number(value).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function setStatus(message, type = "info") {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  statusEl.classList.remove("hidden");
}

function clearStatus() {
  statusEl.classList.add("hidden");
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

if (pdfInput) {
  pdfInput.addEventListener("change", () => {
    showSelectedFile(pdfFileNameEl, pdfInput.files[0]);
  });
}

if (excelInput) {
  excelInput.addEventListener("change", () => {
    showSelectedFile(excelFileNameEl, excelInput.files[0]);
  });
}

function renderResults(data) {
  summaryEl.innerHTML = `
    <div class="summary-item">
      <span>Items</span>
      <strong>${data.item_count}</strong>
    </div>
    <div class="summary-item">
      <span>Gross Total</span>
      <strong>$${formatMoney(data.gross_total)}</strong>
    </div>
    <div class="summary-item">
      <span>Output file</span>
      <strong>${data.filename}</strong>
    </div>
  `;

  tableBody.innerHTML = data.items
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

  downloadLink.href = data.download_url;
  downloadLink.download = data.filename;
  resultsCard.classList.remove("hidden");
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  clearStatus();

  const pdfFile = pdfInput ? pdfInput.files[0] : null;
  const excelFile = excelInput ? excelInput.files[0] : null;

  if (!pdfFile || !excelFile) {
    setStatus("Please select both PDF and Excel files.", "error");
    return;
  }

  const formData = new FormData();
  formData.append("pdf", pdfFile);
  formData.append("excel", excelFile);

  submitBtn.disabled = true;
  submitBtn.textContent = "Processing...";
  setStatus("Extracting PDF data and writing Excel...", "info");

  try {
    const response = await fetch("/api/proforma/generate", {
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
    setStatus("Proforma generated successfully.", "info");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = "Generate Proforma";
  }
});
