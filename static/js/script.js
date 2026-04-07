/* PDF → Excel UI */

const uploadArea = document.getElementById("uploadArea");
const fileInput = document.getElementById("fileInput");
const browseButton = document.getElementById("browseButton");
const convertBtn = document.getElementById("convertBtn");
const resetBtn = document.getElementById("resetBtn");
const downloadBtn = document.getElementById("downloadBtn");
const progressBar = document.getElementById("progressBar");
const progressFill = document.getElementById("progressFill");
const progressLabel = document.getElementById("progressLabel");
const globalLoader = document.getElementById("globalLoader");
const globalLoaderMsg = document.getElementById("globalLoaderMsg");
const globalLoaderHint = document.getElementById("globalLoaderHint");
const fileCountEl = document.getElementById("fileCount");
const summaryFileCount = document.getElementById("summaryFileCount");
const summaryPageCount = document.getElementById("summaryPageCount");
const pdfInfo = document.getElementById("pdfInfo");
const excelInfo = document.getElementById("excelInfo");
const excelRowInfo = document.getElementById("excelRowInfo");
const pdfViewer = document.getElementById("pdfViewer");
const pdfControls = document.getElementById("pdfControls");
const prevPageBtn = document.getElementById("prevPage");
const nextPageBtn = document.getElementById("nextPage");
const pageInfo = document.getElementById("pageInfo");
const excelHeader = document.getElementById("excelHeader");
const excelBody = document.getElementById("excelBody");
const reportModal = document.getElementById("reportModal");
const closeModal = document.getElementById("closeModal");
const closeModalBtn = document.getElementById("closeModalBtn");
const confirmDownload = document.getElementById("confirmDownload");
const modalPageCount = document.getElementById("modalPageCount");
const modalCellCount = document.getElementById("modalCellCount");
const modalTime = document.getElementById("modalTime");
const modalFileSize = document.getElementById("modalFileSize");
const totalFilesCounter = document.getElementById("totalFilesCounter");
const todayFilesCounter = document.getElementById("todayFilesCounter");
const currentDateEl = document.getElementById("currentDate");
const engineSelect = document.getElementById("engineSelect");
const engineHint = document.getElementById("engineHint");
const engineFidelityWarn = document.getElementById("engineFidelityWarn");
const convertapiOptionsRow = document.getElementById("convertapiOptionsRow");
const excelFidelityBanner = document.getElementById("excelFidelityBanner");

let currentFileId = null;
let currentPdfFile = null;
let pageCount = 0;
let currentPage = 1;
let previewPageUrls = [];
let pollTimer = null;
let convertStartTime = 0;
let lastOcrNoticeFileId = null;
/** Same-tick duplicate "change" guard (some browsers fire twice) */
let lastFileSelectKey = "";
let lastFileSelectAt = 0;

function setProgress(pct, labelText) {
  progressBar.style.display = "block";
  progressFill.style.width = `${Math.min(100, Math.max(0, pct))}%`;
  if (progressLabel) {
    if (labelText !== undefined && labelText !== null && labelText !== "") {
      progressLabel.style.display = "block";
      progressLabel.textContent = labelText;
    }
  }
}

/** Full-screen overlay for upload and slow /preview requests */
function setGlobalLoader(show, message, hint) {
  if (!globalLoader) return;
  if (globalLoaderMsg && message) globalLoaderMsg.textContent = message;
  if (globalLoaderHint) {
    if (hint === false || hint === "") {
      globalLoaderHint.style.display = "none";
    } else {
      globalLoaderHint.style.display = "block";
      globalLoaderHint.textContent = hint || "Large PDFs can take a while.";
    }
  }
  globalLoader.classList.toggle("d-none", !show);
  globalLoader.setAttribute("aria-busy", show ? "true" : "false");
}

function resetUi() {
  currentFileId = null;
  currentPdfFile = null;
  pageCount = 0;
  currentPage = 1;
  previewPageUrls = [];
  lastOcrNoticeFileId = null;
  if (pollTimer) {
    clearInterval(pollTimer);
    pollTimer = null;
  }
  fileCountEl.textContent = "0";
  summaryFileCount.textContent = "0";
  summaryPageCount.textContent = "0";
  pdfInfo.textContent = "No PDF loaded.";
  excelInfo.textContent = "Convert your PDF to populate this preview.";
  excelRowInfo.textContent = "Showing 0 rows";
  downloadBtn.disabled = true;
  progressBar.style.display = "none";
  progressFill.style.width = "0%";
  if (progressLabel) {
    progressLabel.style.display = "none";
    progressLabel.textContent = "";
  }
  setGlobalLoader(false);
  if (uploadArea) uploadArea.classList.remove("is-busy");
  pdfControls.style.display = "none";
  pdfViewer.innerHTML =
    '<div class="pdf-placeholder-inner text-center p-5 text-muted"><i class="fas fa-file-pdf fa-4x mb-3 d-block"></i><h3>No PDF to Display</h3><p>Upload a PDF to preview it here.</p></div>';
  excelHeader.innerHTML = "";
  excelBody.innerHTML =
    '<tr><td colspan="12" class="text-center p-5 text-muted">No data yet. Run conversion to see Excel structure.</td></tr>';
  if (excelFidelityBanner) {
    excelFidelityBanner.textContent = "";
    excelFidelityBanner.classList.add("d-none");
  }
  if (engineFidelityWarn) engineFidelityWarn.classList.add("d-none");
}

function syncEngineFidelityUI() {
  const isConvertApi = engineSelect && engineSelect.value === "convertapi";
  if (engineFidelityWarn && engineSelect) {
    engineFidelityWarn.classList.toggle("d-none", !isConvertApi);
  }
  if (convertapiOptionsRow && engineSelect) {
    convertapiOptionsRow.classList.toggle("d-none", !isConvertApi);
  }
}

function setExcelFidelityBanner(meta) {
  if (!excelFidelityBanner) return;
  if (meta && meta.conversion_engine === "convertapi" && meta.visual_match_note) {
    excelFidelityBanner.textContent = meta.visual_match_note;
    excelFidelityBanner.classList.remove("d-none");
  } else {
    excelFidelityBanner.textContent = "";
    excelFidelityBanner.classList.add("d-none");
  }
}

function formatBytes(n) {
  if (!n && n !== 0) return "0 KB";
  const u = ["B", "KB", "MB", "GB"];
  let i = 0;
  let v = n;
  while (v >= 1024 && i < u.length - 1) {
    v /= 1024;
    i++;
  }
  return `${v.toFixed(i > 0 ? 1 : 0)} ${u[i]}`;
}

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s;
  return d.innerHTML;
}

function showNotification(message, type = "info") {
  const notification = document.createElement("div");
  notification.style.cssText = `
        position: fixed; top: 20px; right: 20px; padding: 1rem 1.5rem; border-radius: 8px;
        color: white; font-weight: 600; z-index: 9999; display: flex; align-items: center;
        gap: 0.8rem; box-shadow: 0 4px 12px rgba(0,0,0,0.15); animation: slideIn 0.3s ease; max-width: 400px;
    `;

  // Default/info should also be green as requested
  if (type === "error") notification.style.background = "var(--error-red)";
  else notification.style.background = "var(--success-green)";

  let icon = "check-circle";
  if (type === "error") icon = "exclamation-circle";

  const i = document.createElement("i");
  i.className = `fas fa-${icon}`;
  const span = document.createElement("span");
  span.textContent = String(message ?? "");
  notification.appendChild(i);
  notification.appendChild(span);

  document.body.appendChild(notification);

  setTimeout(() => {
    notification.style.animation = "slideOut 0.3s ease";
    setTimeout(() => {
      if (notification.parentNode) notification.parentNode.removeChild(notification);
    }, 300);
  }, 5000);

  if (!document.querySelector("#notification-styles")) {
    const style = document.createElement("style");
    style.id = "notification-styles";
    style.textContent = `
            @keyframes slideIn { from { transform: translateX(100%); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
            @keyframes slideOut { from { transform: translateX(0); opacity: 1; } to { transform: translateX(100%); opacity: 0; } }
        `;
    document.head.appendChild(style);
  }
}

/** Bootstrap toast (top-right). Falls back to alert if Toast unavailable. */
function showAppToast(title, body, opts = {}) {
  const el = document.getElementById("appToast");
  const titleEl = document.getElementById("appToastTitle");
  const bodyEl = document.getElementById("appToastBody");
  const headerEl = document.getElementById("appToastHeader");
  const iconEl = document.getElementById("appToastIcon");
  if (!el || !titleEl || !bodyEl || !headerEl) {
    window.alert(`${title}\n\n${body}`);
    return;
  }
  if (typeof bootstrap === "undefined" || !bootstrap.Toast) {
    window.alert(`${title}\n\n${body}`);
    return;
  }

  titleEl.textContent = title;
  bodyEl.textContent = body;

  const variant = opts.variant || "info";
  const closeBtn = headerEl.querySelector(".btn-close");
  headerEl.className = "toast-header";
  headerEl.style.background = "";
  headerEl.style.color = "";
  if (closeBtn) closeBtn.className = "btn-close";

  if (variant === "danger") {
    headerEl.classList.add("bg-danger", "text-white");
    if (iconEl) iconEl.className = "fas fa-circle-exclamation me-2";
    if (closeBtn) closeBtn.classList.add("btn-close-white");
  } else if (variant === "success") {
    headerEl.classList.add("bg-success", "text-white");
    if (iconEl) iconEl.className = "fas fa-check-circle me-2";
    if (closeBtn) closeBtn.classList.add("btn-close-white");
  } else if (variant === "warning") {
    headerEl.classList.add("bg-warning", "text-dark");
    if (iconEl) iconEl.className = "fas fa-exclamation-triangle me-2";
  } else {
    headerEl.style.background = "linear-gradient(135deg, #1e3c72, #064a9c)";
    headerEl.style.color = "#fff";
    if (iconEl) iconEl.className = "fas fa-bell me-2";
    if (closeBtn) closeBtn.classList.add("btn-close-white");
  }

  const delay = opts.delay ?? 5000;
  const t = bootstrap.Toast.getOrCreateInstance(el, { autohide: true, delay });
  t.show();
}

function renderPdfPage() {
  if (!previewPageUrls.length) return;
  const url = previewPageUrls[currentPage - 1];
  pdfViewer.innerHTML = "";
  const wrap = document.createElement("div");
  wrap.className = "pdf-canvas-container text-center";
  const img = document.createElement("img");
  img.src = url;
  img.alt = `Page ${currentPage}`;
  img.className = "img-fluid shadow-sm border";
  img.style.maxWidth = "100%";
  wrap.appendChild(img);
  pdfViewer.appendChild(wrap);
  pageInfo.textContent = `Page ${currentPage} of ${pageCount}`;
  prevPageBtn.disabled = currentPage <= 1;
  nextPageBtn.disabled = currentPage >= pageCount;
}

function bumpStats() {
  try {
    const k = "pdf2xlsx_total";
    const d = new Date().toDateString();
    const dk = `pdf2xlsx_day_${d}`;
    const t = parseInt(localStorage.getItem(k) || "0", 10) + 1;
    localStorage.setItem(k, String(t));
    const td = parseInt(localStorage.getItem(dk) || "0", 10) + 1;
    localStorage.setItem(dk, String(td));
    totalFilesCounter.textContent = String(t);
    todayFilesCounter.textContent = String(td);
  } catch (_) {
    /* ignore */
  }
}

function loadPreviewMeta() {
  if (!currentFileId) return Promise.resolve();
  return fetch(`/preview/${currentFileId}`)
    .then((r) => {
      if (r.status === 404) return Promise.reject(new Error("preview_404"));
      return r.json();
    })
    .then((data) => {
      pageCount = data.page_count || 0;
      previewPageUrls = (data.pdf_pages || []).map((p) => p.url);
      if (previewPageUrls.length) {
        currentPage = 1;
        pdfControls.style.display = "flex";
        renderPdfPage();
        pdfInfo.textContent = `${pageCount} page(s) — server preview`;
      }
      const ep = data.excel_preview || {};
      const rows = ep.rows || [];
      const headers = ep.headers || [];
      if (headers.length && rows.length) {
        excelHeader.innerHTML = `<tr>${headers
          .map((h) => `<th>${escapeHtml(String(h))}</th>`)
          .join("")}</tr>`;
        excelBody.innerHTML = rows
          .map(
            (row) =>
              `<tr>${row.map((c) => `<td>${escapeHtml(String(c))}</td>`).join("")}</tr>`
          )
          .join("");
        excelRowInfo.textContent = `Showing ${rows.length} preview rows`;
        excelInfo.textContent = "Excel layout preview (subset of cells)";
      }
    });
}

async function uploadFile(file) {
  const fd = new FormData();
  fd.append("file", file);
  const res = await fetch("/upload", { method: "POST", body: fd });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || "Upload failed");
  }
  return res.json();
}

function buildConvertQueryParams() {
  const eng = engineSelect && engineSelect.value ? engineSelect.value : "local";
  const params = new URLSearchParams();
  params.set("engine", eng);
  if (eng === "convertapi") {
    const sheetOne = document.getElementById("sheetOne");
    const singleSheet = sheetOne && sheetOne.checked ? "1" : "0";
    params.set("single_sheet", singleSheet);
    const inc = document.getElementById("includeFormatting");
    params.set("include_formatting", inc && inc.checked ? "1" : "0");
  }
  return params.toString();
}

function pollConvert() {
  if (!currentFileId) return;
  const q = buildConvertQueryParams();
  fetch(`/convert/${currentFileId}?${q}`)
    .then((r) => {
      if (r.status === 404) {
        if (pollTimer) {
          clearInterval(pollTimer);
          pollTimer = null;
        }
        convertBtn.disabled = false;
        showNotification(
          "Session not found — re-upload the PDF. If this persists, the server must run Gunicorn with a single worker (see README).",
          "error"
        );
        return Promise.reject(new Error("convert_404"));
      }
      return r.json();
    })
    .then((data) => {
      if (!data) return;
      const p = data.progress ?? 0;
      const stepLabel =
        data.message || `Working… ${Math.round(p)}%`;
      setProgress(p, stepLabel);
      if (data.status === "done") {
        clearInterval(pollTimer);
        pollTimer = null;
        convertBtn.disabled = false;
        downloadBtn.disabled = false;
        setProgress(100, "Complete");
        bumpStats();
        const meta = data.excel_meta || {};
        const rows = meta.preview_rows || [];
        summaryPageCount.textContent = String(meta.page_count || pageCount || 0);
        excelInfo.textContent =
          meta.conversion_engine === "convertapi"
            ? "ConvertAPI: table data in Excel (not a visual copy of the PDF). Preview below."
            : "Conversion complete. Preview updates below.";
        setExcelFidelityBanner(meta);
        const elapsed = ((Date.now() - convertStartTime) / 1000).toFixed(1);
        modalPageCount.textContent = String(meta.page_count || pageCount || 0);
        const cells = (meta.rows_written || 0) * (meta.cols_written || 0);
        modalCellCount.textContent = String(cells || rows.length || 0);
        modalTime.textContent = `${elapsed}s`;
        modalFileSize.textContent = currentPdfFile
          ? formatBytes(currentPdfFile.size)
          : "—";
        setGlobalLoader(
          true,
          "Loading preview panels…",
          "Updating PDF thumbnails and Excel grid from the server."
        );
        loadPreviewMeta()
          .then(() => {
            showNotification("Conversion complete. Excel ready.", "success");
            if (meta.used_ocr && lastOcrNoticeFileId !== currentFileId) {
              lastOcrNoticeFileId = currentFileId;
              showNotification("OCR used (scanned/empty-text pages detected).", "success");
            }
          })
          .catch(() => {})
          .finally(() => setGlobalLoader(false));
      } else if (data.status === "error") {
        clearInterval(pollTimer);
        pollTimer = null;
        convertBtn.disabled = false;
        showNotification(data.message || "Conversion failed", "error");
      }
    })
    .catch(() => {});
}

browseButton.addEventListener("click", (e) => {
  e.preventDefault();
  e.stopPropagation();
  fileInput.click();
});

uploadArea.addEventListener("click", (e) => {
  if (browseButton.contains(e.target)) {
    return;
  }
  fileInput.click();
});

["dragenter", "dragover"].forEach((ev) => {
  uploadArea.addEventListener(ev, (e) => {
    e.preventDefault();
    uploadArea.classList.add("drag-over");
  });
});

["dragleave", "drop"].forEach((ev) => {
  uploadArea.addEventListener(ev, (e) => {
    e.preventDefault();
    uploadArea.classList.remove("drag-over");
  });
});

uploadArea.addEventListener("drop", (e) => {
  const f = e.dataTransfer.files[0];
  if (f && f.name.toLowerCase().endsWith(".pdf")) {
    // Do not assign fileInput.files here — that can fire "change" and run upload twice
    handleFileSelect(f);
  }
});

fileInput.addEventListener("change", () => {
  const f = fileInput.files[0];
  if (f) handleFileSelect(f);
});

async function handleFileSelect(file) {
  if (!file) return;
  const dedupeKey = `${file.name}|${file.size}|${file.lastModified}`;
  const now = Date.now();
  if (dedupeKey === lastFileSelectKey && now - lastFileSelectAt < 80) {
    return;
  }
  lastFileSelectKey = dedupeKey;
  lastFileSelectAt = now;

  resetUi();
  currentPdfFile = file;
  fileCountEl.textContent = "1";
  summaryFileCount.textContent = "1";
  pdfInfo.textContent = `Selected: ${file.name}`;
  if (uploadArea) uploadArea.classList.add("is-busy");
  setGlobalLoader(true, "Uploading PDF…", "Sending your file to the server.");
  try {
    const data = await uploadFile(file);
    currentFileId = data.file_id;
    pageCount = data.page_count || 0;
    summaryPageCount.textContent = String(pageCount);
    setGlobalLoader(
      true,
      "Loading preview from server…",
      "Building PDF thumbnails and metadata. This can take 10–30 seconds for large files."
    );
    await loadPreviewMeta();
    syncEngineFidelityUI();
  } catch (e) {
    showNotification(e.message || String(e), "error");
    resetUi();
  } finally {
    if (uploadArea) uploadArea.classList.remove("is-busy");
    setGlobalLoader(false);
  }
}

convertBtn.addEventListener("click", () => {
  if (!currentFileId) {
    showNotification("Please upload a PDF first.", "error");
    return;
  }
  convertBtn.disabled = true;
  convertStartTime = Date.now();
  setProgress(5, "Starting conversion…");
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(pollConvert, 700);
  pollConvert();
});

resetBtn.addEventListener("click", () => {
  fileInput.value = "";
  lastFileSelectKey = "";
  lastFileSelectAt = 0;
  resetUi();
  showNotification("Reset tool: Successfully reset the tool", "success");
});

prevPageBtn.addEventListener("click", () => {
  if (currentPage > 1) {
    currentPage -= 1;
    renderPdfPage();
  }
});

nextPageBtn.addEventListener("click", () => {
  if (currentPage < pageCount) {
    currentPage += 1;
    renderPdfPage();  
  }
});

downloadBtn.addEventListener("click", () => {
  reportModal.style.display = "flex";
});

closeModal.addEventListener("click", () => {
  reportModal.style.display = "none";
});
closeModalBtn.addEventListener("click", () => {
  reportModal.style.display = "none";
});

confirmDownload.addEventListener("click", () => {
  if (!currentFileId) return;
  window.location.href = `/download/${currentFileId}`;
  reportModal.style.display = "none";
});

document.querySelectorAll(".info-tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    document.querySelectorAll(".info-tab").forEach((t) => t.classList.remove("active"));
    document.querySelectorAll(".tab-pane-info").forEach((p) => p.classList.remove("active"));
    tab.classList.add("active");
    const id = tab.getAttribute("data-tab");
    const pane = document.getElementById(`${id}-tab`);
    if (pane) pane.classList.add("active");
  });
});

document.addEventListener("DOMContentLoaded", () => {
  resetUi();
  fileInput.value = "";
  lastFileSelectKey = "";
  lastFileSelectAt = 0;

  showNotification(
    "Page has been reset successfully",
    "info"
  );

  try {
    const k = "pdf2xlsx_total";
    const d = new Date().toDateString();
    const dk = `pdf2xlsx_day_${d}`;
    totalFilesCounter.textContent = localStorage.getItem(k) || "0";
    todayFilesCounter.textContent = localStorage.getItem(dk) || "0";
  } catch (_) {
    totalFilesCounter.textContent = "0";
    todayFilesCounter.textContent = "0";
  }
  if (currentDateEl) {
    currentDateEl.textContent = new Date().toLocaleDateString();
  }

  fetch("/api/config")
    .then((r) => r.json())
    .then((cfg) => {
      const opt = document.getElementById("optConvertapi");
      if (opt && cfg.convertapi_configured) {
        opt.disabled = false;
        opt.removeAttribute("title");
      }
      if (engineHint && cfg && !cfg.convertapi_configured) {
        engineHint.textContent =
          "Local engine works out of the box. Copy env.example to .env and add CONVERTAPI_SECRET / CONVERTAPI_SECRET_SANDBOX (see https://www.convertapi.com/a/authentication).";
      } else if (engineHint && cfg && cfg.convertapi_configured && cfg.convertapi_env) {
        engineHint.textContent = `ConvertAPI tokens loaded. Active env: ${cfg.convertapi_env} (set CONVERTAPI_ENV=sandbox or production in .env).`;
      }
      syncEngineFidelityUI();
    })
    .catch(() => {});

  if (engineSelect) {
    engineSelect.addEventListener("change", syncEngineFidelityUI);
  }
});
