/* ================= TEXT SIZE CONTROLS ================= */
const increaseFontBtn = document.getElementById("increaseFontBtn");
const decreaseFontBtn = document.getElementById("decreaseFontBtn");
const resetFontBtn = document.getElementById("resetFontBtn");
const root = document.documentElement;

const FONT_SCALE_MIN = 0.9;
const FONT_SCALE_STEP = 0.1;
const DEFAULT_FONT_SCALE = 1;

function setFontScale(nextScale) {
  const safeScale = Math.max(FONT_SCALE_MIN, nextScale);
  root.style.setProperty("--font-scale", safeScale);
  root.style.setProperty("--ui-scale", safeScale);
  localStorage.setItem("toolkit-font-scale", String(safeScale));

  decreaseFontBtn.disabled = safeScale <= FONT_SCALE_MIN;
}

const storedScale = Number(localStorage.getItem("toolkit-font-scale") || DEFAULT_FONT_SCALE);
setFontScale(Number.isFinite(storedScale) ? storedScale : DEFAULT_FONT_SCALE);

increaseFontBtn.addEventListener("click", () => {
  const currentScale = Number(root.style.getPropertyValue("--font-scale")) || DEFAULT_FONT_SCALE;
  setFontScale(currentScale + FONT_SCALE_STEP);
});

decreaseFontBtn.addEventListener("click", () => {
  const currentScale = Number(root.style.getPropertyValue("--font-scale")) || DEFAULT_FONT_SCALE;
  setFontScale(currentScale - FONT_SCALE_STEP);
});

resetFontBtn.addEventListener("click", () => {
  setFontScale(DEFAULT_FONT_SCALE);
});

/* ================= EXCEL COMPARISON ================= */
const form = document.getElementById("compareForm");
const statusEl = document.getElementById("status");
const resultsSection = document.getElementById("resultsSection");
const summaryEl = document.getElementById("summary");
const compareBtn = document.getElementById("compareBtn");

const compareSelectBar = document.getElementById("compareSelectBar");
const compareFileSelect = document.getElementById("compareFileSelect");

const noDiffMessage = document.getElementById("noDiffMessage");
const diffContent = document.getElementById("diffContent");
const sheetTabsEl = document.getElementById("sheetTabs");
const hideEqualRowsCheckbox = document.getElementById("hideEqualRows");

const excelGrid = document.getElementById("excelGrid");
const excelGridHead = document.getElementById("excelGridHead");
const excelGridBody = document.getElementById("excelGridBody");
const sheetOnlyMessage = document.getElementById("sheetOnlyMessage");

const mergeSection = document.getElementById("mergeSection");
const baseFileSelect = document.getElementById("baseFileSelect");
const bulkFile1Btn = document.getElementById("bulkFile1Btn");
const bulkFile2Btn = document.getElementById("bulkFile2Btn");
const downloadMergeBtn = document.getElementById("downloadMergeBtn");
const mergeStatus = document.getElementById("mergeStatus");

let currentData = null;   // seluruh respons API: {job_id, file1_name, compare_names, comparisons}
let compareIndex = 0;     // indeks arah perbandingan yang sedang ditampilkan
let currentSheet = null;  // nama sheet yang sedang ditampilkan
let mergeChoices = {};    // key -> "file1" | "file2"

function bindFileLabel(inputId, labelId) {
  const input = document.getElementById(inputId);
  const label = document.getElementById(labelId);
  input.addEventListener("change", () => {
    const names = Array.from(input.files || []).map((file) => file.name);
    label.textContent = names.length ? names.join(", ") : "Belum ada file dipilih";
  });
}
bindFileLabel("file1", "file1Name");
bindFileLabel("file2", "file2Name");
bindFileLabel("template", "templateName");

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeAttr(str) {
  return escapeHtml(str).replace(/"/g, "&quot;");
}

function formatVal(v) {
  if (v === "" || v === null || v === undefined) return "";
  return escapeHtml(String(v));
}

function mergeKey(idx, sheet, row, col) {
  return `${idx}||${sheet}||${row}||${col}`;
}

function getDefaultSource(status) {
  // file1 tidak punya nilai -> defaultnya pakai file2, selain itu default file1.
  return status === "missing1" ? "file2" : "file1";
}

function renderTemplateBadge(data) {
  const badge = document.getElementById("templateBadge");
  if (!badge) return;
  if (data.template_detected) {
    const sourceLabel = data.template_source === "uploaded" ? "diunggah manual" : "terdeteksi otomatis";
    const nameLabel = data.template_name ? escapeHtml(data.template_name) : "sebuah file";
    badge.innerHTML = `<span class="template-badge-icon">🧩</span> Memakai <b>${nameLabel}</b> sebagai template/skeleton (${sourceLabel}) — header tabel mengikuti file ini.`;
    badge.className = "template-badge active";
  } else {
    badge.innerHTML = `<span class="template-badge-icon">—</span> Tidak memakai template. Header tabel mengikuti header umum yang ditemukan di file yang diunggah.`;
    badge.className = "template-badge inactive";
  }
  badge.hidden = false;
}

function forEachCorrectableCell(comp, callback) {
  for (const sheetName of Object.keys(comp.sheets)) {
    const sheetData = comp.sheets[sheetName];
    for (const row of sheetData.rows) {
      for (const cell of row.cells) {
        if (cell.status === "diff" || cell.status === "missing1" || cell.status === "missing2" || cell.status === "empty_vs_value") {
          callback(sheetName, row.row, cell);
        }
      }
    }
  }
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  statusEl.className = "status";
  statusEl.textContent = "Membandingkan file…";
  resultsSection.hidden = true;
  mergeStatus.textContent = "";
  compareBtn.disabled = true;

  const formData = new FormData(form);

  try {
    const res = await fetch("/api/compare", { method: "POST", body: formData });
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Perbandingan gagal.");
    }

    currentData = data;
    compareIndex = 0;
    currentSheet = null;
    mergeChoices = {};

    compareFileSelect.innerHTML = "";
    const comparisonLabels = data.compare_names || [];
    comparisonLabels.forEach((name, idx) => {
      const opt = document.createElement("option");
      opt.value = String(idx);
      opt.textContent = name;
      compareFileSelect.appendChild(opt);
    });
    compareFileSelect.value = "0";
    compareSelectBar.hidden = comparisonLabels.length <= 1;

    renderTemplateBadge(data);
    statusEl.textContent = "";

    resultsSection.hidden = false;
    renderComparison();
  } catch (err) {
    statusEl.className = "status error";
    statusEl.textContent = err.message;
  } finally {
    compareBtn.disabled = false;
  }
});

compareFileSelect.addEventListener("change", () => {
  compareIndex = Number(compareFileSelect.value);
  currentSheet = null;
  renderComparison();
});

hideEqualRowsCheckbox.addEventListener("change", () => {
  excelGrid.classList.toggle("hide-equal-rows", hideEqualRowsCheckbox.checked);
});

function renderComparison() {
  const comp = currentData.comparisons[compareIndex];

  const hasAnyDifference =
    comp.summary.total_differences > 0 ||
    comp.sheets_only_in_file1.length > 0 ||
    comp.sheets_only_in_file2.length > 0;

  const leftName = comp.left_name || currentData.file1_name || "File 1";
  const rightName = comp.right_name || currentData.file2_name || "File 2";
  const headerRowInfo = currentSheet && comp.sheets[currentSheet]
    ? comp.sheets[currentSheet].header_row_used
    : null;

  let summaryHtml = `
    <span><b>${comp.summary.total_differences}</b> sel berbeda</span>
    <span>Arah: ${escapeHtml(comp.compare_name || `${leftName} → ${rightName}`)}</span>
    <span>Sheet yang dibandingkan: ${comp.summary.sheets_compared.map(escapeHtml).join(", ") || "tidak ada"}</span>
    <span>Baris header terdeteksi: ${headerRowInfo ?? "-"}</span>
  `;
  if (comp.sheets_only_in_file1.length) {
    summaryHtml += `<span>Hanya ada di ${escapeHtml(leftName)}: ${comp.sheets_only_in_file1.map(escapeHtml).join(", ")}</span>`;
  }
  if (comp.sheets_only_in_file2.length) {
    summaryHtml += `<span>Hanya ada di ${escapeHtml(rightName)}: ${comp.sheets_only_in_file2.map(escapeHtml).join(", ")}</span>`;
  }
  summaryEl.innerHTML = summaryHtml;

  if (!hasAnyDifference) {
    noDiffMessage.hidden = false;
    diffContent.hidden = true;
    return;
  }

  noDiffMessage.hidden = true;
  diffContent.hidden = false;

  if (!currentSheet || !comp.summary.sheets_compared.includes(currentSheet)) {
    currentSheet =
      comp.summary.sheets_compared[0] ||
      comp.sheets_only_in_file1[0] ||
      comp.sheets_only_in_file2[0] ||
      null;
  }

  renderSheetTabs(comp);
  renderGrid();
}

function renderSheetTabs(comp) {
  sheetTabsEl.innerHTML = "";

  const makeTab = (actualName, label, extraClass) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `sheet-tab ${extraClass || ""} ${currentSheet === actualName ? "active" : ""}`.trim();
    btn.textContent = label;
    btn.addEventListener("click", () => {
      currentSheet = actualName;
      renderComparison();
    });
    sheetTabsEl.appendChild(btn);
  };

  comp.summary.sheets_compared.forEach((name) => makeTab(name, name));
  comp.sheets_only_in_file1.forEach((name) => makeTab(name, `${name} ⚠`, "warn"));
  comp.sheets_only_in_file2.forEach((name) => makeTab(name, `${name} ⚠`, "warn"));
}

function buildCellHTML(sheet, rowLabel, cell) {
  const status = cell.status;

  if (status === "same") {
    return `<td class="excel-cell same">${formatVal(cell.v1)}</td>`;
  }
  if (status === "empty") {
    return `<td class="excel-cell empty"></td>`;
  }

  const key = mergeKey(compareIndex, sheet, rowLabel, cell.col);
  const selected = mergeChoices[key] || getDefaultSource(status);
  const keyAttr = escapeAttr(key);

  if (status === "diff") {
    return `<td class="excel-cell diff">
      <span class="val-choice ${selected === "file1" ? "selected" : ""}" data-key="${keyAttr}" data-source="file1">${formatVal(cell.v1)}</span>
      <span class="val-choice ${selected === "file2" ? "selected" : ""}" data-key="${keyAttr}" data-source="file2">${formatVal(cell.v2)}</span>
    </td>`;
  }
  if (status === "empty_vs_value") {
    return `<td class="excel-cell empty-vs-value">
      <span class="val-choice ${selected === "file1" ? "selected" : ""}" data-key="${keyAttr}" data-source="file1">${formatVal(cell.v1)}</span>
      <span class="val-choice ${selected === "file2" ? "selected" : ""}" data-key="${keyAttr}" data-source="file2">${formatVal(cell.v2)}</span>
    </td>`;
  }
  if (status === "missing1") {
    return `<td class="excel-cell missing">
      <span class="val-choice disabled ${selected === "file1" ? "selected" : ""}" data-key="${keyAttr}" data-source="file1">(tidak ada)</span>
      <span class="val-choice ${selected === "file2" ? "selected" : ""}" data-key="${keyAttr}" data-source="file2">${formatVal(cell.v2)}</span>
    </td>`;
  }
  // missing2
  return `<td class="excel-cell missing">
    <span class="val-choice ${selected === "file1" ? "selected" : ""}" data-key="${keyAttr}" data-source="file1">${formatVal(cell.v1)}</span>
    <span class="val-choice disabled ${selected === "file2" ? "selected" : ""}" data-key="${keyAttr}" data-source="file2">(tidak ada)</span>
  </td>`;
}

function buildRowHTML(sheet, row) {
  const hasDiff = row.cells.some(
    (c) => c.status === "diff" || c.status === "missing1" || c.status === "missing2" || c.status === "empty_vs_value"
  );
  let html = `<tr data-hasdiff="${hasDiff ? 1 : 0}"><th class="row-header">${escapeHtml(row.row)}</th>`;
  for (const cell of row.cells) {
    html += buildCellHTML(sheet, row.row, cell);
  }
  html += `</tr>`;
  return html;
}

function renderGrid() {
  const comp = currentData.comparisons[compareIndex];

  if (!currentSheet || !comp.sheets[currentSheet]) {
    excelGrid.hidden = true;
    sheetOnlyMessage.hidden = false;
    mergeSection.hidden = true;
    if (currentSheet && comp.sheets_only_in_file1.includes(currentSheet)) {
      sheetOnlyMessage.textContent = `Sheet "${currentSheet}" hanya ada di ${comp.left_name || currentData.file1_name || "File 1"}, sehingga tidak ada yang bisa dibandingkan.`;
    } else if (currentSheet && comp.sheets_only_in_file2.includes(currentSheet)) {
      sheetOnlyMessage.textContent = `Sheet "${currentSheet}" hanya ada di ${comp.right_name || currentData.file2_name || "File 2"}, sehingga tidak ada yang bisa dibandingkan.`;
    } else {
      sheetOnlyMessage.textContent = "Tidak ada sheet untuk ditampilkan.";
    }
    return;
  }

  excelGrid.hidden = false;
  sheetOnlyMessage.hidden = true;
  mergeSection.hidden = false;

  const sheetData = comp.sheets[currentSheet];
  const colHeaders = Array.isArray(sheetData.display_col_headers) && sheetData.display_col_headers.length
    ? sheetData.display_col_headers
    : (Array.isArray(sheetData.reference_col_headers_display) && sheetData.reference_col_headers_display.length
      ? sheetData.reference_col_headers_display
      : (Array.isArray(sheetData.reference_col_headers)
        ? sheetData.reference_col_headers
        : (Array.isArray(sheetData.col_headers) ? sheetData.col_headers : [])));
  const rowHeaders = Array.isArray(sheetData.reference_row_headers)
    ? sheetData.reference_row_headers
    : (Array.isArray(sheetData.row_headers) ? sheetData.row_headers : []);

  const leftName = comp.left_name || currentData.file1_name || "File 1";
  let headHtml = `<th class="corner-header">${escapeHtml(leftName)}</th>`;
  for (const colLabel of colHeaders) {
    headHtml += `<th>${escapeHtml(colLabel)}</th>`;
  }
  excelGridHead.innerHTML = headHtml;

  let bodyHtml = "";
  for (const row of sheetData.rows) {
    const rowLabel = row.row;
    if (rowHeaders.length && !rowHeaders.includes(rowLabel)) {
      continue;
    }
    bodyHtml += buildRowHTML(currentSheet, row);
  }
  excelGridBody.innerHTML = bodyHtml || `<tr><td>Tidak ada data.</td></tr>`;

  excelGrid.classList.toggle("hide-equal-rows", hideEqualRowsCheckbox.checked);
}

// Delegasi klik untuk memilih nilai file1/file2 pada sel yang berbeda (dipakai fitur merge)
excelGridBody.addEventListener("click", (e) => {
  const span = e.target.closest(".val-choice");
  if (!span || span.classList.contains("disabled")) return;

  const key = span.dataset.key;
  const source = span.dataset.source;
  mergeChoices[key] = source;

  const td = span.parentElement;
  td.querySelectorAll(".val-choice").forEach((s) => s.classList.remove("selected"));
  span.classList.add("selected");
});

/* ================= GABUNGKAN FILE (MERGE) ================= */

bulkFile1Btn.addEventListener("click", () => {
  if (!currentData) return;
  const comp = currentData.comparisons[compareIndex];
  forEachCorrectableCell(comp, (sheet, rowLabel, cell) => {
    mergeChoices[mergeKey(compareIndex, sheet, rowLabel, cell.col)] = "file1";
  });
  renderGrid();
});

bulkFile2Btn.addEventListener("click", () => {
  if (!currentData) return;
  const comp = currentData.comparisons[compareIndex];
  forEachCorrectableCell(comp, (sheet, rowLabel, cell) => {
    mergeChoices[mergeKey(compareIndex, sheet, rowLabel, cell.col)] = "file2";
  });
  renderGrid();
});

downloadMergeBtn.addEventListener("click", async () => {
  if (!currentData || !currentData.job_id) {
    mergeStatus.className = "status error";
    mergeStatus.textContent = "Belum ada hasil perbandingan untuk digabungkan.";
    return;
  }

  const comp = currentData.comparisons[compareIndex];
  const choices = [];
  forEachCorrectableCell(comp, (sheet, rowLabel, cell) => {
    const key = mergeKey(compareIndex, sheet, rowLabel, cell.col);
    const source = mergeChoices[key] || getDefaultSource(cell.status);
    choices.push({ sheet, row: rowLabel, col: cell.col, source });
  });

  mergeStatus.className = "status";
  mergeStatus.textContent = "Membuat file gabungan…";
  downloadMergeBtn.disabled = true;

  try {
    const res = await fetch("/api/merge", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        job_id: currentData.job_id,
        target_index: compareIndex,
        base_file: baseFileSelect.value,
        choices,
      }),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || "Gagal membuat file gabungan.");
    }

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "file_gabungan.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

    mergeStatus.textContent = "File gabungan berhasil diunduh.";
  } catch (err) {
    mergeStatus.className = "status error";
    mergeStatus.textContent = err.message;
  } finally {
    downloadMergeBtn.disabled = false;
  }
});