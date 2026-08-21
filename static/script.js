(function () {
  "use strict";

  // ---------- Font-size toolbar (unchanged UX from the original app) ----------
  const root = document.documentElement;
  const FONT_SCALE_MIN = 0.85;
  const FONT_SCALE_MAX = 1.3;
  const FONT_SCALE_STEP = 0.05;
  let fontScale = 1;

  function applyFontScale() {
    root.style.setProperty("--font-scale", fontScale.toFixed(2));
    const dec = document.getElementById("decreaseFontBtn");
    const inc = document.getElementById("increaseFontBtn");
    if (dec) dec.disabled = fontScale <= FONT_SCALE_MIN + 1e-6;
    if (inc) inc.disabled = fontScale >= FONT_SCALE_MAX - 1e-6;
  }

  document.getElementById("decreaseFontBtn")?.addEventListener("click", () => {
    fontScale = Math.max(FONT_SCALE_MIN, fontScale - FONT_SCALE_STEP);
    applyFontScale();
  });
  document.getElementById("increaseFontBtn")?.addEventListener("click", () => {
    fontScale = Math.min(FONT_SCALE_MAX, fontScale + FONT_SCALE_STEP);
    applyFontScale();
  });
  document.getElementById("resetFontBtn")?.addEventListener("click", () => {
    fontScale = 1;
    applyFontScale();
  });
  applyFontScale();

  // ---------- Elements ----------
  const fileInput = document.getElementById("file");
  const fileNameLabel = document.getElementById("fileName");
  const uploadStatus = document.getElementById("uploadStatus");
  const settingsSection = document.getElementById("settingsSection");
  const sheetSelect = document.getElementById("sheetSelect");
  const fontNameInput = document.getElementById("fontNameInput");
  const fontSizeInput = document.getElementById("fontSizeInput");
  const cleanBtn = document.getElementById("cleanBtn");
  const cleanStatus = document.getElementById("cleanStatus");
  const resultsSection = document.getElementById("resultsSection");
  const summary = document.getElementById("summary");
  const downloadBtn = document.getElementById("downloadBtn");
  const excelGridHead = document.getElementById("excelGridHead");
  const excelGridBody = document.getElementById("excelGridBody");
  const previewNote = document.getElementById("previewNote");

  let currentJobId = null;

  function setStatus(el, message, isError) {
    el.textContent = message || "";
    el.classList.toggle("error", !!isError);
  }

  function resetForNewUpload() {
    settingsSection.hidden = true;
    resultsSection.hidden = true;
    setStatus(cleanStatus, "");
    currentJobId = null;
  }

  // ---------- Step 1: upload + inspect ----------
  fileInput.addEventListener("change", async () => {
    const file = fileInput.files[0];
    resetForNewUpload();

    if (!file) {
      fileNameLabel.textContent = "Belum ada file dipilih";
      return;
    }
    fileNameLabel.textContent = file.name;

    const formData = new FormData();
    formData.append("file", file);

    setStatus(uploadStatus, "Membaca file...", false);

    try {
      const res = await fetch("/api/inspect", { method: "POST", body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Gagal membaca file.");

      currentJobId = data.job_id;
      sheetSelect.innerHTML = "";
      data.sheets.forEach((s) => {
        const opt = document.createElement("option");
        opt.value = s;
        opt.textContent = s;
        sheetSelect.appendChild(opt);
      });

      settingsSection.hidden = false;
      setStatus(uploadStatus, `File terbaca: ${data.file_name} (${data.sheets.length} sheet)`, false);
    } catch (err) {
      setStatus(uploadStatus, err.message, true);
    }
  });

  // ---------- Step 2: clean ----------
  cleanBtn.addEventListener("click", async () => {
    if (!currentJobId) return;

    const payload = {
      job_id: currentJobId,
      sheet: sheetSelect.value,
      font_name: fontNameInput.value || "Arial",
      font_size: fontSizeInput.value || 9,
    };

    cleanBtn.disabled = true;
    resultsSection.hidden = true;
    setStatus(cleanStatus, "Membersihkan file...", false);

    try {
      const res = await fetch("/api/clean", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Gagal membersihkan file.");

      renderResults(data);
      setStatus(cleanStatus, "", false);
    } catch (err) {
      setStatus(cleanStatus, err.message, true);
    } finally {
      cleanBtn.disabled = false;
    }
  });

  function renderResults(data) {
    summary.innerHTML = `<span>Sheet: <b>${escapeHtml(data.sheet)}</b></span>
      <span>Font dipakai: <b>${escapeHtml(data.font_name)}, ${data.font_size}pt</b></span>
      <span>Baris transaksi disimpan: <b>${data.total_rows}</b></span>`;

    downloadBtn.href = `/api/download/${data.job_id}`;

    excelGridHead.innerHTML = "";
    data.headers.forEach((h) => {
      const th = document.createElement("th");
      th.textContent = h;
      excelGridHead.appendChild(th);
    });

    excelGridBody.innerHTML = "";
    data.preview_rows.forEach((row) => {
      const tr = document.createElement("tr");
      row.forEach((val) => {
        const td = document.createElement("td");
        td.className = "excel-cell";
        td.textContent = val;
        tr.appendChild(td);
      });
      excelGridBody.appendChild(tr);
    });

    previewNote.textContent = data.preview_truncated
      ? `Menampilkan 20 dari ${data.total_rows} baris. Unduh file untuk melihat semuanya.`
      : `Menampilkan semua ${data.total_rows} baris.`;

    resultsSection.hidden = false;
  }

  function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
  }
})();
