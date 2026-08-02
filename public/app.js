// ============================================================
//  Pivot Stock Masuk Odoo — Client-Side Logic
// ============================================================

/* ========= STATE ========= */
let selectedFile = null;
let lastDownloadUrl = null;

/* ========= ELEMENTS ========= */
const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const fileChosen = document.getElementById("fileChosen");
const fileName = document.getElementById("fileName");
const fileSize = document.getElementById("fileSize");
const processBtn = document.getElementById("processBtn");
const processBtnText = document.getElementById("processBtnText");
const resultCard = document.getElementById("resultCard");
const downloadBtn = document.getElementById("downloadBtn");

/* ========= HTML ESCAPE ========= */
// Build entities dynamically to avoid formatter collisions in HTML files
const AMP = String.fromCharCode(38) + "amp;";
const LT_ENT = String.fromCharCode(38) + "lt;";
const GT_ENT = String.fromCharCode(38) + "gt;";
const QUOT_ENT = String.fromCharCode(38) + "quot;";
const APOS_ENT = String.fromCharCode(38) + "#39;";

function escape(str) {
    const s = String(str);
    return s
        .replace(/&/g, AMP)
        .replace(/</g, LT_ENT)
        .replace(/>/g, GT_ENT)
        .replace(/"/g, QUOT_ENT)
        .replace(/'/g, APOS_ENT);
}

function escapeAttr(str) {
    return escape(str);
}

/* ========= TOAST ========= */
function toast(message, type = "info") {
    const container = document.getElementById("toastContainer");
    const t = document.createElement("div");
    t.className = `toast ${type}`;
    const icon = type === "success"
        ? '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>'
        : '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M12 9v3.75m9-.75a9 9 0 11-18 0 9 9 0 0118 0zm-9 3.75h.008v.008H12v-.008z"/></svg>';
    t.innerHTML = `${icon}<span>${escape(message)}</span>`;
    container.appendChild(t);
    setTimeout(() => t.remove(), 3300);
}

/* ========= FILE HANDLING ========= */
function formatSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}

function setFile(file) {
    if (!file) return;
    if (!/\.xlsx$/i.test(file.name)) {
        toast("Hanya file .xlsx yang diperbolehkan.", "error");
        return;
    }
    if (file.size > 50 * 1024 * 1024) {
        toast("File terlalu besar. Maksimal 50 MB.", "error");
        return;
    }
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatSize(file.size);
    fileChosen.classList.remove("hidden");
    processBtn.disabled = false;
}

function removeFile() {
    selectedFile = null;
    fileInput.value = "";
    fileChosen.classList.add("hidden");
    processBtn.disabled = true;
}

/* ========= DRAG & DROP ========= */
["dragenter", "dragover"].forEach((ev) =>
    dropZone.addEventListener(ev, (e) => {
        e.preventDefault();
        dropZone.classList.add("dragging");
    })
);
["dragleave", "drop"].forEach((ev) =>
    dropZone.addEventListener(ev, (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragging");
    })
);
dropZone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files[0];
    if (file) setFile(file);
});
dropZone.addEventListener("click", (e) => {
    if (e.target.tagName !== "BUTTON") fileInput.click();
});
fileInput.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) setFile(file);
});

/* ========= PROCESS ========= */
async function processFile() {
    if (!selectedFile) return;
    processBtn.disabled = true;
    processBtnText.innerHTML = '<span class="spinner"></span> Memproses...';

    const formData = new FormData();
    formData.append("file", selectedFile);

    try {
        const res = await fetch("/api/pivot", { method: "POST", body: formData });
        const data = await res.json();
        if (!res.ok || !data.success) {
            throw new Error(data.error || "Gagal memproses file.");
        }
        renderResult(data);
        toast("File berhasil diproses!", "success");
    } catch (err) {
        toast(err.message, "error");
        console.error(err);
    } finally {
        processBtn.disabled = false;
        processBtnText.innerHTML = `
          <svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
            <path stroke-linecap="round" stroke-linejoin="round" d="M9.813 15.904L9 18.75l-.813-2.846a4.5 4.5 0 00-3.09-3.09L2.25 12l2.846-.813a4.5 4.5 0 003.09-3.09L9 5.25l.813 2.846a4.5 4.5 0 003.09 3.09L15.75 12l-2.846.813a4.5 4.5 0 00-3.09 3.09zM18.259 8.715L18 9.75l-.259-1.035a3.375 3.375 0 00-2.455-2.456L14.25 6l1.036-.259a3.375 3.375 0 002.455-2.456L18 2.25l.259 1.035a3.375 3.375 0 002.456 2.456L21.75 6l-1.035.259a3.375 3.375 0 00-2.456 2.456zM16.894 20.567L16.5 21.75l-.394-1.183a2.25 2.25 0 00-1.423-1.423L13.5 18.75l1.183-.394a2.25 2.25 0 001.423-1.423l.394-1.183.394 1.183a2.25 2.25 0 001.423 1.423l1.183.394-1.183.394a2.25 2.25 0 00-1.423 1.423z" />
          </svg>
          Proses Sekarang
        `;
    }
}

function renderResult(data) {
    const s = data.stats;

    const statsHtml = [
        { label: "Baris Sumber", value: s.totalSourceRows.toLocaleString("id-ID"), sub: "Total baris di file", icon: rowsIcon(), cls: "" },
        { label: "Dikelompokkan", value: s.groupedRows.toLocaleString("id-ID"), sub: "Baris valid diproses", icon: checkIcon(), cls: "success" },
        { label: "Skip", value: s.skippedRows.toLocaleString("id-ID"), sub: "Data tidak valid", icon: skipIcon(), cls: "warning" },
        { label: "Prefix Dibersihkan", value: s.cleanedPrefixCount.toLocaleString("id-ID"), sub: "PP- / P- dihapus", icon: broomIcon(), cls: "info" },
        { label: "Baris Pivot", value: s.pivotRowCount.toLocaleString("id-ID"), sub: "Hasil pengelompokan", icon: tableIcon(), cls: "success" },
        { label: "Total Qty", value: s.pivotTotalQty.toLocaleString("id-ID"), sub: "Sumber = Pivot ✓", icon: sumIcon(), cls: "success" },
    ];

    const statGrid = document.getElementById("statGrid");
    statGrid.innerHTML = statsHtml.map((stat, i) => `
        <div class="stat-card ${escapeAttr(stat.cls)}" style="animation-delay:${i * 60}ms">
          <div class="label">${stat.icon}${escape(stat.label)}</div>
          <div class="value">${escape(stat.value)}</div>
          <div class="sub">${escape(stat.sub)}</div>
        </div>
      `).join("");

    const tbody = document.getElementById("resultBody");
    tbody.innerHTML = data.preview.map((row, idx) => `
        <tr class="${idx % 2 === 1 ? "zebra" : ""}">
          <td>${idx + 1}</td>
          <td><strong>${escape(row.skuBatch)}</strong></td>
          <td>${escape(row.sku)}</td>
          <td>${escape(row.batch)}</td>
          <td class="qty-col">${row.qty.toLocaleString("id-ID")}</td>
        </tr>
      `).join("");

    document.getElementById("rowInfo").textContent =
        `Menampilkan ${data.preview.length} baris pertama dari total ${data.totalPivotRows.toLocaleString("id-ID")} baris pivot.`;

    lastDownloadUrl = data.downloadUrl;
    downloadBtn.onclick = () => {
        const a = document.createElement("a");
        a.href = data.downloadUrl;
        a.download = data.filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        toast("Mendownload file pivot...", "success");
    };

    resultCard.classList.remove("hidden");
    resultCard.scrollIntoView({ behavior: "smooth", block: "start" });
}

function resetAll() {
    removeFile();
    resultCard.classList.add("hidden");
    window.scrollTo({ top: 0, behavior: "smooth" });
    lastDownloadUrl = null;
}

/* ========= ICONS ========= */
function rowsIcon() {
    return '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5M3.75 17.25h16.5"/></svg>';
}
function checkIcon() {
    return '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12.75l6 6 9-13.5"/></svg>';
}
function skipIcon() {
    return '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126z"/></svg>';
}
function broomIcon() {
    return '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M16.023 9.348h4.992v-.001M2.985 19.644v-4.992m0 0h4.992m-4.993 0l3.181 3.183a8.25 8.25 0 0013.803-3.7M4.031 9.865a8.25 8.25 0 0113.803-3.7l3.181 3.182m0-4.991v4.99"/></svg>';
}
function tableIcon() {
    return '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M3.375 19.5h17.25m-17.25 0a1.125 1.125 0 01-1.125-1.125M3.375 19.5h7.5c.621 0 1.125-.504 1.125-1.125m-9.75 0V5.625m0 12.75v-1.5c0-.621.504-1.125 1.125-1.125m18.375 2.625V5.625m0 12.75c0 .621-.504 1.125-1.125 1.125m1.125-1.125v-1.5c0-.621-.504-1.125-1.125-1.125m0 3.75h-7.5A1.125 1.125 0 0112 18.375m9.75-12.75h-7.5c-.621 0-1.125.504-1.125 1.125m8.25-1.125c0 .621-.504 1.125-1.125 1.125m-1.125 0h.001m0 0h.001"/></svg>';
}
function sumIcon() {
    return '<svg fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M12 4.5v15m0 0l6.75-6.75M12 19.5L5.25 12.75"/></svg>';
}
