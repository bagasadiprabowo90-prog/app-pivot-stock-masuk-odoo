// ============================================================
//  App Pivot Stock Masuk Odoo — Web Server
//  Upload file Transfer (stock.picking) => Pivot SKU + Batch + Qty
// ============================================================

import express from "express";
import multer from "multer";
import ExcelJS from "exceljs";
import { randomUUID } from "node:crypto";
import path from "node:path";
import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

const UPLOADS_DIR = path.join(__dirname, "uploads");
const PUBLIC_DIR = path.join(__dirname, "public");

await fs.mkdir(UPLOADS_DIR, { recursive: true });
await fs.mkdir(PUBLIC_DIR, { recursive: true });

// Railway uses an ephemeral filesystem. Keep generated workbooks in memory
// for a short period instead of relying on a file that may disappear after a
// restart or redeploy.
const DOWNLOAD_TTL_MS = 10 * 60 * 1000;
const downloads = new Map();

// ---------- Multer ----------
const storage = multer.diskStorage({
    destination: UPLOADS_DIR,
    filename: (req, file, cb) => {
        const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
        const safe = file.originalname.replace(/[^a-zA-Z0-9._-]/g, "_");
        cb(null, `${ts}__${safe}`);
    },
});

const upload = multer({
    storage,
    limits: { fileSize: 50 * 1024 * 1024 }, // 50 MB
    fileFilter: (req, file, cb) => {
        if (/\.xlsx$/i.test(file.originalname)) cb(null, true);
        else cb(new Error("Hanya file .xlsx yang diperbolehkan."));
    },
});

// ---------- Middleware ----------
app.use(express.static(PUBLIC_DIR, { extensions: ["html"] }));
app.use(express.json());

// Cleanup helper
async function safeUnlink(p) {
    try { await fs.unlink(p); } catch { /* ignore */ }
}

function saveDownload(buffer, filename) {
    const token = randomUUID();
    downloads.set(token, {
        buffer,
        filename,
        expiresAt: Date.now() + DOWNLOAD_TTL_MS,
    });

    const timer = setTimeout(() => downloads.delete(token), DOWNLOAD_TTL_MS);
    timer.unref?.();
    return token;
}

// ---------- Routes ----------
app.get("/", (_req, res) => {
    res.sendFile(path.join(PUBLIC_DIR, "index.html"));
});

app.get("/api/health", (_req, res) => {
    res.json({ ok: true, time: new Date().toISOString() });
});

// Pivot endpoint
app.post("/api/pivot", upload.single("file"), async (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: "File tidak ditemukan." });
    }

    const uploadedPath = req.file.path;

    try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(uploadedPath);

        const sheet = wb.worksheets[0];
        if (!sheet) throw new Error("File Excel kosong atau rusak.");

        // Headers (row 1)
        const headers = [];
        const headerRow = sheet.getRow(1);
        const headerColCount = headerRow.cellCount || headerRow.values.length;
        for (let c = 1; c <= headerColCount; c++) {
            const v = headerRow.getCell(c).value;
            headers[c - 1] = v == null ? "" : String(v).trim();
        }

        const batchCol = headers.indexOf("Operations/Lot/Serial Number");
        const productCol = headers.indexOf("Operations/Product");
        const qtyCol = headers.indexOf("Operations/Qty Done");

        if (batchCol < 0 || productCol < 0 || qtyCol < 0) {
            throw new Error(
                "Kolom wajib tidak ditemukan. Pastikan file memiliki header: 'Operations/Lot/Serial Number', 'Operations/Product', dan 'Operations/Qty Done'."
            );
        }

        // Group rows
        const grouped = new Map();
        let sourceLineCount = 0;
        let sourceTotalQty = 0;
        let cleanedPrefixCount = 0;
        let skippedRowCount = 0;

        sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber === 1) return; // skip header
            const rawBatch = String(row.getCell(batchCol + 1).value ?? "").trim();
            const product = String(row.getCell(productCol + 1).value ?? "").trim();
            const qtyRaw = row.getCell(qtyCol + 1).value;
            const quantity = typeof qtyRaw === "number" ? qtyRaw : Number(qtyRaw);
            const skuMatch = product.match(/^\s*\[([^\]]+)\]/);

            if (!skuMatch || !Number.isFinite(quantity)) {
                skippedRowCount += 1;
                return;
            }

            // Batch kosong tetap diproses (dianggap string kosong),
            // supaya semua baris tanpa batch dari SKU yang sama digabung.
            const batch = rawBatch.replace(/^(?:PP-|P-)/i, "").trim();
            if (rawBatch && batch !== rawBatch) cleanedPrefixCount += 1;

            const sku = skuMatch[1].trim();
            const key = `${sku}\u0000${batch}`;
            grouped.set(key, (grouped.get(key) ?? 0) + quantity);
            sourceLineCount += 1;
            sourceTotalQty += quantity;
        });

        const pivotRows = [...grouped.entries()]
            .map(([key, quantity]) => {
                const [sku, batch] = key.split("\u0000");
                return [sku, batch, quantity];
            })
            .sort(
                (a, b) =>
                    a[0].localeCompare(b[0], "en", { numeric: true, sensitivity: "base" }) ||
                    a[1].localeCompare(b[1], "en", { numeric: true, sensitivity: "base" })
            );

        const pivotTotalQty = pivotRows.reduce((sum, r) => sum + r[2], 0);
        if (Math.abs(sourceTotalQty - pivotTotalQty) > 1e-9) {
            throw new Error(
                `Total Qty tidak cocok: sumber ${sourceTotalQty}, pivot ${pivotTotalQty}.`
            );
        }

        // Build output workbook
        const outWb = new ExcelJS.Workbook();
        outWb.creator = "App Pivot Stock Masuk Odoo";
        outWb.created = new Date();

        const outSheet = outWb.addWorksheet("Pivot SKU Batch", {
            views: [{ state: "frozen", ySplit: 1, showGridLines: false }],
        });

        // Header row
        outSheet.getRow(1).values = ["SKU Batch", "Qty"];
        outSheet.getRow(1).height = 28;
        const headerStyle = {
            fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } },
            font: { bold: true, color: { argb: "FFFFFFFF" }, size: 11, name: "Calibri" },
            alignment: { horizontal: "center", vertical: "middle" },
            border: {
                top: { style: "thin", color: { argb: "FF0B5D56" } },
                left: { style: "thin", color: { argb: "FF0B5D56" } },
                bottom: { style: "thin", color: { argb: "FF0B5D56" } },
                right: { style: "thin", color: { argb: "FF0B5D56" } },
            },
        };
        outSheet.getRow(1).eachCell((c) => Object.assign(c, headerStyle));

        // Data rows
        pivotRows.forEach(([sku, batch, qty], idx) => {
            const r = outSheet.getRow(idx + 2);
            r.values = [`${sku}${batch}`, qty];
            r.height = 20;
            const a = r.getCell(1);
            a.numFmt = "@";
            a.font = { color: { argb: "FF1F2937" }, size: 10 };
            a.alignment = { vertical: "middle" };

            const b = r.getCell(2);
            b.numFmt = "#,##0";
            b.font = { color: { argb: "FF1F2937" }, size: 10 };
            b.alignment = { vertical: "middle", horizontal: "right" };

            [a, b].forEach((c) => {
                c.border = { bottom: { style: "hair", color: { argb: "FFE5E7EB" } } };
            });

            // Zebra row banding
            if (idx % 2 === 1) {
                a.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
                b.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
            }
        });

        // Column widths
        outSheet.getColumn(1).width = 38;
        outSheet.getColumn(2).width = 16;

        // Generate output in memory so downloads work on Railway's ephemeral filesystem.
        const now = new Date();
        const dateLabel = now.toISOString().slice(0, 10);
        const outName = `Pivot Gabungan SKU Batch dan Qty - ${dateLabel}.xlsx`;
        const outBuffer = Buffer.from(await outWb.xlsx.writeBuffer());
        const downloadToken = saveDownload(outBuffer, outName);

        // Preview payload (first 100 rows)
        const preview = pivotRows.slice(0, 100).map(([sku, batch, qty]) => ({
            skuBatch: `${sku}${batch}`,
            sku,
            batch,
            qty,
        }));

        res.json({
            success: true,
            stats: {
                totalSourceRows: sheet.rowCount - 1,
                groupedRows: sourceLineCount,
                skippedRows: skippedRowCount,
                cleanedPrefixCount,
                pivotRowCount: pivotRows.length,
                sourceTotalQty,
                pivotTotalQty,
            },
            preview,
            totalPivotRows: pivotRows.length,
            filename: outName,
            downloadUrl: `/api/download/${downloadToken}`,
        });
    } catch (err) {
        console.error("[pivot error]", err);
        res.status(500).json({ error: err.message || "Terjadi kesalahan saat memproses file." });
    } finally {
        await safeUnlink(uploadedPath);
    }
});

// Download endpoint (short-lived in-memory result)
app.get("/api/download/:token", (req, res) => {
    const item = downloads.get(req.params.token);
    if (!item || item.expiresAt <= Date.now()) {
        downloads.delete(req.params.token);
        return res.status(404).json({ error: "File output sudah kedaluwarsa atau tidak ditemukan." });
    }

    res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
        "Content-Disposition",
        `attachment; filename*=UTF-8''${encodeURIComponent(item.filename)}`
    );
    res.send(item.buffer);
});

// Error handler
app.use((err, _req, res, _next) => {
    console.error("[server error]", err);
    res.status(500).json({ error: err.message || "Internal server error" });
});

app.listen(PORT, () => {
    console.log(`\n🚀 App Pivot Stock Masuk Odoo`);
    console.log(`   Berjalan di http://localhost:${PORT}\n`);
});
