// scripts/test-skip-rule-inproc.mjs
// Validate the new skip rules: batch kosong tetap diproses, hanya !sku atau !qty yang skip.
//
// Fixture:
//   SKU001 + BATCH-A1             -> 10 (no prefix)
//   SKU001 + P-BATCH-B1           -> 5  (P- prefix stripped)
//   SKU001 + PP-BATCH-B2          -> 3  (PP- prefix stripped)
//   SKU002 + ""                   -> 7  (batch KOSONG, diproses -- bukti rule baru)
//   SKU002 + null                 -> 4  (batch null=KOSONG, digabung jadi 1 key)
//   SKU003-no-brackets            -> skip (no [SKU] regex match)
//   SKU999-Test qty='bukan-angka' -> skip (NaN)
//
// Hasil yang diharapkan:
//   groupedRows = 5  (3 SKU001 + 2 SKU002)
//   pivotRowCount = 4 (BATCH-A1, BATCH-B1, BATCH-B2, "")
//   skippedRows = 2
//   cleanedPrefixCount = 2 (P-BATCH-B1 & PP-BATCH-B2 kena strip)
//   pivotTotalQty = sourceTotalQty
import ExcelJS from "exceljs";
import fs from "node:fs/promises";

const XLSX = "uploads/_test_skip_rule.xlsx";
await fs.mkdir("uploads", { recursive: true });

const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet("Sheet1");
ws.addRow([
    "Operations/Product",
    "Operations/Lot/Serial Number",
    "Operations/Qty Done",
]);
ws.addRow(["[SKU001] Paracetamol 500mg", "BATCH-A1", 10]);
ws.addRow(["[SKU001] Paracetamol 500mg", "P-BATCH-B1", 5]);
ws.addRow(["[SKU001] Paracetamol 500mg", "PP-BATCH-B2", 3]);
ws.addRow(["[SKU002] Amoxicillin 500mg", "", 7]);
ws.addRow(["[SKU002] Amoxicillin 500mg", null, 4]);
ws.addRow(["SKU-no-brackets", "BATCH-X", 99]);        // harus skip (no [SKU])
ws.addRow(["[SKU999] Test", "BATCH-E", "bukan-angka"]); // harus skip (NaN)
await wb.xlsx.writeFile(XLSX);

// --- Replicate server logic exactly ---
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.readFile(XLSX);
const sheet = wb2.worksheets[0];
const headers = [];
const hc = sheet.getRow(1).cellCount || sheet.getRow(1).values.length;
for (let c = 1; c <= hc; c++) {
    headers[c - 1] = String(sheet.getRow(1).getCell(c).value ?? "").trim();
}
const batchCol = headers.indexOf("Operations/Lot/Serial Number");
const productCol = headers.indexOf("Operations/Product");
const qtyCol = headers.indexOf("Operations/Qty Done");

const grouped = new Map();
let groupedRows = 0, sourceTotalQty = 0, cleanedPrefixCount = 0, skippedRows = 0;

sheet.eachRow({ includeEmpty: false }, (row, n) => {
    if (n === 1) return;
    const rawBatch = String(row.getCell(batchCol + 1).value ?? "").trim();
    const product = String(row.getCell(productCol + 1).value ?? "").trim();
    const qtyRaw = row.getCell(qtyCol + 1).value;
    const quantity = typeof qtyRaw === "number" ? qtyRaw : Number(qtyRaw);
    const skuMatch = product.match(/^\s*\[([^\]]+)\]/);

    if (!skuMatch || !Number.isFinite(quantity)) {
        skippedRows += 1;
        return;
    }
    const batch = rawBatch.replace(/^(?:PP-|P-)/i, "").trim();
    if (rawBatch && batch !== rawBatch) cleanedPrefixCount += 1;
    const sku = skuMatch[1].trim();
    const key = `${sku}\u0000${batch}`;
    grouped.set(key, (grouped.get(key) ?? 0) + quantity);
    groupedRows += 1;
    sourceTotalQty += quantity;
});

const pivotRows = [...grouped.entries()]
    .map(([k, q]) => {
        const [sku, batch] = k.split("\u0000");
        return { sku, batch, qty: q };
    })
    .sort((a, b) =>
        a.sku.localeCompare(b.sku, "en", { numeric: true }) ||
        a.batch.localeCompare(b.batch, "en", { numeric: true })
    );

const pivotTotalQty = pivotRows.reduce((s, r) => s + r.qty, 0);

const checks = [
    ["groupedRows", groupedRows, 5],
    ["pivotRowCount", pivotRows.length, 4],
    ["skippedRows", skippedRows, 2],
    ["cleanedPrefixCount", cleanedPrefixCount, 2],
    ["pivotTotalQty", pivotTotalQty, sourceTotalQty],
];

console.log("== TEST RESULTS ==");
let failed = 0;
for (const [k, got, exp] of checks) {
    const ok = got === exp;
    console.log(`  ${ok ? "OK  " : "FAIL"} ${k}: got=${got} expected=${exp}`);
    if (!ok) failed += 1;
}
console.log("\n== PIVOT ROWS ==");
for (const r of pivotRows) console.log(`  sku=[${r.sku}] batch=[${r.batch}] qty=${r.qty}`);

if (failed === 0) {
    console.log("\nALL TESTS PASSED");
    process.exit(0);
} else {
    console.log(`\n${failed} TESTS FAILED`);
    process.exit(1);
}
