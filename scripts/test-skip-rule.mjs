// scripts/test-skip-rule.mjs
// Fixture generator & smoke-test untuk aturan skip batch.
//
// Skenario:
//   A) SKU + Batch valid + qty valid -> proses (prefix dibersihkan)
//   B) SKU + batch KOSONG (current change: harus diproses, key = SKU+empty batch)
//   C) SKU tanpa [SKU] di product -> skip
//   D) qty invalid (NaN) -> skip
import ExcelJS from "exceljs";
import fs from "node:fs/promises";

const XLSX = "uploads/_test_skip_rule.xlsx";

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
ws.addRow(["[SKU003] Vitamin C Tanpa SKU", "BATCH-D", 99]);
ws.addRow(["[SKU999] Test", "BATCH-E", "bukan-angka"]);

await fs.mkdir("uploads", { recursive: true });
await wb.xlsx.writeFile(XLSX);
console.log("fixture ready:", XLSX);
