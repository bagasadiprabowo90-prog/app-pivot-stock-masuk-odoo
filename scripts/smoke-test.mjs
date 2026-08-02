import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";
import fs from "node:fs/promises";
import ExcelJS from "exceljs";

const ROOT = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const PORT = 3137;
const BASE_URL = `http://127.0.0.1:${PORT}`;
const fixturePath = path.join(ROOT, "uploads", "_smoke_test.xlsx");

async function createFixture() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    sheet.addRow([
        "Operations/Product",
        "Operations/Lot/Serial Number",
        "Operations/Qty Done",
    ]);
    sheet.addRow(["[SKU001] Product A", "BATCH-A", 10]);
    sheet.addRow(["[SKU001] Product A", "P-BATCH-A", 2]);
    sheet.addRow(["[SKU002] Product B", "", 7]);
    sheet.addRow(["Product without SKU", "BATCH-X", 9]);
    sheet.addRow(["[SKU003] Invalid Qty", "BATCH-Y", "not-a-number"]);
    await workbook.xlsx.writeFile(fixturePath);
}

async function waitForHealth() {
    for (let attempt = 0; attempt < 30; attempt += 1) {
        try {
            const response = await fetch(`${BASE_URL}/api/health`);
            if (response.ok) return;
        } catch {
            // Server is still starting.
        }
        await new Promise((resolve) => setTimeout(resolve, 200));
    }
    throw new Error("Server tidak merespons health check.");
}

function assert(condition, message) {
    if (!condition) throw new Error(message);
}

const child = spawn(process.execPath, ["server.js"], {
    cwd: ROOT,
    env: { ...process.env, PORT: String(PORT) },
    stdio: ["ignore", "pipe", "pipe"],
});

let serverLogs = "";
child.stdout.on("data", (chunk) => { serverLogs += chunk.toString(); });
child.stderr.on("data", (chunk) => { serverLogs += chunk.toString(); });

try {
    await createFixture();
    await waitForHealth();

    const form = new FormData();
    form.append(
        "file",
        new Blob([await fs.readFile(fixturePath)], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        "smoke-test.xlsx"
    );

    const pivotResponse = await fetch(`${BASE_URL}/api/pivot`, {
        method: "POST",
        body: form,
    });
    const result = await pivotResponse.json();

    assert(pivotResponse.ok && result.success, result.error || "Pivot request gagal.");
    assert(result.stats.totalSourceRows === 5, "totalSourceRows tidak sesuai.");
    assert(result.stats.groupedRows === 3, "groupedRows tidak sesuai.");
    assert(result.stats.skippedRows === 2, "skippedRows tidak sesuai.");
    assert(result.stats.cleanedPrefixCount === 1, "cleanedPrefixCount tidak sesuai.");
    assert(result.stats.pivotRowCount === 2, "pivotRowCount tidak sesuai.");
    assert(result.stats.sourceTotalQty === 19, "sourceTotalQty tidak sesuai.");
    assert(result.stats.pivotTotalQty === 19, "pivotTotalQty tidak sesuai.");
    assert(result.preview.some((row) => row.sku === "SKU002" && row.batch === "" && row.qty === 7),
        "Batch kosong tidak ikut diproses.");

    const downloadResponse = await fetch(`${BASE_URL}${result.downloadUrl}`);
    const downloadBuffer = Buffer.from(await downloadResponse.arrayBuffer());
    assert(downloadResponse.ok, "Download output gagal.");
    assert(downloadResponse.headers.get("content-type")?.includes("spreadsheetml.sheet"),
        "Content-Type download tidak sesuai.");
    assert(downloadBuffer.subarray(0, 2).toString() === "PK", "Output bukan file XLSX valid.");

    console.log("HTTP SMOKE TEST PASSED");
    console.log(JSON.stringify({ stats: result.stats, downloadBytes: downloadBuffer.length }, null, 2));
} catch (error) {
    console.error("HTTP SMOKE TEST FAILED:", error.message);
    if (serverLogs) console.error(serverLogs);
    process.exitCode = 1;
} finally {
    await fs.rm(fixturePath, { force: true });
    child.kill();
}
