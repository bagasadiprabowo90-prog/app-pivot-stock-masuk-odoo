# Pivot Stock Masuk Odoo

Aplikasi web untuk mengubah file Transfer/`stock.picking` dari Odoo menjadi
pivot **SKU + Batch + Qty**.

## Fitur

- Upload file `.xlsx` dengan batas ukuran 50 MB.
- Membaca kolom:
  - `Operations/Product`
  - `Operations/Lot/Serial Number`
  - `Operations/Qty Done`
- Mengambil SKU dari format produk seperti `[SKU001] Nama Produk`.
- Menghapus prefix batch `P-` dan `PP-`.
- Batch kosong tetap diproses dan digabung berdasarkan SKU.
- Menampilkan statistik, preview hasil, dan mengunduh file Excel pivot.

## Menjalankan secara lokal

```bash
npm install
npm start
```

Buka <http://localhost:3000>.

Untuk mode development dengan auto-restart:

```bash
npm run dev
```

## Deploy ke Railway

Project ini adalah aplikasi Node.js/Express, bukan Streamlit/Python.
Railway akan mendeteksi `package.json` dan menjalankan `npm start`.

1. Push repository ini ke GitHub.
2. Di Railway pilih **New Project → Deploy from GitHub repo**.
3. Pilih repository ini.
4. Tunggu build selesai, lalu pilih **Settings → Networking → Generate Domain**.
5. Buka domain yang dibuat Railway.

Konfigurasi deployment sudah tersedia di `railway.json` dan `Procfile`. Railway
menyediakan `PORT` secara otomatis; server membaca nilai tersebut.

### Catatan filesystem

File upload hanya dipakai selama proses dan dihapus setelah selesai. File hasil
Excel disimpan sementara di memory selama 10 menit menggunakan token download,
sehingga aplikasi tidak membutuhkan Railway Volume. Jika service restart atau
token kedaluwarsa, proses upload perlu diulang.

## Struktur input

Baris header harus memiliki nama kolom persis seperti berikut:

```text
Operations/Product
Operations/Lot/Serial Number
Operations/Qty Done
```

Baris akan di-skip hanya jika produk tidak memiliki `[SKU]` di awal atau Qty
bukan angka yang valid. Batch kosong **tidak** menyebabkan baris di-skip.

## Validasi lokal

```bash
node --check server.js
node scripts/test-skip-rule-inproc.mjs
node scripts/smoke-test.mjs
```
