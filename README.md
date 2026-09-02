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

## Menjalankan secara lokal dengan Streamlit

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Streamlit akan menampilkan alamat aplikasi di browser.

Untuk mode development dengan auto-restart:

```bash
npm run dev
```

## Deploy ke Streamlit Community Cloud

1. Push perubahan repository ke GitHub.
2. Buka <https://share.streamlit.io> dan pilih **Create app**.
3. Pilih repository, branch `main`, lalu set **Main file path** menjadi `app.py`.
4. Klik **Deploy**.

Konfigurasi ukuran upload dibatasi menjadi 10 MB melalui `.streamlit/config.toml`.

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
