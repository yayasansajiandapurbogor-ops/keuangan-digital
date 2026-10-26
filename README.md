# Keuangan Digital

Aplikasi web keuangan keluarga — catat transaksi, dompet, budget, target tabungan, tagihan dengan pengingat email, serta laporan Excel/PDF.

- **Frontend:** HTML/CSS/JS statis (GitHub Pages) — folder ini
- **Backend:** Google Apps Script REST API (`Kode.gs`, dipasang terpisah di script.google.com)
- **Login:** Masuk dengan Google, hanya untuk 2 akun terdaftar

## Struktur

```
index.html        Kerangka aplikasi + layar login
css/style.css     Desain "Keuangan Digital Emerald"
js/config.js      ← ISI: GAS_URL & GOOGLE_CLIENT_ID
js/api.js         fetch() ke Apps Script + sesi login
js/app.js         Seluruh halaman (SPA), sinkronisasi, grafik, laporan
```

Panduan lengkap: **PANDUAN-INSTALASI.md**.
