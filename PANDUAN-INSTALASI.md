# Keuangan Digital v2 — Panduan Instalasi
**Backend:** Google Apps Script (REST API JSON) • **Frontend:** GitHub Pages • **Login:** akun Google (whitelist 2 akun)

```
Browser (HP/Laptop)
  └─ https://USERNAME.github.io/keuangan-digital/   ← frontend (folder ZIP ini)
        │  fetch() JSON  (+ token sesi setelah login Google)
        ▼
  Google Apps Script  /exec                          ← backend (Kode.gs)
        └─ Google Sheets • Drive • Gmail (pengingat)
```

Kerjakan **berurutan**. Ada 3 nilai yang saling terhubung — catat di Notepad:

| Nilai | Contoh | Didapat di |
|---|---|---|
| ① Alamat GitHub Pages | `https://budi.github.io/keuangan-digital/` | Tahap A |
| ② Google Client ID | `1234-abcd.apps.googleusercontent.com` | Tahap B |
| ③ URL Apps Script | `https://script.google.com/macros/s/AKfy…/exec` | Tahap C |

---

## Tahap A — Tentukan alamat GitHub Pages

1. Daftar/masuk di **https://github.com**. Username Anda menjadi bagian alamat situs.
2. Nama repository yang disarankan: **`keuangan-digital`**.
3. Alamat situs nanti: `https://USERNAME.github.io/keuangan-digital/` → catat sebagai **①**.
   *Origin*-nya (dipakai di Tahap B) adalah `https://USERNAME.github.io` — tanpa garis miring & tanpa nama repo.

## Tahap B — Buat Google Client ID (untuk tombol "Masuk dengan Google")

1. Buka **https://console.cloud.google.com** dengan akun Google suami/pemilik → buat **Project baru** (nama bebas, mis. `Keuangan Digital`).
2. Menu **Google Auth Platform** (dulu: *APIs & Services → OAuth consent screen*) → **Get started**:
   - App name: `Keuangan Digital`, User support email: email Anda
   - Audience: **External**
   - Contact email: email Anda → **Create**
3. Menu **Audience** → bagian **Test users** → **Add users** → masukkan **kedua email** (suami & istri) → Save.
   *(Mode "Testing" justru cocok: hanya akun yang didaftarkan yang bisa memakai tombol login.)*
4. Menu **Clients** (dulu: *Credentials → Create credentials → OAuth client ID*) → **Create client**:
   - Application type: **Web application**
   - Name: `Keuangan Digital Web`
   - **Authorized JavaScript origins** → Add URI → `https://USERNAME.github.io`
   - Authorized redirect URIs: **kosongkan**
   - **Create** → salin **Client ID** → catat sebagai **②**.

> Perubahan origin bisa butuh 5 menit sampai beberapa jam sebelum aktif.

## Tahap C — Pasang backend (Google Apps Script)

1. Buka **https://script.google.com** → **Proyek baru** → beri nama `Keuangan Digital API`.
   *(Sudah memasang versi 1? Pakai proyek yang sama: ganti isi Kode.gs, hapus file `Index`, `Stylesheet`, `JavaScript`. Data di spreadsheet tetap aman.)*
2. ⚙️ **Setelan Proyek** → Zona waktu **(GMT+07:00) Jakarta**.
3. Ganti isi `Code.gs` dengan **Kode.gs** (rename menjadi `Kode`). Di bagian **KONFIG** isi:
   ```js
   PENGGUNA: [
     { email: '', nama: 'Budi', label: 'Suami' },              // '' = akun Anda (pemilik)
     { email: 'email.istri@gmail.com', nama: 'Siti', label: 'Istri' }
   ],
   GOOGLE_CLIENT_ID: '②',
   FRONTEND_URL: '①',
   ```
4. Pilih fungsi **`setupAppEnvironment`** → **▶ Jalankan** → izinkan akses (jika muncul "Google belum memverifikasi aplikasi ini": **Lanjutan → Buka … (tidak aman)** — normal untuk script pribadi).
   Cek **Log eksekusi**: harus ada ✅ Setup selesai, 🔐 Client ID, folder & spreadsheet. Setup **aman dijalankan ulang** (tidak menduplikasi data).
5. **Terapkan → Deployment baru → ⚙️ Aplikasi web**:

   | Pengaturan | Pilih |
   |---|---|
   | Jalankan sebagai | **Saya (Me)** |
   | Yang memiliki akses | **Siapa saja (Anyone)** |

   **Deploy** → salin URL berakhiran `/exec` → catat sebagai **③**.

> "Anyone" di sini aman: seluruh data hanya bisa dibaca/diubah dengan **token sesi** yang diberikan server setelah login Google terverifikasi dan email cocok dengan whitelist. Tanpa login, API hanya mengembalikan nama & warna aplikasi.

## Tahap D — Siapkan frontend

1. Ekstrak **keuangan-digital.zip**. Folder hasil ekstrak (mis. `Documents\keuangan-digital`) **adalah folder repository** — isinya harus langsung:
   ```
   keuangan-digital\
   ├── index.html        ← harus terlihat di sini
   ├── README.md
   ├── PANDUAN-INSTALASI.md
   ├── css\style.css
   └── js\config.js, api.js, app.js
   ```
2. Buka **`js\config.js`** dengan Notepad, isi:
   ```js
   GAS_URL: '③',
   GOOGLE_CLIENT_ID: '②'
   ```
   Simpan.

## Tahap E — Upload ke GitHub & aktifkan Pages (lewat terminal)

> ⚠️ Jangan memakai "Upload files" di web GitHub — struktur folder `css/` & `js/` bisa rusak dan situs tampil tanpa tampilan.

1. Install Git: **https://git-scm.com/download/win** (pengaturan default). Buka **PowerShell**, cek: `git --version`.
2. Identitas (sekali saja):
   ```powershell
   git config --global user.name "Nama Anda"
   git config --global user.email "email-akun-github@gmail.com"
   ```
3. Di github.com: **+ → New repository** → nama `keuangan-digital` → **Public** → **jangan** centang README/.gitignore/license → Create.
4. Masuk ke folder hasil ekstrak (cara cepat: buka foldernya di File Explorer, klik address bar, ketik `powershell`, Enter). Cek dulu:
   ```powershell
   dir
   ```
   Baris yang tampil **wajib** berisi `index.html`, `css`, `js`. Kalau yang terlihat malah folder `keuangan-digital`, masuk dulu: `cd keuangan-digital`.
5. Kirim ke GitHub — satu per satu:
   ```powershell
   git init
   git add .
   git commit -m "Upload pertama"
   git branch -M main
   git remote add origin https://github.com/USERNAME/keuangan-digital.git
   git push -u origin main
   ```
   Saat diminta **Password**, tempel **Personal Access Token** (bukan password GitHub): buat di **https://github.com/settings/tokens** → *Generate new token (classic)* → centang **repo** → salin `ghp_…`. Saat menempel, layar memang **tidak menampilkan apa pun** — tetap tekan Enter.
6. Repository → **Settings → Pages** → Source **Deploy from a branch** → Branch **main / (root)** → Save. Centang **Enforce HTTPS**.
7. Tunggu 1–2 menit → buka **①** → klik **Masuk dengan Google**.

## Tahap F — Mulai pakai

1. Suami masuk → dashboard tampil (berisi data contoh).
2. Kirim alamat **①** ke istri → masuk dengan akun Google-nya.
3. Setelah mencoba: **Pengaturan → Data & Cadangan → Hapus Data Contoh**, lalu buat dompet asli beserta saldo awal.
4. Di HP: Chrome → ⋮ → **Tambahkan ke layar utama**.

## Memperbarui

- **Frontend** (setelah mengubah file): dari folder yang sama
  ```powershell
  git add .
  git commit -m "Update"
  git push
  ```
  Lalu Ctrl+Shift+R di browser.
- **Backend:** Terapkan → **Kelola deployment** → ✏️ → Versi: **Versi baru** → Deploy. URL **③** tetap sama, `config.js` tidak perlu diubah.

---

## Pemecahan Masalah

| Gejala | Solusi |
|---|---|
| Layar "Konfigurasi Belum Lengkap" | `js/config.js` belum diisi/belum di-push. Isi ② & ③, lalu `git add . / commit / push` |
| Tombol Google error *origin_mismatch* / "origin is not allowed" | Di Google Cloud → Clients → tambahkan origin `https://USERNAME.github.io` (tanpa `/` di akhir). Tunggu ±5 menit |
| "Access blocked: app has not completed verification" | Tambahkan kedua email di **Audience → Test users** |
| "GOOGLE_CLIENT_ID belum diisi di KONFIG" | Isi KONFIG di Kode.gs → jalankan `setupAppEnvironment` lagi |
| "Token bukan untuk aplikasi ini" | Client ID di `config.js` dan di Kode.gs berbeda — samakan |
| "Respons server bukan JSON" / gagal terhubung | Deploy ulang dengan **Execute as: Me** + **Anyone**; pastikan URL ③ berakhiran `/exec` (bukan `/dev`) |
| "Akses Tidak Diizinkan" | Email belum ada di whitelist — tambahkan di KONFIG lalu jalankan setup, atau lewat Pengaturan → Pengguna Terdaftar |
| Situs 404 | Di repo GitHub, `index.html` harus terlihat di halaman depan. Kalau yang terlihat folder lain, `git init` dijalankan di folder yang salah |
| Tampilan polos tanpa warna | Folder `css/` & `js/` tidak ikut — gunakan terminal, bukan "Upload files" |
| Data diubah langsung di spreadsheet tidak muncul | Pengaturan → Data & Cadangan → **Bersihkan Cache Server** |
| Masih tampil versi lama | Ctrl+Shift+R atau buka jendela Incognito |

**Keamanan singkat:** token sesi berlaku 14 hari dan disimpan di server dalam bentuk hash; menghapus email dari whitelist langsung mencabut aksesnya; tombol **Keluar** menghapus sesi dan data cache di perangkat.
