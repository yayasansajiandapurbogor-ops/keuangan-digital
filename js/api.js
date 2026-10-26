'use strict';
/* ============================================================
   KEUANGAN DIGITAL — Lapisan API (js/api.js)
   Semua komunikasi ke Google Apps Script lewat fetch() JSON.
   POST memakai Content-Type text/plain agar TIDAK memicu CORS
   preflight (yang tidak didukung Apps Script).
   ============================================================ */
const CFG = window.KD_CONFIG || {};
const CFG_OK = /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(String(CFG.GAS_URL || '')) && !/GANTI/.test(String(CFG.GAS_URL));
const CLIENT_OK = /\.apps\.googleusercontent\.com$/.test(String(CFG.GOOGLE_CLIENT_ID || '')) && !/^GANTI/.test(String(CFG.GOOGLE_CLIENT_ID));

/** Sesi login (token dari server setelah verifikasi akun Google). */
const AUTH = {
  get token() { try { return localStorage.getItem('kd_sesi') || ''; } catch (e) { return ''; } },
  get user() { try { return JSON.parse(localStorage.getItem('kd_user') || 'null'); } catch (e) { return null; } },
  set(token, user) { try { localStorage.setItem('kd_sesi', token); localStorage.setItem('kd_user', JSON.stringify(user)); } catch (e) { /* abaikan */ } },
  clear() { try { localStorage.removeItem('kd_sesi'); localStorage.removeItem('kd_user'); } catch (e) { /* abaikan */ } }
};

/** Kirim body JSON ke Apps Script dan kembalikan objek respons apa adanya. */
async function apiRaw(body, timeoutMs = 60000) {
  if (!CFG_OK) throw Object.assign(new Error('GAS_URL di js/config.js belum diisi.'), { code: 'CONFIG' });
  const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
  const timer = ctrl ? setTimeout(() => ctrl.abort(), timeoutMs) : null;
  try {
    const res = await fetch(CFG.GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body),
      redirect: 'follow',
      signal: ctrl ? ctrl.signal : undefined
    });
    const txt = await res.text();
    try { return JSON.parse(txt); }
    catch (e) { throw Object.assign(new Error('Respons server bukan JSON. Pastikan Web App di-deploy "Execute as: Me" & "Anyone", lalu URL /exec benar.'), { code: 'NET' }); }
  } catch (err) {
    if (err.code) throw err;
    throw Object.assign(new Error(err.name === 'AbortError' ? 'Server terlalu lama merespons.' : 'Tidak dapat terhubung ke server (periksa internet).'), { code: 'NET' });
  } finally { if (timer) clearTimeout(timer); }
}

/**
 * Panggil action API yang butuh sesi. Resolve { success, data, message }.
 * Reject dengan err.code: 'AUTH' (sesi habis), 'NET' (jaringan), atau tanpa code (ditolak server).
 */
async function api(action, ...args) {
  const res = await apiRaw({ action, token: AUTH.token, args });
  if (res && res.success) return res;
  const err = new Error((res && res.message) || 'Terjadi kesalahan di server.');
  err.code = res && res.code;
  if (err.code === 'AUTH' && typeof onSesiBerakhir === 'function') onSesiBerakhir();
  throw err;
}

/** Info publik (nama/warna aplikasi) untuk layar login — GET sederhana tanpa preflight. */
async function apiInfo() {
  if (!CFG_OK) return null;
  try { const r = await fetch(CFG.GAS_URL + '?action=info'); const j = await r.json(); return j.success ? j.data : null; }
  catch (e) { return null; }
}
