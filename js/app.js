'use strict';
/* ============================================================
   KEUANGAN DIGITAL — Frontend (JavaScript.html)
   Prinsip gas-instant-ux:
   1) SPA: semua halaman sudah ada; navigasi = tampil/sembunyi (0 ms)
   2) Optimistic UI: perubahan langsung tampil, disimpan ke antrean
      (localStorage) lalu dikirim ke server di latar belakang
   3) Data awal disisipkan saat doGet → tanpa round-trip tambahan
   4) Batch: banyak perubahan dikirim dalam satu panggilan sinkronkan()
   ============================================================ */

// ════════════════════════════════════════════════════════════
// 1. UTILITAS
// ════════════════════════════════════════════════════════════
const TZ = 'Asia/Jakarta';
const CHART_URL = 'https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js';
const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ic = (n, cls = '') => `<span class="ms ${cls}" aria-hidden="true">${n}</span>`;
const NF = new Intl.NumberFormat('id-ID', { maximumFractionDigits: 0 });
const fmtN = (n) => NF.format(Math.round(+n || 0));
const rp = (n, sign) => {
  const v = Math.round(+n || 0);
  const s = 'Rp ' + NF.format(Math.abs(v));
  if (v < 0) return '-' + s;
  return sign && v > 0 ? '+' + s : s;
};
const rpS = (n) => {
  const v = +n || 0, a = Math.abs(v), neg = v < 0 ? '-' : '';
  const f = (x, d) => x.toFixed(d).replace('.', ',').replace(/,0+$/, '').replace(/(,\d*?)0+$/, '$1');
  if (a >= 1e9) return `${neg}Rp ${f(a / 1e9, 2)}M`;
  if (a >= 1e6) return `${neg}Rp ${f(a / 1e6, a >= 1e8 ? 0 : a >= 1e7 ? 1 : 2)}jt`;
  if (a >= 1e3) return `${neg}Rp ${f(a / 1e3, 0)}rb`;
  return `${neg}Rp ${fmtN(a)}`;
};
const pct = (x, d = 0) => (isFinite(x) ? (x * 100).toFixed(d).replace('.', ',') : '0') + '%';
const clamp = (x, a, b) => Math.max(a, Math.min(b, x));
const uid = (p) => `${p}_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
const parseNum = (s) => { const n = Number(String(s ?? '').replace(/[^\d]/g, '')); return isFinite(n) ? n : 0; };
const sum = (arr, f) => arr.reduce((s, x) => s + (+f(x) || 0), 0);
const short = (s, n = 18) => (String(s).length > n ? String(s).slice(0, n - 1) + '…' : String(s));
const RE_EMAIL = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/;

// Tanggal (selalu zona Asia/Jakarta, format string YYYY-MM-DD)
const BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const BLN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
const _fISO = new Intl.DateTimeFormat('en-CA', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit' });
const _fHM = new Intl.DateTimeFormat('en-GB', { timeZone: TZ, hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
const today = () => _fISO.format(new Date());
const nowHMS = () => _fHM.format(new Date()).replace(/^24/, '00');
const nowStamp = () => `${today()} ${nowHMS()}`;
const jamJkt = () => +nowHMS().slice(0, 2);
const pD = (s) => { const [y, m, d] = String(s).split('-').map(Number); return new Date(Date.UTC(y, (m || 1) - 1, d || 1)); };
const isoD = (d) => d.toISOString().slice(0, 10);
const addDays = (s, n) => { const d = pD(s); d.setUTCDate(d.getUTCDate() + n); return isoD(d); };
const addMonths = (s, n) => {
  const [y, m, d] = s.split('-').map(Number);
  const t = new Date(Date.UTC(y, m - 1 + n, 1));
  const last = new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 0)).getUTCDate();
  t.setUTCDate(Math.min(d || 1, last));
  return isoD(t);
};
const diffDays = (a, b) => Math.round((pD(b) - pD(a)) / 864e5);
const mKey = (s) => String(s).slice(0, 7);
const mEnd = (mk) => { const [y, m] = mk.split('-').map(Number); return isoD(new Date(Date.UTC(y, m, 0))); };
const mDays = (mk) => +mEnd(mk).slice(8);
const mAdd = (mk, n) => addMonths(mk + '-01', n).slice(0, 7);
const mLabel = (mk, short) => { const [y, m] = mk.split('-').map(Number); return `${(short ? BLN : BULAN)[m - 1]} ${y}`; };
const validDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(s || '') && isoD(pD(s)) === s;
const tgl = (s, mode = 'short') => {
  if (!s) return '-';
  const d = pD(s), D = d.getUTCDate(), M = d.getUTCMonth(), Y = d.getUTCFullYear();
  if (mode === 'long') return `${D} ${BULAN[M]} ${Y}`;
  if (mode === 'day') return `${HARI[d.getUTCDay()]}, ${D} ${BULAN[M]} ${Y}`;
  if (mode === 'dm') return `${D} ${BLN[M]}`;
  if (mode === 'num') return `${String(D).padStart(2, '0')}/${String(M + 1).padStart(2, '0')}/${Y}`;
  return `${D} ${BLN[M]} ${Y}`;
};
const jam = (stamp) => String(stamp || '').slice(11, 16);
const relHari = (s) => {
  const d = diffDays(today(), s);
  if (d === 0) return 'Hari ini';
  if (d === 1) return 'Besok';
  if (d === -1) return 'Kemarin';
  return d > 1 ? `${d} hari lagi` : `${-d} hari lalu`;
};

/** Angka → kata (Rupiah), contoh 650000 → "Enam ratus lima puluh ribu rupiah" */
function terbilang(n) {
  n = Math.floor(Math.abs(+n || 0));
  if (!n) return '';
  const s = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas'];
  const t = (x) => {
    if (x < 12) return s[x];
    if (x < 20) return t(x - 10) + ' belas';
    if (x < 100) return t(Math.floor(x / 10)) + ' puluh ' + t(x % 10);
    if (x < 200) return 'seratus ' + t(x - 100);
    if (x < 1000) return t(Math.floor(x / 100)) + ' ratus ' + t(x % 100);
    if (x < 2000) return 'seribu ' + t(x - 1000);
    if (x < 1e6) return t(Math.floor(x / 1000)) + ' ribu ' + t(x % 1000);
    if (x < 1e9) return t(Math.floor(x / 1e6)) + ' juta ' + t(x % 1e6);
    if (x < 1e12) return t(Math.floor(x / 1e9)) + ' miliar ' + t(x % 1e9);
    return t(Math.floor(x / 1e12)) + ' triliun ' + t(x % 1e12);
  };
  const r = (t(n) + ' rupiah').replace(/\s+/g, ' ').trim();
  return r.charAt(0).toUpperCase() + r.slice(1);
}

// Warna
const hexRgb = (h) => { h = String(h || '#000000').replace('#', ''); if (h.length === 3) h = h.split('').map((c) => c + c).join(''); return [0, 2, 4].map((i) => parseInt(h.substr(i, 2), 16) || 0); };
const rgbHex = (a) => '#' + a.map((v) => clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0')).join('').toUpperCase();
const mix = (a, b, t) => { const A = hexRgb(a), B = hexRgb(b); return rgbHex(A.map((v, i) => v + (B[i] - v) * t)); };
const rgba = (h, a) => `rgba(${hexRgb(h).join(',')},${a})`;
const lum = (h) => { const c = hexRgb(h).map((v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); }); return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]; };
const isDark = () => document.documentElement.dataset.theme === 'dark';
const fg = (hex) => (isDark() ? mix(hex, '#FFFFFF', 0.28) : lum(hex) > 0.38 ? mix(hex, '#000000', 0.38) : hex);
const tint = (hex) => `background:${rgba(hex, isDark() ? 0.18 : 0.12)};color:${fg(hex)}`;
const cssVar = (n) => getComputedStyle(document.documentElement).getPropertyValue(n).trim();

// localStorage aman
const LS = {
  get(k, d) { try { const v = localStorage.getItem(k); return v === null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* abaikan */ } },
  del(k) { try { localStorage.removeItem(k); } catch (e) { /* abaikan */ } }
};

/** Memuat script CDN secara malas (Chart.js, SheetJS). */
const _scripts = {};
function loadScript(src) {
  if (_scripts[src]) return _scripts[src];
  _scripts[src] = new Promise((ok, fail) => {
    const s = document.createElement('script');
    s.src = src; s.async = true;
    s.onload = () => ok(); s.onerror = () => { delete _scripts[src]; fail(new Error('Gagal memuat ' + src)); };
    document.head.appendChild(s);
  });
  return _scripts[src];
}


// ════════════════════════════════════════════════════════════
// 2. STATE (sumber data tunggal di browser)
// ════════════════════════════════════════════════════════════
const TX_COLS = ['id', 'tanggal', 'tipe', 'dompet_id', 'dompet_tujuan_id', 'kategori_id', 'nominal', 'catatan', 'dibuat_oleh', 'dibuat_pada', 'ref'];
const S = {
  me: null, users: [], settings: {}, dompet: [], kategori: [], tx: [], budget: [], target: [], tagihan: [],
  txDari: '', versi: 0, meta: {}, rev: 0,
  dMap: new Map(), kMap: new Map(), uMap: new Map(), txMap: new Map()
};
let MEMO = {};
const memo = (k, fn) => (k in MEMO ? MEMO[k] : (MEMO[k] = fn()));

function loadBoot(b) {
  S.me = b.me; S.me.email = String(S.me.email).toLowerCase();
  S.users = (b.users || []).map((u) => ({ ...u, email: String(u.email).toLowerCase() }));
  if (!S.users.length) S.users = [S.me];
  S.settings = b.pengaturan || {};
  S.dompet = (b.dompet || []).map((d) => ({ ...d, saldo: +d.saldo || 0, saldo_awal: +d.saldo_awal || 0, urutan: +d.urutan || 0, aktif: d.aktif !== false }));
  S.kategori = b.kategori || [];
  S.budget = (b.budget || []).map((x) => ({ ...x, batas_nominal: +x.batas_nominal || 0 }));
  S.target = (b.target || []).map((x) => ({ ...x, target_nominal: +x.target_nominal || 0, terkumpul: +x.terkumpul || 0 }));
  S.tagihan = (b.tagihan || []).map((x) => ({ ...x, nominal: +x.nominal || 0 }));
  const c = b.tx.c;
  S.tx = b.tx.r.map((r) => { const o = {}; c.forEach((k, i) => (o[k] = r[i])); o.nominal = +o.nominal || 0; return o; });
  S.txDari = b.txDari;
  S.versi = b.versi;
  S.meta = b.meta || {};
  reindex();
  sortTx();
  bump();
}
function reindex() {
  S.dMap = new Map(S.dompet.map((d) => [d.id, d]));
  S.kMap = new Map(S.kategori.map((k) => [k.id, k]));
  S.uMap = new Map(S.users.map((u) => [u.email, u]));
  S.txMap = new Map(S.tx.map((t) => [t.id, t]));
}
function sortTx() {
  S.tx.sort((a, b) => (a.tanggal < b.tanggal ? 1 : a.tanggal > b.tanggal ? -1 : String(b.dibuat_pada).localeCompare(String(a.dibuat_pada))));
}
function bump() { S.rev++; MEMO = {}; }

// Lookup & label
const dompetAktif = () => memo('dpa', () => S.dompet.filter((d) => d.aktif).sort((a, b) => (a.urutan || 99) - (b.urutan || 99) || a.nama.localeCompare(b.nama)));
const dompetSemua = () => memo('dps', () => S.dompet.slice().sort((a, b) => (b.aktif - a.aktif) || (a.urutan || 99) - (b.urutan || 99)));
const totalSaldo = () => sum(dompetAktif(), (d) => d.saldo);
const katByTipe = (tipe) => S.kategori.filter((k) => k.tipe === tipe);
const userOf = (email) => S.uMap.get(String(email || '').toLowerCase()) || { email, nama: String(email || '?').split('@')[0], label: '' };
const userIdx = (email) => S.users.findIndex((u) => u.email === String(email || '').toLowerCase());
const familyName = () => S.users.map((u) => u.nama).join(' & ');
function avatar(email, cls = '') {
  const u = userOf(email);
  return `<span class="avatar ${cls} ${userIdx(email) === 1 ? 'alt' : ''}" title="${esc(u.nama)}">${esc(String(u.nama || '?').charAt(0).toUpperCase())}</span>`;
}
const whoMini = (email) => `<span class="who-mini">${avatar(email, 'xs')}${esc(userOf(email).nama)}</span>`;
function dpCode(d) {
  if (d.jenis === 'tunai') return 'CASH';
  const w = String(d.nama).trim().split(/\s+/);
  if (w[0].length <= 4) return w[0].toUpperCase();
  return (w[0][0] + (w[1] ? w[1][0] : w[0][1] || '')).toUpperCase();
}
function dpBadge(d, sm) {
  if (!d) return `<span class="wallet-badge ${sm ? 'sm' : ''}">?</span>`;
  if (d.ikon_url) return `<span class="wallet-badge ${sm ? 'sm' : ''}"><img src="${esc(d.ikon_url)}" alt="" loading="lazy" referrerpolicy="no-referrer"></span>`;
  return `<span class="wallet-badge ${sm ? 'sm' : ''}" style="color:${fg(d.warna)};background:${rgba(d.warna, 0.1)};border-color:${rgba(d.warna, 0.28)}">${esc(dpCode(d))}</span>`;
}
const JENIS_DP = { rekening: ['Rekening Bank', 'account_balance'], 'e-wallet': ['E-Wallet', 'account_balance_wallet'], tunai: ['Uang Tunai', 'payments'], lainnya: ['Lainnya', 'wallet'] };
const katPill = (k) => (k ? `<span class="cat-pill" style="${tint(k.warna)}">${esc(k.nama)}</span>` : '<span class="cat-pill" style="background:var(--surface-3);color:var(--text-3)">—</span>');
const katWell = (k, cls = '') => `<span class="icon-well ${cls}" style="${tint(k ? k.warna : '#94A3B8')}">${ic(k ? k.ikon : 'category')}</span>`;

function txTitle(t) {
  if (t.catatan) return t.catatan;
  if (t.tipe === 'transfer') return 'Transfer ke ' + (S.dMap.get(t.dompet_tujuan_id)?.nama || 'dompet lain');
  return S.kMap.get(t.kategori_id)?.nama || 'Transaksi';
}
function txWell(t, cls = '') {
  if (t.tipe === 'transfer') return `<span class="icon-well round ${cls}" style="${tint('#3B82F6')}">${ic('swap_horiz')}</span>`;
  const k = S.kMap.get(t.kategori_id);
  return `<span class="icon-well round ${cls}" style="${tint(k ? k.warna : '#94A3B8')}">${ic(k ? k.ikon : 'category')}</span>`;
}
function txAmt(t) {
  const c = t.tipe === 'masuk' ? 'in' : t.tipe === 'keluar' ? 'out' : 'tr';
  const s = t.tipe === 'masuk' ? '+' : t.tipe === 'keluar' ? '-' : '';
  return `<span class="amt ${c} money">${s}${rp(t.nominal)}</span>`;
}
function txDompet(t) {
  const a = S.dMap.get(t.dompet_id)?.nama || '?';
  if (t.tipe !== 'transfer') return esc(a);
  return `${esc(short(a, 14))} → ${esc(short(S.dMap.get(t.dompet_tujuan_id)?.nama || '?', 14))}`;
}
const txKatPill = (t) => (t.tipe === 'transfer' ? '<span class="cat-pill" style="' + tint('#3B82F6') + '">Transfer Dompet</span>' : katPill(S.kMap.get(t.kategori_id)));

// Agregasi bulanan (dihitung sekali per perubahan data)
function agg() {
  return memo('agg', () => {
    const M = new Map();
    for (const t of S.tx) {
      const mk = t.tanggal.slice(0, 7);
      let a = M.get(mk);
      if (!a) { a = { in: 0, out: 0, n: 0, cat: {}, catN: {}, domOut: {} }; M.set(mk, a); }
      const n = +t.nominal || 0;
      a.n++;
      if (t.tipe === 'masuk') { a.in += n; a.cat[t.kategori_id] = (a.cat[t.kategori_id] || 0) + n; a.catN[t.kategori_id] = (a.catN[t.kategori_id] || 0) + 1; }
      else if (t.tipe === 'keluar') {
        a.out += n; a.cat[t.kategori_id] = (a.cat[t.kategori_id] || 0) + n; a.catN[t.kategori_id] = (a.catN[t.kategori_id] || 0) + 1;
        a.domOut[t.dompet_id] = (a.domOut[t.dompet_id] || 0) + n;
      }
    }
    return M;
  });
}
const EMPTY_AGG = { in: 0, out: 0, n: 0, cat: {}, catN: {}, domOut: {} };
const monthAgg = (mk) => agg().get(mk) || EMPTY_AGG;
function txRange(a, b) { return S.tx.filter((t) => t.tanggal >= a && t.tanggal <= b); }

// Budget
function budgetMonth(mk) {
  return memo('bud' + mk, () => {
    const a = monthAgg(mk);
    return S.budget.filter((b) => b.bulan === mk).map((b) => {
      const k = S.kMap.get(b.kategori_id);
      const used = a.cat[b.kategori_id] || 0;
      const limit = +b.batas_nominal || 0;
      const p = limit ? used / limit : 0;
      return { b, k, used, limit, p, sisa: limit - used, n: a.catN[b.kategori_id] || 0 };
    }).filter((x) => x.k).sort((x, y) => y.limit - x.limit);
  });
}
const levelOf = (p) => (p >= 1 ? 'danger' : p >= 0.8 ? 'warn' : 'ok');

// Tagihan
function billStatus(g) {
  if (g.status === 'lunas') return 'lunas';
  const d = diffDays(today(), g.tanggal_jatuh_tempo);
  if (d < 0) return 'terlambat';
  if (d <= 1) return 'jatuh_tempo';
  return 'akan_datang';
}
const BILL_ST = { lunas: ['Lunas', 'ok'], terlambat: ['Terlambat', 'danger'], jatuh_tempo: ['Jatuh Tempo', 'warn'], akan_datang: ['Akan Datang', 'info'] };
const paidThisMonth = (g) => !!g.terakhir_dibayar && mKey(g.terakhir_dibayar) === mKey(today());
function nextDue(g) {
  if (g.periode === 'bulanan') return addMonths(g.tanggal_jatuh_tempo, 1);
  if (g.periode === 'tahunan') return addMonths(g.tanggal_jatuh_tempo, 12);
  if (g.periode === 'custom') return addDays(g.tanggal_jatuh_tempo, Math.max(1, +g.interval_hari || 30));
  return '';
}
const PERIODE = { sekali: 'Sekali Bayar', bulanan: 'Bulanan', tahunan: 'Tahunan', custom: 'Custom' };
const unpaidBills = () => memo('unpaid', () => S.tagihan.filter((g) => g.status !== 'lunas').sort((a, b) => a.tanggal_jatuh_tempo.localeCompare(b.tanggal_jatuh_tempo)));
const attentionBills = () => unpaidBills().filter((g) => diffDays(today(), g.tanggal_jatuh_tempo) <= 7);
function dueText(g) {
  const st = billStatus(g);
  if (st === 'lunas') return ['Lunas', 'green'];
  const d = diffDays(today(), g.tanggal_jatuh_tempo);
  if (d < 0) return [`${tgl(g.tanggal_jatuh_tempo, 'dm')} — Terlambat ${-d} hari`, 'red'];
  if (d === 0) return [`${tgl(g.tanggal_jatuh_tempo, 'dm')} — Hari ini!`, 'red'];
  if (d === 1) return [`${tgl(g.tanggal_jatuh_tempo, 'dm')} — Besok!`, 'red'];
  return [`${tgl(g.tanggal_jatuh_tempo, 'dm')} (${d} hari lagi)`, d <= 3 ? 'gold' : 'green'];
}

// Target tabungan
const targetPct = (t) => (t.target_nominal ? t.terkumpul / t.target_nominal : 0);
const targetAktif = () => S.target.filter((t) => t.status !== 'dicairkan');

// ════════════════════════════════════════════════════════════
// 3. MESIN SINKRONISASI (Outbox + Optimistic UI)
// ════════════════════════════════════════════════════════════
const SYNC = { outbox: [], busy: false, timer: null, retry: 0, refreshing: false, offline: false, refreshAfter: false, lastOk: Date.now() };
let OB_KEY = 'kd_outbox';

/** Masukkan perubahan ke antrean (disimpan di localStorage agar tidak hilang). */
function enqueue(type, payload) {
  const op = { opId: uid('op'), type, payload: JSON.parse(JSON.stringify(payload)), t: Date.now() };
  SYNC.outbox.push(op);
  LS.set(OB_KEY, SYNC.outbox);
  setSyncPill();
  scheduleFlush(300);
  return op;
}
function scheduleFlush(ms) { clearTimeout(SYNC.timer); SYNC.timer = setTimeout(flushOutbox, ms); }

/** Kirim antrean ke server dalam SATU panggilan sinkronkan() (fetch). */
function flushOutbox() {
  if (SYNC.busy || !SYNC.outbox.length || !AUTH.token) return;
  SYNC.busy = true;
  setSyncPill();
  const batch = SYNC.outbox.slice(0, 40);
  api('sinkronkan', batch).then((res) => {
    SYNC.busy = false;
    SYNC.retry = 0; SYNC.offline = false; SYNC.lastOk = Date.now();
    const sent = new Set(batch.map((o) => o.opId));
    SYNC.outbox = SYNC.outbox.filter((o) => !sent.has(o.opId));
    LS.set(OB_KEY, SYNC.outbox);
    const d = res.data || {};
    const gagal = (d.results || []).filter((r) => !r.ok);
    (d.results || []).forEach((r) => { if (r.ok && r.data) echoServer(r.data); });
    const orangLain = S.versi && d.vSebelum !== S.versi;
    S.versi = d.vSesudah;
    if (gagal.length) {
      toast(gagal.map((g) => g.message).filter(Boolean).slice(0, 2).join(' • ') || 'Sebagian perubahan ditolak server.', 'error', 6000);
      SYNC.refreshAfter = true;
    }
    if (orangLain) SYNC.refreshAfter = true;
    setSyncPill();
    if (SYNC.outbox.length) scheduleFlush(60);
    else if (SYNC.refreshAfter) { SYNC.refreshAfter = false; refreshData({ silent: true, reason: gagal.length ? 'fix' : 'remote' }); }
  }).catch((err) => {
    SYNC.busy = false;
    if (err.code === 'AUTH') { setSyncPill(); return; } // antrean tetap disimpan, dikirim lagi setelah login
    retryLater(err.message);
  });
}
function retryLater(msg) {
  SYNC.offline = true;
  SYNC.retry = Math.min(60000, (SYNC.retry || 1500) * 2);
  setSyncPill();
  scheduleFlush(SYNC.retry);
  if (msg && !/sibuk|network|jaringan|NetworkError/i.test(msg)) console.warn('Sinkronisasi tertunda:', msg);
}
/** Lengkapi data lokal dengan nilai dari server (mis. waktu dibuat_pada resmi). */
function echoServer(data) {
  const fix = (t) => { if (t && t.id && S.txMap.has(t.id)) S.txMap.get(t.id).dibuat_pada = t.dibuat_pada; };
  if (data.tanggal && data.tipe) fix(data);
  if (data.tx) fix(data.tx);
}
function setSyncPill() {
  const el = $('#syncPill');
  if (!el) return;
  const n = SYNC.outbox.length;
  let cls = 'ok', icon = 'cloud_done', txt = 'Tersinkron';
  if (SYNC.offline && n) { cls = 'off'; icon = 'cloud_off'; txt = `Offline · ${n} tertunda`; }
  else if (n || SYNC.busy || SYNC.refreshing) { cls = 'busy'; icon = 'sync'; txt = n ? `Menyimpan ${n}…` : 'Memperbarui…'; }
  el.className = 'sync-pill ' + cls;
  el.innerHTML = `${ic(icon)}<span class="txt">${txt}</span>`;
  el.title = cls === 'ok' ? 'Semua perubahan tersimpan di Google Sheets' : txt;
  el.dataset.act = cls === 'off' ? 'sync-now' : '';
}

/** Ambil ulang seluruh data dari server (dipakai bila pasangan mengubah data). */
function refreshData(opt = {}) {
  if (SYNC.refreshing) return;
  SYNC.refreshing = true;
  setSyncPill();
  api('ambilDataAwal').then((res) => {
    SYNC.refreshing = false;
    setSyncPill();
    if (SYNC.outbox.length) { SYNC.refreshAfter = true; return; } // tunggu antrean kosong
    const sebelum = S.versi;
    simpanCacheBoot(res.data);
    loadBoot(res.data);
    applyTheme(S.settings);
    renderShell();
    rerender(true);
    if (opt.reason === 'remote' && sebelum !== S.versi) toast('Data diperbarui dari anggota keluarga.', 'info');
    else if (!opt.silent) toast('Data dimuat ulang dari server.', 'success');
    if (opt.done) opt.done();
  }).catch((err) => { SYNC.refreshing = false; setSyncPill(); if (!opt.silent && err.code !== 'AUTH') toast(err.message, 'error'); });
}
/** Polling versi ringan: deteksi perubahan dari pasangan. */
function pollVersi() {
  if (document.hidden || SYNC.busy || SYNC.refreshing || SYNC.outbox.length || !AUTH.token) return;
  api('cekVersi').then((res) => { if (res.data.v !== S.versi) refreshData({ silent: true, reason: 'remote' }); }).catch(() => {});
}

// ════════════════════════════════════════════════════════════
// 4. TOAST, MODAL, POPOVER
// ════════════════════════════════════════════════════════════
function toast(msg, type = 'success', ms = 3200) {
  const root = $('#toast-root');
  if (!root) return;
  const icons = { success: 'check', error: 'error', info: 'info', warn: 'warning' };
  const el = document.createElement('div');
  el.className = 'toast ' + type;
  el.innerHTML = `<span class="ti">${ic(icons[type] || 'info')}</span><div>${esc(msg)}</div><button aria-label="Tutup">${ic('close', 'sm')}</button>`;
  const close = () => { el.classList.add('out'); setTimeout(() => el.remove(), 200); };
  el.querySelector('button').onclick = close;
  root.appendChild(el);
  while (root.children.length > 3) root.firstElementChild.remove();
  setTimeout(close, ms);
}

let MODAL = null;
function openModal({ title, icon = '', iconCls = '', body = '', foot = '', size = '', onOpen, onClose }) {
  closeModal(true);
  const root = $('#modal-root');
  root.innerHTML = `<div class="modal-backdrop" data-act="modal-close"></div>
    <div class="modal ${size}" role="dialog" aria-modal="true" aria-labelledby="mdlTitle">
      <div class="modal-h">${icon ? `<span class="modal-ic ${iconCls}">${ic(icon)}</span>` : ''}<h3 id="mdlTitle">${esc(title)}</h3>
        <button class="btn-icon" data-act="modal-close" aria-label="Tutup">${ic('close')}</button></div>
      <div class="modal-b">${body}</div>
      ${foot ? `<div class="modal-f">${foot}</div>` : ''}
    </div>`;
  root.classList.add('open');
  document.body.classList.add('modal-open');
  const el = root.querySelector('.modal');
  MODAL = { el, onClose, prevFocus: document.activeElement };
  if (onOpen) onOpen(el);
  const f = el.querySelector('[autofocus]') || el.querySelector('.modal-b input:not([type=hidden]), .modal-b select, .modal-b textarea');
  if (f && window.matchMedia('(min-width: 768px)').matches) setTimeout(() => f.focus(), 30);
  return el;
}
function closeModal(silent) {
  const root = $('#modal-root');
  if (!root || !root.classList.contains('open')) return;
  const m = MODAL;
  MODAL = null;
  root.classList.remove('open');
  root.innerHTML = '';
  document.body.classList.remove('modal-open');
  if (m && m.onClose && !silent) m.onClose();
  if (m && m.prevFocus && m.prevFocus.focus) try { m.prevFocus.focus({ preventScroll: true }); } catch (e) { /* abaikan */ }
}
/** Dialog konfirmasi → Promise<boolean> */
function confirmBox({ title, text, ok = 'Ya, Lanjutkan', danger = false, icon }) {
  return new Promise((resolve) => {
    let done = false;
    openModal({
      title, size: 'sm', icon: icon || (danger ? 'delete' : 'help'), iconCls: danger ? 'red' : '',
      body: `<p class="muted" style="line-height:1.6">${text}</p>`,
      foot: `<div class="grow"></div><button class="btn btn-ghost" data-act="modal-close">Batal</button><button class="btn ${danger ? 'btn-danger' : 'btn-primary'}" id="cfOk">${esc(ok)}</button>`,
      onOpen: (el) => { el.querySelector('#cfOk').onclick = () => { done = true; closeModal(true); resolve(true); }; },
      onClose: () => { if (!done) resolve(false); }
    });
  });
}

let POP = null;
function openPopover(anchor, html, width) {
  closePopover();
  const el = document.createElement('div');
  el.className = 'popover';
  if (width) el.style.width = width + 'px';
  el.innerHTML = html;
  document.body.appendChild(el);
  const r = anchor.getBoundingClientRect();
  const w = el.offsetWidth;
  el.style.top = Math.min(r.bottom + 8, window.innerHeight - 80) + 'px';
  el.style.left = clamp(r.right - w, 12, window.innerWidth - w - 12) + 'px';
  POP = { el, anchor };
  return el;
}
function closePopover() { if (POP) { POP.el.remove(); POP = null; } }
document.addEventListener('mousedown', (e) => { if (POP && !POP.el.contains(e.target) && !POP.anchor.contains(e.target)) closePopover(); });
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') { if (POP) closePopover(); else if (MODAL) closeModal(); else if (document.body.classList.contains('mode-print')) closePrint(); }
});

/** Tombol dengan status memuat (untuk operasi yang menunggu server). */
async function withBtn(btn, fn) {
  if (!btn) return fn();
  const html = btn.innerHTML;
  btn.classList.add('loading');
  btn.innerHTML = `<span class="spinner"></span>${esc(btn.dataset.loading || 'Memproses…')}`;
  try { return await fn(); } finally { btn.classList.remove('loading'); btn.innerHTML = html; }
}

/** Unduh file dari Blob / base64 (tanpa membuka tab baru). */
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name; a.rel = 'noopener';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}
function b64Blob(b64, type) {
  const bin = atob(b64); const u8 = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  return new Blob([u8], { type });
}

// ════════════════════════════════════════════════════════════
// 5. ROUTER SPA (tampil/sembunyi — tanpa reload, tanpa ubah URL)
// ════════════════════════════════════════════════════════════
const PAGES = {};
let CUR = 'dashboard';
const NAV_OF = { 'tagihan-form': 'tagihan' };
function goPage(page, opt = {}) {
  if (!PAGES[page]) page = 'dashboard';
  if (CUR === 'pengaturan' && page !== 'pengaturan' && typeof revertLiveTheme === 'function') revertLiveTheme();
  closePopover();
  const prev = CUR;
  CUR = page;
  if (prev !== page) {
    // kosongkan halaman yang ditinggalkan: DOM ringan & tidak ada ID ganda
    const old = document.getElementById('page-' + prev);
    if (old) old.innerHTML = '';
    if (typeof CHARTS !== 'undefined') Object.keys(CHARTS).forEach((id) => { if (!CHARTS[id].canvas || !CHARTS[id].canvas.isConnected) { CHARTS[id].destroy(); delete CHARTS[id]; } });
  }
  $$('.page').forEach((p) => p.classList.toggle('active', p.id === 'page-' + page));
  const nav = NAV_OF[page] || page;
  $$('[data-go]').forEach((b) => b.classList.toggle('active', b.dataset.go === nav && !b.dataset.f));
  closeDrawer();
  if (PAGES[page].prep) PAGES[page].prep(opt);
  renderPage(page);
  document.title = `${PAGES[page].t} · ${S.settings.nama_aplikasi || 'Keuangan Digital'}`;
  window.scrollTo({ top: 0, behavior: 'instant' in document.documentElement.style ? 'instant' : 'auto' });
}
function renderPage(p) {
  try { PAGES[p].r(); } catch (err) {
    console.error(err);
    $('#page-' + p).innerHTML = `<div class="card empty"><span class="icon-well">${ic('error')}</span><b>Halaman gagal ditampilkan</b>${esc(err.message)}</div>`;
  }
}
let _rr = 0;
/** Render ulang halaman aktif (digabung per frame agar tidak berulang). */
function rerender(now) {
  const run = () => { renderShell(); if (!PAGES[CUR] || !PAGES[CUR].noAuto) renderPage(CUR); };
  if (now) { run(); return; }
  cancelAnimationFrame(_rr);
  _rr = requestAnimationFrame(run);
}
function openDrawer() { $('#sidebar').classList.add('open'); $('#scrim').classList.add('show'); }
function closeDrawer() { $('#sidebar')?.classList.remove('open'); $('#scrim')?.classList.remove('show'); }

// Delegasi event global
const ACT = {};   // klik: data-act
const ON = {};    // input: data-in
const CHG = {};   // change: data-ch
document.addEventListener('click', (e) => {
  const go = e.target.closest('[data-go]');
  if (go) { e.preventDefault(); goPage(go.dataset.go, { ...go.dataset }); return; }
  const a = e.target.closest('[data-act]');
  if (!a) return;
  const fn = ACT[a.dataset.act];
  if (fn) { e.preventDefault(); fn(a, e); }
});
document.addEventListener('input', (e) => { const el = e.target.closest('[data-in]'); if (el && ON[el.dataset.in]) ON[el.dataset.in](el, e); });
document.addEventListener('change', (e) => { const el = e.target.closest('[data-ch]'); if (el && CHG[el.dataset.ch]) CHG[el.dataset.ch](el, e); });
document.addEventListener('keydown', (e) => {
  const el = e.target.closest('[data-enter]');
  if (el && e.key === 'Enter' && ACT[el.dataset.enter]) { e.preventDefault(); ACT[el.dataset.enter](el, e); }
});

ACT['modal-close'] = () => closeModal();
ACT['drawer-open'] = openDrawer;
ACT['drawer-close'] = closeDrawer;

/** Input nominal Rupiah: format ribuan otomatis + terbilang. */
ON.money = (el) => {
  const n = parseNum(el.value);
  const s = n ? fmtN(n) : '';
  if (el.value !== s) { el.value = s; try { el.setSelectionRange(s.length, s.length); } catch (e) { /* abaikan */ } }
  if (el.dataset.words) { const w = $(el.dataset.words); if (w) w.textContent = n ? terbilang(n) : ''; }
  if (el.dataset.after && ON[el.dataset.after]) ON[el.dataset.after](el);
};
const moneyVal = (sel) => parseNum(($(sel) || {}).value);

// Komponen kecil reusable
const progress = (p, cls = '', lvl) => `<div class="progress ${cls}"><i class="${lvl || levelOf(p)}" style="width:${clamp(p * 100, 0, 100).toFixed(1)}%"></i></div>`;
function ring(p, size = 64, stroke = 7, color) {
  const r = (size - stroke) / 2, c = 2 * Math.PI * r, v = clamp(p, 0, 1);
  const col = color || (p >= 1 ? 'var(--danger)' : p >= 0.8 ? 'var(--gold)' : 'var(--primary)');
  return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" aria-hidden="true"><circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="var(--primary-50)" stroke-width="${stroke}"/><circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="${col}" stroke-width="${stroke}" stroke-linecap="round" stroke-dasharray="${(c * v).toFixed(2)} ${c.toFixed(2)}"/></svg>`;
}
const emptyBox = (icon, title, text = '', action = '') => `<div class="empty"><span class="icon-well">${ic(icon)}</span><b>${esc(title)}</b>${text ? `<div class="small">${text}</div>` : ''}${action ? `<div class="mt-2">${action}</div>` : ''}</div>`;
function pager(total, page, per, act) {
  const pages = Math.max(1, Math.ceil(total / per));
  if (total <= per) return `<div class="pager"><span>Menampilkan <b>${total}</b> data</span></div>`;
  const from = (page - 1) * per + 1, to = Math.min(total, page * per);
  const nums = [];
  for (let i = 1; i <= pages; i++) if (i === 1 || i === pages || Math.abs(i - page) <= 1) nums.push(i); else if (nums[nums.length - 1] !== '…') nums.push('…');
  return `<div class="pager"><span>Menampilkan <b>${from}–${to}</b> dari <b>${fmtN(total)}</b> data</span><div class="pages">
    <button data-act="${act}" data-p="${page - 1}" ${page <= 1 ? 'disabled' : ''} aria-label="Sebelumnya">${ic('chevron_left', 'sm')}</button>
    ${nums.map((n) => (n === '…' ? '<button disabled>…</button>' : `<button data-act="${act}" data-p="${n}" class="${n === page ? 'on' : ''}">${n}</button>`)).join('')}
    <button data-act="${act}" data-p="${page + 1}" ${page >= pages ? 'disabled' : ''} aria-label="Berikutnya">${ic('chevron_right', 'sm')}</button></div></div>`;
}
const optList = (items, sel, empty) => (empty ? `<option value="">${esc(empty)}</option>` : '') + items.map((x) => `<option value="${esc(x.id)}" ${x.id === sel ? 'selected' : ''}>${esc(x.nama)}${x.extra ? esc(x.extra) : ''}</option>`).join('');
const dompetOpts = (sel, withSaldo = true, empty, list) => optList((list || dompetAktif()).map((d) => ({ id: d.id, nama: d.nama, extra: withSaldo ? ` (Saldo: ${rp(d.saldo)})` : '' })), sel, empty);


// ════════════════════════════════════════════════════════════
// 6. AKSI DATA — Optimistic UI (ubah state → render → antrean)
// ════════════════════════════════════════════════════════════
const DP_COLS = ['id', 'nama', 'jenis', 'saldo_awal', 'ikon_url', 'warna', 'aktif', 'keterangan', 'urutan'];
const pick = (o, cols) => cols.reduce((r, c) => ((r[c] = o[c] ?? ''), r), {});
const txPayload = (t) => pick(t, TX_COLS);

const DB = {
  /** Efek transaksi ke saldo dompet (sign +1 tambah, -1 batalkan). */
  applyTx(t, sign) {
    const n = (+t.nominal || 0) * sign;
    const a = S.dMap.get(t.dompet_id), b = S.dMap.get(t.dompet_tujuan_id);
    if (t.tipe === 'masuk') { if (a) a.saldo += n; }
    else if (t.tipe === 'keluar') { if (a) a.saldo -= n; }
    else { if (a) a.saldo -= n; if (b) b.saldo += n; }
    const now = nowStamp();
    if (a) a.terakhir = now;
    if (b && t.tipe === 'transfer') b.terakhir = now;
  },
  simpanTx(t, kirim = true) {
    t = { ...t, nominal: +t.nominal || 0 };
    const lama = S.txMap.get(t.id);
    if (lama) { DB.applyTx(lama, -1); Object.assign(lama, t); DB.applyTx(lama, 1); }
    else {
      t.dibuat_pada = t.dibuat_pada || nowStamp(); t.ref = t.ref || ''; t.dompet_tujuan_id = t.dompet_tujuan_id || ''; t.kategori_id = t.kategori_id || '';
      S.tx.push(t); S.txMap.set(t.id, t); DB.applyTx(t, 1);
    }
    sortTx(); bump();
    if (kirim) enqueue('tx.simpan', txPayload(S.txMap.get(t.id)));
    rerender();
  },
  hapusTx(id) {
    const t = S.txMap.get(id);
    if (!t) return;
    DB.applyTx(t, -1);
    S.tx = S.tx.filter((x) => x.id !== id); S.txMap.delete(id);
    bump(); enqueue('tx.hapus', { id }); rerender();
  },
  simpanDompet(d) {
    const lama = S.dMap.get(d.id);
    if (lama) { const delta = (+d.saldo_awal || 0) - (+lama.saldo_awal || 0); Object.assign(lama, d); lama.saldo += delta; }
    else S.dompet.push({ ...d, saldo: +d.saldo_awal || 0, terakhir: nowStamp(), dibuat_pada: nowStamp() });
    reindex(); bump(); enqueue('dompet.simpan', pick(S.dMap.get(d.id), DP_COLS)); rerender();
  },
  hapusDompet(id) { S.dompet = S.dompet.filter((d) => d.id !== id); reindex(); bump(); enqueue('dompet.hapus', { id }); rerender(); },
  urutanDompet(ids) { ids.forEach((id, i) => { const d = S.dMap.get(id); if (d) d.urutan = i + 1; }); bump(); enqueue('dompet.urutan', { ids }); rerender(); },
  simpanKategori(k) {
    const lama = S.kMap.get(k.id);
    if (lama) Object.assign(lama, k); else S.kategori.push(k);
    reindex(); bump(); enqueue('kategori.simpan', k); rerender();
  },
  hapusKategori(id) { S.kategori = S.kategori.filter((k) => k.id !== id); reindex(); bump(); enqueue('kategori.hapus', { id }); rerender(); },
  simpanBudget(list, hapusIds = []) {
    list.forEach((b) => {
      const i = S.budget.findIndex((x) => x.id === b.id || (x.bulan === b.bulan && x.kategori_id === b.kategori_id));
      if (i >= 0) S.budget[i] = { ...S.budget[i], ...b }; else S.budget.push({ ...b });
      enqueue('budget.simpan', b);
    });
    hapusIds.forEach((id) => { S.budget = S.budget.filter((x) => x.id !== id); enqueue('budget.hapus', { id }); });
    bump(); rerender();
  },
  simpanTarget(t) {
    const i = S.target.findIndex((x) => x.id === t.id);
    if (i >= 0) S.target[i] = { ...S.target[i], ...t }; else S.target.push({ ...t, dibuat_pada: nowStamp() });
    bump(); enqueue('target.simpan', t); rerender();
  },
  hapusTarget(id) { S.target = S.target.filter((t) => t.id !== id); bump(); enqueue('target.hapus', { id }); rerender(); },
  setorTarget(id, nominal, tx) {
    const t = S.target.find((x) => x.id === id);
    if (!t) return;
    if (tx) DB.simpanTx({ ...tx, tipe: 'transfer', ref: 'target:' + id }, false);
    const baru = Math.max(0, t.terkumpul + nominal);
    let st = t.status || 'aktif';
    if (st === 'aktif' && baru >= t.target_nominal) st = 'tercapai';
    if (st === 'tercapai' && baru < t.target_nominal) st = 'aktif';
    t.terkumpul = baru; t.status = st;
    bump(); enqueue('target.setor', { id, nominal, tx: tx ? txPayload(S.txMap.get(tx.id)) : null }); rerender();
  },
  simpanTagihan(g) {
    const i = S.tagihan.findIndex((x) => x.id === g.id);
    if (i >= 0) S.tagihan[i] = { ...S.tagihan[i], ...g }; else S.tagihan.push({ ...g, dibuat_pada: nowStamp() });
    bump(); enqueue('tagihan.simpan', g); rerender();
  },
  hapusTagihan(id) { S.tagihan = S.tagihan.filter((g) => g.id !== id); bump(); enqueue('tagihan.hapus', { id }); rerender(); },
  bayarTagihan(g, { tanggal, dompet_id, nominal, buatTx, lanjut, oleh }) {
    let tx = null;
    if (buatTx) {
      const kat = g.kategori_id && S.kMap.has(g.kategori_id) ? g.kategori_id : (S.kMap.has('kat_tagihan') ? 'kat_tagihan' : katByTipe('pengeluaran')[0]?.id);
      tx = { id: uid('tx'), tanggal, tipe: 'keluar', dompet_id, dompet_tujuan_id: '', kategori_id: kat, nominal, catatan: 'Bayar ' + g.nama, dibuat_oleh: oleh || S.me.email, ref: 'tagihan:' + g.id };
      DB.simpanTx(tx, false);
    }
    const jt = g.periode !== 'sekali' && lanjut ? nextDue(g) : '';
    g.terakhir_dibayar = tanggal;
    if (jt) { g.tanggal_jatuh_tempo = jt; g.status = 'aktif'; } else g.status = 'lunas';
    bump(); enqueue('tagihan.bayar', { id: g.id, tanggal, jatuh_tempo_baru: jt, tx: tx ? txPayload(S.txMap.get(tx.id)) : null }); rerender();
    return jt;
  },
  simpanPengaturan(obj) {
    S.settings = { ...S.settings, ...obj };
    bump(); enqueue('pengaturan.simpan', obj); applyTheme(S.settings); rerender();
  },
  simpanPengguna(list) {
    S.users = list.map((u) => ({ ...u, email: u.email.toLowerCase(), peran: 'Pemilik' }));
    reindex(); bump(); enqueue('pengguna.simpan', { list: S.users }); rerender();
  }
};

// ════════════════════════════════════════════════════════════
// 7. KERANGKA (sidebar, topbar, notifikasi, akun, tema)
// ════════════════════════════════════════════════════════════
function sapaan() {
  const h = jamJkt();
  if (h >= 4 && h < 11) return 'Selamat Pagi';
  if (h >= 11 && h < 15) return 'Selamat Siang';
  if (h >= 15 && h < 18) return 'Selamat Sore';
  return 'Selamat Malam';
}
/** Status kesehatan keuangan bulan berjalan. */
function health() {
  return memo('health', () => {
    const a = monthAgg(mKey(today()));
    const over = budgetMonth(mKey(today())).filter((x) => x.p >= 1).length;
    const telat = unpaidBills().filter((g) => billStatus(g) === 'terlambat').length;
    const rate = a.in ? (a.in - a.out) / a.in : (a.out ? -1 : 0);
    if (a.out > a.in && a.in > 0) return { label: 'Perlu Waspada', cls: 'warn', rate };
    if (telat || over > 1) return { label: 'Perlu Perhatian', cls: 'warn', rate };
    return { label: 'Keuangan Sehat', cls: 'ok', rate };
  });
}
function notifItems() {
  return memo('notif', () => {
    const items = [];
    unpaidBills().forEach((g) => {
      const d = diffDays(today(), g.tanggal_jatuh_tempo);
      if (d > 7) return;
      items.push({ icon: d < 0 ? 'error' : 'event_upcoming', color: d <= 1 ? '#EB5757' : '#F5B13D', title: g.nama, sub: `${rp(g.nominal)} • ${d < 0 ? 'terlambat ' + -d + ' hari' : relHari(g.tanggal_jatuh_tempo).toLowerCase()}`, act: 'bill-pay', id: g.id });
    });
    budgetMonth(mKey(today())).filter((x) => x.p >= 0.8).forEach((x) => {
      items.push({ icon: x.p >= 1 ? 'warning' : 'donut_large', color: x.p >= 1 ? '#EB5757' : '#F5B13D', title: `Budget ${x.k.nama} ${pct(x.p)}`, sub: x.p >= 1 ? `Melebihi pagu ${rp(-x.sisa)}` : `Sisa ${rp(x.sisa)}`, go: 'budget' });
    });
    return items;
  });
}
function renderShell() {
  const set = S.settings;
  const nama = set.nama_aplikasi || 'Keuangan Digital';
  $$('[data-slot="app-name"]').forEach((e) => (e.textContent = nama));
  $$('[data-slot="app-sub"]').forEach((e) => (e.textContent = set.subjudul || 'Keluarga ' + familyName()));
  const logo = set.logo_url ? `<img src="${esc(set.logo_url)}" alt="Logo ${esc(nama)}" referrerpolicy="no-referrer">` : ic('eco');
  $$('[data-slot="logo"]').forEach((e) => { if (e.dataset.logo !== logo) { e.innerHTML = logo; e.dataset.logo = logo; } });
  const td = $('[data-slot="today"] span:last-child');
  if (td) td.textContent = tgl(today(), 'day');
  $('[data-slot="avatars"]').innerHTML = S.users.map((u) => avatar(u.email)).join('');
  $('[data-slot="family"]').textContent = familyName();
  $('[data-slot="me"]').textContent = 'Masuk sebagai ' + S.me.nama;
  const att = attentionBills().length;
  const bt = $('[data-slot="badge-tagihan"]');
  bt.textContent = att; bt.hidden = !att;
  const navT = $('.nav-item[data-go="tagihan"]');
  if (att) navT.dataset.badge = att; else delete navT.dataset.badge;
  const tAktif = S.target.filter((t) => t.status === 'aktif').length;
  const bTab = $('[data-slot="badge-tabungan"]');
  bTab.hidden = true; bTab.textContent = tAktif;
  const nc = notifItems().length;
  const nb = $('[data-slot="notif-count"]');
  nb.textContent = nc; nb.hidden = !nc;
  $('[data-slot="bn-dot"]').hidden = !att;
  const bud = budgetMonth(mKey(today()));
  const lim = sum(bud, (x) => x.limit), used = sum(bud, (x) => x.used);
  const tip = $('[data-slot="side-tip"]');
  if (lim) {
    const p = used / lim;
    tip.innerHTML = `<b>${ic(p >= 1 ? 'warning' : 'verified_user', 'sm')}${p >= 1 ? 'Pagu Terlampaui' : p >= 0.8 ? 'Mendekati Batas' : 'Keluarga Hemat'}</b>Pagu bulan ini ${pct(p)} terpakai. ${p >= 1 ? 'Tahan pengeluaran non-prioritas.' : 'Pertahankan hingga akhir bulan!'}`;
    tip.hidden = false;
  } else tip.hidden = true;
  const tb = $('#btnTheme .ms');
  if (tb) tb.textContent = isDark() ? 'light_mode' : 'dark_mode';
  setSyncPill();
}

ACT.theme = () => {
  const next = isDark() ? 'light' : 'dark';
  document.documentElement.dataset.theme = next;
  try { localStorage.setItem('kd_theme', next); } catch (e) { /* abaikan */ }
  applyTheme(S.settings);
  rerender(true);
};
ACT.privasi = () => {
  const on = !document.documentElement.classList.contains('privasi');
  document.documentElement.classList.toggle('privasi', on);
  try { localStorage.setItem('kd_privasi', on ? '1' : '0'); } catch (e) { /* abaikan */ }
  rerender(true);
};
ACT.notif = (el) => {
  const items = notifItems();
  openPopover(el, `<div class="pop-h">Notifikasi <span class="pill ${items.length ? 'pill-danger' : 'pill-ok'}">${items.length} perlu perhatian</span></div>
    ${items.length ? items.map((n) => `<button class="pop-i" ${n.act ? `data-act="${n.act}" data-id="${n.id}"` : `data-go="${n.go}"`}><span class="icon-well sm" style="${tint(n.color)}">${ic(n.icon)}</span><span class="grow"><b>${esc(n.title)}</b><small>${esc(n.sub)}</small></span>${ic('chevron_right', 'sm faint')}</button>`).join('')
      : `<div class="empty" style="padding:20px">${ic('notifications_off', 'lg faint')}<div class="small mt-1">Tidak ada tagihan atau budget yang perlu diperhatikan.</div></div>`}
    <hr><button class="pop-i" data-go="tagihan">${ic('receipt', 'faint')}<span class="grow"><b>Lihat semua tagihan</b></span></button>`);
};
ACT.usermenu = (el) => {
  openPopover(el, `<div class="pop-h"><span>Akun Keluarga</span></div>
    ${S.users.map((u) => `<div class="pop-i">${avatar(u.email)}<span class="grow"><b>${esc(u.nama)} ${u.email === S.me.email ? '<span class="pill pill-ok" style="height:18px">Anda</span>' : ''}</b><small>${esc(u.email)}</small></span></div>`).join('')}
    <hr>
    <button class="pop-i" data-act="theme">${ic(isDark() ? 'light_mode' : 'dark_mode', 'faint')}<span class="grow"><b>${isDark() ? 'Mode Terang' : 'Mode Gelap'}</b></span></button>
    <button class="pop-i" data-act="privasi">${ic(document.documentElement.classList.contains('privasi') ? 'visibility' : 'visibility_off', 'faint')}<span class="grow"><b>${document.documentElement.classList.contains('privasi') ? 'Tampilkan Nominal' : 'Sembunyikan Nominal'}</b><small>Mode privasi saat di tempat umum</small></span></button>
    <button class="pop-i" data-act="reload">${ic('sync', 'faint')}<span class="grow"><b>Muat Ulang Data</b><small>Ambil data terbaru dari Google Sheets</small></span></button>
    <button class="pop-i" data-act="profil">${ic('group', 'faint')}<span class="grow"><b>Profil Keluarga</b></span></button>
    <hr><button class="pop-i" data-act="logout">${ic('logout', 'neg')}<span class="grow"><b class="neg">Keluar</b></span></button>`, 300);
};
ACT.reload = () => { closePopover(); if (SYNC.outbox.length) { toast('Menunggu perubahan tersimpan dulu…', 'info'); SYNC.refreshAfter = true; flushOutbox(); return; } refreshData(); };
ACT.profil = () => { closePopover(); goPage('pengaturan', { tab: 'pengguna' }); };
ACT.logout = () => {
  closePopover(); closeDrawer();
  openModal({
    title: 'Keluar dari Aplikasi', icon: 'logout', size: 'sm',
    body: `<p class="muted" style="line-height:1.6">Anda masuk sebagai <b>${esc(S.me.nama)}</b> (${esc(S.me.email)}).</p>
      ${SYNC.outbox.length ? `<div class="note mt-2">${ic('warning')}<div>Masih ada <b>${SYNC.outbox.length}</b> perubahan yang belum terkirim. Perubahan tetap tersimpan di perangkat ini dan dikirim saat Anda masuk lagi.</div></div>` : ''}
      <div class="stack mt-3">
        <button class="link-row" data-act="do-logout" data-switch="1"><span class="icon-well sm">${ic('switch_account')}</span><span class="grow"><b>Ganti Akun Google</b><small>Masuk dengan akun pasangan Anda</small></span></button>
        <button class="link-row" data-act="do-logout"><span class="icon-well sm" style="${tint('#EB5757')}">${ic('logout')}</span><span class="grow"><b>Keluar dari Aplikasi</b><small>Sesi di perangkat ini diakhiri</small></span></button>
      </div>`,
    foot: '<div class="grow"></div><button class="btn btn-ghost" data-act="modal-close">Batal</button>'
  });
};

// Pencarian global → halaman Transaksi
const _globalSearch = debounce((v) => { TXF.q = v; TXF.page = 1; if (CUR !== 'transaksi') goPage('transaksi', { keep: 1 }); else renderTxResults(); }, 250);
ON.gsearch = (el) => _globalSearch(el.value);

// ════════════════════════════════════════════════════════════
// 8. GRAFIK (Chart.js dimuat malas)
// ════════════════════════════════════════════════════════════
const CHARTS = {};
function whenChart(cb) {
  if (window.Chart) { cb(); return; }
  loadScript(CHART_URL).then(() => {
    Chart.defaults.font.family = "Inter, system-ui, sans-serif";
    Chart.defaults.font.size = 11;
    Chart.defaults.maintainAspectRatio = false;
    Chart.defaults.plugins.legend.display = false;
    Chart.defaults.animation.duration = 420;
    cb();
  }).catch(() => { /* grafik opsional: abaikan bila CDN gagal */ });
}
function mkChart(id, cfg) {
  const cv = document.getElementById(id);
  if (!cv) return;
  if (CHARTS[id]) { CHARTS[id].destroy(); delete CHARTS[id]; }
  CHARTS[id] = new Chart(cv, cfg);
}
function CC() {
  return { text: cssVar('--text-3'), grid: cssVar('--divider'), primary: cssVar('--primary'), accent: cssVar('--accent'), surface: cssVar('--surface'), p50: cssVar('--primary-50'), danger: cssVar('--danger'), gold: cssVar('--gold'), border: cssVar('--border') };
}
const TIP = () => ({
  backgroundColor: isDark() ? '#E7F1EB' : '#13231B', titleColor: isDark() ? '#13231B' : '#fff', bodyColor: isDark() ? '#13231B' : '#fff',
  padding: 10, cornerRadius: 10, displayColors: true, boxPadding: 4, titleFont: { weight: '600' }
});
const plugDashed = {
  id: 'kdDashed',
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((ds, di) => {
      if (!ds.kdDashed) return;
      const meta = chart.getDatasetMeta(di);
      ctx.save(); ctx.setLineDash([4, 3]); ctx.lineWidth = 1.5; ctx.strokeStyle = ds.kdDashed;
      meta.data.forEach((bar, i) => {
        if (!ds.data[i]) return;
        const p = bar.getProps(['x', 'y', 'base', 'width'], true);
        const x = p.x - p.width / 2, w = p.width, y = p.y, h = p.base - p.y, r = Math.min(6, w / 2);
        ctx.beginPath();
        ctx.moveTo(x, y + h); ctx.lineTo(x, y + r); ctx.quadraticCurveTo(x, y, x + r, y); ctx.lineTo(x + w - r, y); ctx.quadraticCurveTo(x + w, y, x + w, y + r); ctx.lineTo(x + w, y + h);
        ctx.stroke();
      });
      ctx.restore();
    });
  }
};
const plugAvg = {
  id: 'kdAvg',
  afterDatasetsDraw(chart, args, o) {
    if (!o || !o.value) return;
    const { ctx, chartArea: a, scales } = chart;
    const y = scales.y.getPixelForValue(o.value);
    if (y < a.top || y > a.bottom) return;
    ctx.save(); ctx.setLineDash([3, 4]); ctx.strokeStyle = o.color; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(a.left, y); ctx.lineTo(a.right, y); ctx.stroke();
    ctx.setLineDash([]); ctx.font = '600 10px Inter, sans-serif'; ctx.fillStyle = o.color; ctx.textAlign = 'right';
    ctx.fillText(o.label, a.right, y - 5); ctx.restore();
  }
};
const plugTag = {
  id: 'kdTag',
  afterDatasetsDraw(chart, args, o) {
    if (!o || o.index == null) return;
    const bar = chart.getDatasetMeta(0).data[o.index];
    if (!bar) return;
    const { ctx } = chart;
    const txt = o.text;
    ctx.save(); ctx.font = '700 10.5px "Plus Jakarta Sans", Inter, sans-serif';
    const w = ctx.measureText(txt).width + 14, h = 20, x = clamp(bar.x - w / 2, chart.chartArea.left, chart.chartArea.right - w), y = bar.y - h - 8;
    ctx.fillStyle = o.bg; ctx.beginPath();
    if (ctx.roundRect) ctx.roundRect(x, y, w, h, 6); else ctx.rect(x, y, w, h);
    ctx.fill(); ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(txt, x + w / 2, y + h / 2 + 0.5);
    ctx.restore();
  }
};
function drawDonut(id, items) {
  whenChart(() => {
    const c = CC();
    const has = items.some((i) => i.value > 0);
    mkChart(id, {
      type: 'doughnut',
      data: { labels: has ? items.map((i) => i.label) : ['Belum ada data'], datasets: [{ data: has ? items.map((i) => i.value) : [1], backgroundColor: has ? items.map((i) => i.color) : [c.p50], borderColor: c.surface, borderWidth: 3, hoverOffset: 5, borderRadius: 3 }] },
      options: { cutout: '74%', plugins: { tooltip: has ? { ...TIP(), callbacks: { label: (x) => ` ${x.label}: ${rp(x.parsed)}` } } : { enabled: false } } }
    });
  });
}
const yAxis = (c) => ({ beginAtZero: true, grid: { color: c.grid, drawTicks: false }, border: { display: false }, ticks: { color: c.text, padding: 8, maxTicksLimit: 5, callback: (v) => rpS(v).replace('Rp ', '') } });
const xAxis = (c) => ({ grid: { display: false }, border: { display: false }, ticks: { color: c.text, padding: 6 } });

// ════════════════════════════════════════════════════════════
// 9. DASHBOARD
// ════════════════════════════════════════════════════════════
const DASH = { mk: '' };
PAGES.dashboard = { t: 'Dashboard', r: renderDashboard };
function donutItems(cat, max = 4) {
  const arr = Object.entries(cat).map(([id, v]) => ({ k: S.kMap.get(id), v })).filter((x) => x.v > 0).sort((a, b) => b.v - a.v);
  const top = arr.slice(0, max).map((x) => ({ label: x.k ? x.k.nama : 'Tanpa kategori', value: x.v, color: x.k ? x.k.warna : '#94A3B8' }));
  const rest = sum(arr.slice(max), (x) => x.v);
  if (rest > 0) top.push({ label: 'Lainnya', value: rest, color: '#BAC6BF' });
  return top;
}
function legendHtml(items) {
  const tot = sum(items, (i) => i.value) || 1;
  return items.length ? items.map((i) => `<div class="lg"><i style="background:${i.color}"></i><span title="${esc(i.label)}">${esc(i.label)}<em>${pct(i.value / tot)}</em></span><b class="money">${rp(i.value)}</b></div>`).join('') : '<div class="small faint">Belum ada pengeluaran pada periode ini.</div>';
}
/** Perubahan total saldo dibanding tanggal yang sama bulan lalu (adil terhadap waktu gajian). */
function saldoChange(total) {
  const T = today(), d0 = addMonths(T, -1);
  const net = sum(S.tx.filter((t) => t.tanggal > d0 && t.tanggal <= T && t.tipe !== 'transfer'), (t) => (t.tipe === 'masuk' ? t.nominal : -t.nominal));
  const prev = total - net;
  return prev ? (total - prev) / Math.abs(prev) : 0;
}
function buildInsights() {
  const T = today(), mk = mKey(T), prev = mAdd(mk, -1), day = +T.slice(8);
  const a = monthAgg(mk), b = monthAgg(prev);
  const out = [];
  const prevMTD = sum(S.tx.filter((t) => t.tipe === 'keluar' && mKey(t.tanggal) === prev && +t.tanggal.slice(8) <= day), (t) => t.nominal);
  if (prevMTD > 0 && a.out > 0) {
    const ch = (a.out - prevMTD) / prevMTD;
    out.push({ icon: ch <= 0 ? 'trending_down' : 'trending_up', c: ch <= 0 ? '#2DB37A' : '#EB5757', html: `Pengeluaran bulan ini <b class="money">${rp(a.out)}</b>, ${ch <= 0 ? 'lebih hemat' : 'naik'} <b>${pct(Math.abs(ch))}</b> dibanding periode yang sama bulan lalu.` });
  }
  const topCat = Object.entries(a.cat).filter(([id]) => S.kMap.get(id)?.tipe === 'pengeluaran').sort((x, y) => y[1] - x[1])[0];
  if (topCat && a.out) out.push({ icon: 'pie_chart', c: S.kMap.get(topCat[0]).warna, html: `<b>${esc(S.kMap.get(topCat[0]).nama)}</b> menyerap <b>${pct(topCat[1] / a.out)}</b> dari total pengeluaran bulan ini.` });
  const bud = budgetMonth(mk);
  const over = bud.filter((x) => x.p >= 1), near = bud.filter((x) => x.p >= 0.8 && x.p < 1);
  if (over.length) out.push({ icon: 'warning', c: '#EB5757', html: `Budget <b>${esc(over[0].k.nama)}</b> melebihi pagu <b class="money">${rp(-over[0].sisa)}</b>${over.length > 1 ? ` (+${over.length - 1} kategori lain)` : ''}. Pertimbangkan menahan pos ini hingga akhir bulan.` });
  else if (near.length) out.push({ icon: 'donut_large', c: '#F5B13D', html: `Budget <b>${esc(near[0].k.nama)}</b> sudah <b>${pct(near[0].p)}</b> terpakai, sisa <b class="money">${rp(near[0].sisa)}</b> untuk ${mDays(mk) - day} hari.` });
  const due = attentionBills();
  if (due.length) {
    const tot = sum(due, (g) => g.nominal);
    out.push({ icon: 'event_upcoming', c: '#3B82F6', html: `<b>${due.length} tagihan</b> jatuh tempo dalam 7 hari (total <b class="money">${rp(tot)}</b>). ${tot <= totalSaldo() ? 'Saldo keluarga mencukupi.' : 'Saldo belum mencukupi — siapkan dana.'}` });
  }
  if (a.in > 0) {
    const r = (a.in - a.out) / a.in;
    out.push({ icon: 'savings', c: r >= 0.2 ? '#2DB37A' : '#F5B13D', html: `Tingkat tabungan bulan ini <b>${pct(r)}</b> dari pemasukan — ${r >= 0.3 ? 'sangat sehat' : r >= 0.1 ? 'cukup baik' : r >= 0 ? 'masih tipis' : 'defisit'}.` });
  }
  const lim = sum(bud, (x) => x.limit);
  if (day >= 5 && a.out) {
    const proj = (a.out / day) * mDays(mk);
    out.push({ icon: 'query_stats', c: '#0F5C3F', html: `Rata-rata belanja <b class="money">${rpS(a.out / day)}/hari</b>; proyeksi akhir bulan <b class="money">${rpS(proj)}</b>${lim ? ` (${pct(proj / lim)} dari pagu)` : ''}.` });
  }
  if (!out.length) out.push({ icon: 'lightbulb', c: '#0F5C3F', html: 'Mulai catat transaksi harian — insight otomatis akan muncul di sini.' });
  return out.slice(0, 4);
}
function renderDashboard() {
  const T = today(), mk = mKey(T);
  if (!DASH.mk) DASH.mk = mk;
  const a = monthAgg(mk);
  const total = totalSaldo();
  const ch = saldoChange(total);
  const h = health();
  const unpaid = unpaidBills();
  const att = attentionBills();
  const urgent = att.some((g) => diffDays(T, g.tanggal_jatuh_tempo) <= 1);
  const dompet = dompetAktif();
  const bud = budgetMonth(mk);
  const lim = sum(bud, (x) => x.limit), used = sum(bud, (x) => x.used), bp = lim ? used / lim : 0;
  const recent = S.tx.slice(0, 6);
  const targets = targetAktif().slice(0, 4);
  const dItems = donutItems(monthAgg(DASH.mk).cat ? Object.fromEntries(Object.entries(monthAgg(DASH.mk).cat).filter(([id]) => S.kMap.get(id)?.tipe === 'pengeluaran')) : {});
  const dTotal = monthAgg(DASH.mk).out;
  const months = Array.from({ length: 6 }, (_, i) => mAdd(mk, -i));
  const privasi = document.documentElement.classList.contains('privasi');

  $('#page-dashboard').innerHTML = `<div class="dash">
  <div class="hero a-hero">
    <div class="hero-main">
      <div class="hero-greet"><h1>${sapaan()}, ${esc(S.me.nama)}</h1>${ic('waving_hand', 'wave')}<span class="pill pill-glass"><span class="dot"></span>${h.label}</span></div>
      <p class="hero-sub">Ringkasan arus kas keuangan keluarga hari ini</p>
      <div class="hero-label">TOTAL SALDO BERSAMA <button data-act="privasi" aria-label="${privasi ? 'Tampilkan' : 'Sembunyikan'} nominal">${ic(privasi ? 'visibility_off' : 'visibility')}</button></div>
      <div class="hero-amount-row"><div class="hero-amount money">${rp(total)}</div>
        <span class="pill pill-light pill-lg ${ch < 0 ? 'down' : ''}">${ic(ch < 0 ? 'trending_down' : 'trending_up')}${ch >= 0 ? '+' : ''}${pct(ch, 1)} dibanding bulan lalu</span></div>
    </div>
    <div class="hero-side">
      <div class="glass-ic">${ic('account_balance_wallet')}</div>
      <div><div class="k">ARUS KAS MASUK (${BLN[+mk.slice(5) - 1].toUpperCase()})</div><div class="v money">${rp(a.in)}</div>
      <div class="m">${ic('check_circle')}Target Tabungan Aktif: ${S.target.filter((t) => t.status === 'aktif').length}</div></div>
    </div>
    <div class="hero-mini">
      <div><small>${ic('arrow_downward')}Pemasukan</small><b class="money">${rpS(a.in)}</b></div>
      <div><small>${ic('arrow_upward')}Pengeluaran</small><b class="money">${rpS(a.out)}</b></div>
    </div>
  </div>

  <div class="quick a-quick">
    <button class="qa" data-act="tx-new" data-tipe="keluar"><span class="icon-well">${ic('add_shopping_cart')}</span><span><b>Catat Transaksi</b><small>Pemasukan / Pengeluaran</small></span></button>
    <button class="qa" data-act="tx-new" data-tipe="transfer"><span class="icon-well" style="${tint('#3B82F6')}">${ic('swap_horiz')}</span><span><b>Transfer Dompet</b><small>Antar rekening &amp; e-wallet</small></span></button>
    <button class="qa" data-act="bill-quick"><span class="icon-well" style="${tint('#F5B13D')}">${ic('receipt')}</span><span><b>Bayar Tagihan</b><small class="${att.length ? 'alert' : ''}">${unpaid.length ? unpaid.length + ' tagihan menunggu' : 'Semua lunas'}</small></span></button>
    <button class="qa dashed" data-act="target-new"><span class="icon-well" style="${tint('#2DB37A')}">${ic('flag')}</span><span><b>Tambah Target</b><small>Rencana tabungan baru</small></span></button>
  </div>

  <div class="card a-donut">
    <div class="card-h"><div class="grow"><div class="card-t">Pengeluaran per Kategori</div><div class="card-s">Alokasi belanja rumah tangga</div></div>
      <select class="select sm" style="width:auto" data-ch="dash-month" aria-label="Pilih bulan">${months.map((m) => `<option value="${m}" ${m === DASH.mk ? 'selected' : ''}>${mLabel(m)}</option>`).join('')}</select></div>
    <div class="donut-wrap"><div class="donut"><canvas id="chDonut" aria-label="Grafik donat pengeluaran per kategori" role="img"></canvas><div class="donut-center"><small>Total Pengeluaran</small><b class="money">${rpS(dTotal)}</b></div></div>
      <div class="legend">${legendHtml(dItems)}</div></div>
  </div>

  <div class="card a-bar">
    <div class="card-h"><div class="grow"><div class="card-t">Pengeluaran Bulanan ${T.slice(0, 4)}</div><div class="card-s">Tren belanja Januari – Desember</div></div>
      <div class="chart-legend"><span><i style="background:var(--primary)"></i>Aktual</span><span><i class="dash"></i>Proyeksi</span></div></div>
    <div class="chart-box"><canvas id="chBar" role="img" aria-label="Grafik batang pengeluaran bulanan"></canvas></div>
  </div>

  <div class="card insight a-insight">
    <span class="icon-well">${ic('auto_awesome')}</span>
    <div><div class="card-t">Insight Otomatis <span class="pill pill-ok">Analisis Cerdas</span></div><div class="card-s">Dihitung dari data transaksi, budget, dan tagihan keluarga</div>
      <div class="insight-list">${buildInsights().map((i) => `<div class="insight-item"><span class="ms" style="color:${i.c}">${i.icon}</span><div>${i.html}</div></div>`).join('')}</div></div>
  </div>

  <div class="card a-dompet">
    <div class="card-h"><span class="icon-well sm">${ic('account_balance_wallet')}</span><div class="grow"><div class="card-t">Sebaran Saldo Dompet</div><div class="card-s">${dompet.length} akun aktif</div></div><button class="link" data-go="dompet">Lihat Semua ${ic('arrow_forward')}</button></div>
    <div class="list">${dompet.length ? dompet.map((d) => {
      const p = total > 0 ? Math.max(0, d.saldo) / total : 0;
      return `<div class="li click" data-act="dp-mutasi" data-id="${d.id}">${dpBadge(d)}<div class="grow"><div class="row" style="justify-content:space-between"><span class="semi ellipsis">${esc(d.nama)}</span><b class="money">${rp(d.saldo)}</b></div>
        <div class="row mt-1" style="gap:10px"><div class="grow"><div class="progress thin"><i style="width:${(p * 100).toFixed(1)}%;background:${d.warna}"></i></div></div><span class="xs faint">${pct(p)}</span></div></div></div>`;
    }).join('') : emptyBox('account_balance_wallet', 'Belum ada dompet', '', '<button class="btn btn-soft btn-sm" data-act="dp-new">Tambah Dompet</button>')}</div>
  </div>

  <div class="card a-budget">
    <div class="card-h"><span class="icon-well sm">${ic('donut_large')}</span><div class="grow"><div class="card-t">Progres Budget ${BULAN[+mk.slice(5) - 1]}</div><div class="card-s">${bud.length} pos anggaran</div></div>
      <span class="pill ${bp >= 1 ? 'pill-danger' : bp >= 0.8 ? 'pill-warn' : 'pill-ok'}">${bp >= 1 ? 'Melebihi Kuota' : bp >= 0.8 ? 'Mendekati Batas' : 'Di Bawah Kuota'}</span></div>
    ${lim ? `<div class="mini-box row mb-2"><div class="ring">${ring(bp)}<span class="rt">${pct(bp)}</span></div>
      <div class="grow"><div class="xs faint">Status Penggunaan</div><div class="semi">${pct(bp)} Terpakai</div></div>
      <div class="right"><div class="xs faint">Sisa Anggaran</div><div class="bold money" style="font-size:18px;color:var(--primary-text)">${rp(lim - used)}</div></div></div>
      ${bud.slice(0, 4).map((x) => `<div class="bud-row"><div class="row"><b>${esc(x.k.nama)}</b><span class="${x.p >= 1 ? 'neg' : x.p >= 0.8 ? '' : 'muted'}" style="${x.p >= 0.8 && x.p < 1 ? 'color:var(--gold-700)' : ''}">${pct(x.p)} <span class="faint money">(${rpS(x.used)} / ${rpS(x.limit)})</span></span></div>${progress(x.p)}</div>`).join('')}`
      : emptyBox('donut_large', 'Budget bulan ini belum diatur', 'Tetapkan pagu agar pengeluaran terkendali.', '<button class="btn btn-soft btn-sm" data-act="budget-edit">Atur Budget</button>')}
  </div>

  <div class="card a-tagihan ${urgent ? 'urgent' : ''}">
    <div class="card-h"><span class="icon-well sm" style="${tint('#EB5757')}">${ic('event_upcoming')}</span><div class="grow"><div class="card-t">Tagihan Mendatang</div><div class="card-s">${urgent ? '<span class="neg">Urgensi pembayaran</span>' : 'Pantau jatuh tempo'}</div></div>
      ${att.length ? `<span class="pill pill-danger">${att.length} Jatuh Tempo</span>` : '<span class="pill pill-ok">Aman</span>'}</div>
    <div class="stack" style="gap:10px">${unpaid.slice(0, 3).map((g) => {
      const [txt, cls] = dueText(g);
      const k = S.kMap.get(g.kategori_id);
      return `<button class="bill-item" data-act="bill-pay" data-id="${g.id}">${katWell(k, 'sm')}<span class="grow"><span class="semi ellipsis" style="display:block">${esc(g.nama)}</span><span class="due ${cls}">${cls === 'red' ? '<span class="dot" style="width:6px;height:6px;border-radius:50%;background:currentColor"></span>' : ''}${esc(txt)}</span></span><b class="money">${rp(g.nominal)}</b></button>`;
    }).join('') || emptyBox('task_alt', 'Tidak ada tagihan tertunda')}</div>
    <button class="btn btn-primary btn-block mt-3" data-go="tagihan">${ic('receipt_long')}Lihat Semua Tagihan</button>
  </div>

  <div class="card a-tx">
    <div class="card-h"><div class="grow"><div class="card-t">Transaksi Terbaru</div><div class="card-s">Aktivitas pengeluaran &amp; pemasukan terkini</div></div><button class="link" data-go="transaksi">Lihat Semua ${ic('arrow_forward')}</button></div>
    ${recent.length ? `<div class="hide-m"><div class="tbl-wrap"><table class="tbl"><thead><tr><th>Tanggal &amp; Deskripsi</th><th>Kategori</th><th>Dompet</th><th>Oleh</th><th class="r">Jumlah</th><th></th></tr></thead><tbody>
      ${recent.map((t) => `<tr><td><div class="row">${txWell(t, 'sm')}<div class="grow" style="min-width:150px"><div class="t-title ellipsis">${esc(txTitle(t))}</div><div class="t-sub">${tgl(t.tanggal)}${jam(t.dibuat_pada) ? ', ' + jam(t.dibuat_pada) + ' WIB' : ''}</div></div></div></td>
        <td>${txKatPill(t)}</td><td class="small muted nowrap">${txDompet(t)}</td><td>${whoMini(t.dibuat_oleh)}</td><td class="r">${txAmt(t)}</td>
        <td class="r"><button class="btn-icon" data-act="tx-menu" data-id="${t.id}" aria-label="Aksi">${ic('more_vert')}</button></td></tr>`).join('')}
      </tbody></table></div></div>
      <div class="show-m list" style="display:none">${recent.map(txListItem).join('')}</div>`
      : emptyBox('receipt_long', 'Belum ada transaksi', 'Catat transaksi pertama Anda.', '<button class="btn btn-primary btn-sm" data-act="tx-new">Catat Transaksi</button>')}
  </div>

  <div class="card a-target">
    <div class="card-h"><span class="icon-well sm">${ic('savings')}</span><div class="grow"><div class="card-t">Target Tabungan</div></div><button class="link" data-go="tabungan">Kelola ${ic('arrow_forward')}</button></div>
    <div class="stack" style="gap:10px">${targets.map((t) => {
      const p = targetPct(t), done = p >= 1;
      return done ? `<div class="target-item done row"><span class="icon-well sm round" style="background:var(--primary);color:#fff">${ic('check')}</span><div class="grow"><div class="semi">${esc(t.nama)}</div><div class="xs faint money">${rp(t.terkumpul)} / ${rp(t.target_nominal)}</div></div><span class="pill pill-outline">Lunas 100%</span></div>`
        : `<div class="target-item"><div class="row mb-2">${ic(t.ikon || 'savings', 'faint')}<span class="semi grow ellipsis">${esc(t.nama)}</span><span class="pill pill-ok">${pct(p)}</span></div>${progress(p, '', 'primary')}<div class="row mt-1 xs faint" style="justify-content:space-between"><span>Terkumpul: <b class="money" style="color:var(--text)">${rp(t.terkumpul)}</b></span><span class="money">Target: ${rpS(t.target_nominal)}</span></div></div>`;
    }).join('') || emptyBox('flag', 'Belum ada target', '', '<button class="btn btn-soft btn-sm" data-act="target-new">Buat Target</button>')}</div>
  </div>
  </div>`;

  // grafik
  drawDonut('chDonut', dItems);
  whenChart(() => {
    const c = CC();
    const y = T.slice(0, 4), cur = +mk.slice(5) - 1;
    const actual = Array.from({ length: 12 }, (_, i) => (i <= cur ? monthAgg(`${y}-${String(i + 1).padStart(2, '0')}`).out : null));
    const done = actual.slice(0, cur).filter((v) => v > 0);
    const last3 = actual.slice(Math.max(0, cur - 3), cur).filter((v) => v > 0);
    const proj = last3.length ? sum(last3, (v) => v) / last3.length : actual[cur] || 0;
    const avg = done.length ? sum(done, (v) => v) / done.length : 0;
    mkChart('chBar', {
      type: 'bar',
      data: {
        labels: BLN,
        datasets: [
          { label: 'Aktual', data: actual, backgroundColor: actual.map((_, i) => (i === cur ? c.primary : c.p50)), hoverBackgroundColor: c.accent, borderRadius: 6, borderSkipped: false, barPercentage: 0.72, categoryPercentage: 0.82, grouped: false },
          { label: 'Proyeksi', data: Array.from({ length: 12 }, (_, i) => (i > cur ? proj : null)), backgroundColor: 'transparent', kdDashed: c.accent, borderRadius: 6, barPercentage: 0.72, categoryPercentage: 0.82, grouped: false }
        ]
      },
      options: {
        layout: { padding: { top: 30 } },
        scales: { x: xAxis(c), y: { ...yAxis(c), display: true } },
        plugins: { tooltip: { ...TIP(), callbacks: { label: (x) => ` ${x.dataset.label}: ${rp(x.parsed.y)}` } }, kdAvg: { value: avg, color: c.text, label: avg ? 'Rata-rata: ' + rp(avg) : '' }, kdTag: { index: cur, text: rp(actual[cur] || 0), bg: c.primary } }
      },
      plugins: [plugDashed, plugAvg, plugTag]
    });
  });
}
CHG['dash-month'] = (el) => { DASH.mk = el.value; renderDashboard(); };

// ════════════════════════════════════════════════════════════
// 10. MODAL CATAT TRANSAKSI (pengeluaran / pemasukan / transfer)
// ════════════════════════════════════════════════════════════
let TXM = null;
function openTxModal(opt = {}) {
  const edit = opt.tx ? { ...opt.tx } : null;
  const last = LS.get('kd_last', {});
  TXM = {
    edit, tipe: edit ? edit.tipe : opt.tipe || 'keluar',
    kat: edit ? edit.kategori_id : '', oleh: edit ? edit.dibuat_oleh : S.me.email,
    dompet: edit ? edit.dompet_id : opt.dompet || '', tujuan: edit ? edit.dompet_tujuan_id : opt.tujuan || '', last
  };
  if (!dompetAktif().length) { toast('Buat dompet terlebih dahulu.', 'warn'); goPage('dompet'); return; }
  openModal({
    title: edit ? 'Ubah Transaksi' : 'Catat Transaksi', icon: edit ? 'edit' : 'add_circle',
    body: `
      <div class="seg block solid mb-3" role="tablist" id="txmSeg">
        <button data-act="txm-tipe" data-v="keluar">${ic('arrow_upward')}Pengeluaran</button>
        <button data-act="txm-tipe" data-v="masuk">${ic('arrow_downward')}Pemasukan</button>
        <button data-act="txm-tipe" data-v="transfer">${ic('swap_horiz')}Transfer</button>
      </div>
      <div class="amount-box"><div class="label-xs" id="txmAmtLbl">Jumlah Pengeluaran <span class="neg">*</span></div>
        <div class="amount-row"><span class="cur">Rp</span><input class="amount-input" id="txmAmt" inputmode="numeric" autocomplete="off" data-in="money" data-words="#txmWords" data-after="txm-check" placeholder="0" value="${edit ? fmtN(edit.nominal) : ''}" aria-label="Nominal" autofocus></div>
        <div class="amount-words" id="txmWords">${edit ? terbilang(edit.nominal) : ''}</div></div>
      <div class="form-grid mt-3">
        <div class="field"><label for="txmTgl">Tanggal <span class="req">*</span></label><input type="date" class="input" id="txmTgl" value="${edit ? edit.tanggal : opt.tanggal || today()}" required></div>
        <div class="field"><label for="txmDp" id="txmDpLbl">Dompet / Rekening Sumber <span class="req">*</span></label><select class="select" id="txmDp" data-ch="txm-dp"></select></div>
        <div class="field full" id="txmToWrap"><label for="txmTo">Dompet Tujuan <span class="req">*</span></label><select class="select" id="txmTo"></select></div>
      </div>
      <div class="hint mt-1" id="txmWarn"></div>
      <div class="field mt-3" id="txmKatWrap"><span class="lbl">Kategori <span class="req">*</span></span><div class="opt-grid" id="txmKat"></div><div class="err-msg" id="txmKatErr">Pilih kategori.</div></div>
      <div class="field mt-3"><label for="txmNote">Catatan / Deskripsi Tambahan</label><textarea class="textarea" id="txmNote" maxlength="300" placeholder="Contoh: Pembayaran Listrik PLN Bulan ${BULAN[+today().slice(5, 7) - 1]}">${esc(edit ? edit.catatan : opt.catatan || '')}</textarea></div>
      <div class="field mt-3"><span class="lbl">Dibuat Oleh <span class="req">*</span></span><div class="who-grid" id="txmWho"></div></div>`,
    foot: `${edit ? `<button class="btn btn-danger-soft" data-act="txm-del">${ic('delete')}Hapus</button>` : '<button class="btn btn-ghost" data-act="modal-close">Batal</button>'}
      <div class="grow"></div>
      ${edit ? '<button class="btn btn-ghost" data-act="modal-close">Batal</button>' : '<button class="btn btn-outline" data-act="txm-save" data-more="1">Simpan &amp; Tambah Lagi</button>'}
      <button class="btn btn-primary" data-act="txm-save">${ic('check')}Simpan Transaksi</button>`,
    onOpen: () => txmRender()
  });
}
function txmRender() {
  const m = TXM, tipe = m.tipe;
  $$('#txmSeg button').forEach((b) => b.classList.toggle('on', b.dataset.v === tipe));
  $('#txmAmtLbl').innerHTML = `${tipe === 'masuk' ? 'Jumlah Pemasukan' : tipe === 'transfer' ? 'Nominal Transfer' : 'Jumlah Pengeluaran'} <span class="neg">*</span>`;
  $('#txmDpLbl').innerHTML = `${tipe === 'masuk' ? 'Masuk ke Dompet' : tipe === 'transfer' ? 'Dari Dompet (Sumber)' : 'Dompet / Rekening Sumber'} <span class="req">*</span>`;
  const list = dompetAktif().slice();
  if (m.edit) [m.edit.dompet_id, m.edit.dompet_tujuan_id].forEach((id) => { const d = S.dMap.get(id); if (d && !d.aktif) list.push(d); });
  const dp = m.dompet || m.last[tipe] || list[0]?.id;
  m.dompet = dp;
  $('#txmDp').innerHTML = dompetOpts(dp, true, '', list);
  $('#txmToWrap').hidden = tipe !== 'transfer';
  $('#txmKatWrap').hidden = tipe === 'transfer';
  if (tipe === 'transfer') {
    const to = m.tujuan && m.tujuan !== dp ? m.tujuan : list.find((d) => d.id !== dp)?.id;
    $('#txmTo').innerHTML = dompetOpts(to, true, '', list);
  } else {
    const ktipe = tipe === 'masuk' ? 'pemasukan' : 'pengeluaran';
    if (!m.kat || S.kMap.get(m.kat)?.tipe !== ktipe) m.kat = m.edit && m.edit.tipe === tipe ? m.edit.kategori_id : '';
    const usage = {};
    S.tx.slice(0, 400).forEach((t) => (usage[t.kategori_id] = (usage[t.kategori_id] || 0) + 1));
    const cats = katByTipe(ktipe).slice().sort((a, b) => (a.id === 'kat_lain_k' || a.id === 'kat_lain_m') - (b.id === 'kat_lain_k' || b.id === 'kat_lain_m') || (usage[b.id] || 0) - (usage[a.id] || 0));
    $('#txmKat').innerHTML = cats.map((k) => `<button type="button" class="opt ${k.id === m.kat ? 'on' : ''}" data-act="txm-kat" data-id="${k.id}" aria-pressed="${k.id === m.kat}"><span class="icon-well" style="${tint(k.warna)}">${ic(k.ikon)}</span><span>${esc(k.nama)}</span></button>`).join('')
      + `<button type="button" class="opt" data-act="kat-new" data-tipe="${ktipe}" data-from="txm"><span class="icon-well" style="background:var(--surface-3);color:var(--text-2)">${ic('add')}</span><span>Kategori Baru</span></button>`;
  }
  $('#txmWho').innerHTML = S.users.map((u, i) => `<button type="button" class="who ${u.email === m.oleh ? 'on' : ''}" data-act="txm-who" data-email="${esc(u.email)}"><span class="radio"></span>${avatar(u.email)}<span class="grow"><b>${esc(u.nama)}${u.email === S.me.email ? ' (Saya)' : ''}</b><small>${esc(u.label || (i ? 'Istri' : 'Suami'))}</small></span></button>`).join('');
  ON['txm-check']();
}
ON['txm-check'] = () => {
  const w = $('#txmWarn');
  if (!w || !TXM) return;
  const n = moneyVal('#txmAmt'), d = S.dMap.get($('#txmDp').value);
  let saldo = d ? d.saldo : 0;
  if (TXM.edit && TXM.edit.dompet_id === d?.id && TXM.edit.tipe !== 'masuk') saldo += TXM.edit.nominal;
  w.innerHTML = TXM.tipe !== 'masuk' && d && n > saldo ? `<span style="color:var(--gold-700)">${ic('warning', 'sm')} Saldo ${esc(d.nama)} (${rp(saldo)}) lebih kecil dari nominal — saldo akan minus.</span>` : '';
};
CHG['txm-dp'] = (el) => { TXM.dompet = el.value; if (TXM.tipe === 'transfer') txmRender(); else ON['txm-check'](); };
ACT['txm-tipe'] = (el) => { TXM.tipe = el.dataset.v; TXM.dompet = TXM.edit && TXM.edit.tipe === TXM.tipe ? TXM.edit.dompet_id : ''; txmRender(); };
ACT['txm-kat'] = (el) => { TXM.kat = el.dataset.id; $$('#txmKat .opt').forEach((b) => { b.classList.toggle('on', b.dataset.id === TXM.kat); b.setAttribute('aria-pressed', b.dataset.id === TXM.kat); }); $('#txmKatWrap').classList.remove('invalid'); };
ACT['txm-who'] = (el) => { TXM.oleh = el.dataset.email; $$('#txmWho .who').forEach((b) => b.classList.toggle('on', b.dataset.email === TXM.oleh)); };
ACT['txm-save'] = (el) => {
  const m = TXM, nominal = moneyVal('#txmAmt'), tanggal = $('#txmTgl').value, dp = $('#txmDp').value;
  if (!(nominal > 0)) { toast('Nominal wajib diisi angka positif.', 'warn'); $('#txmAmt').focus(); return; }
  if (!validDate(tanggal)) { toast('Tanggal tidak valid.', 'warn'); return; }
  if (!dp) { toast('Pilih dompet.', 'warn'); return; }
  const tx = { id: m.edit ? m.edit.id : uid('tx'), tanggal, tipe: m.tipe, dompet_id: dp, dompet_tujuan_id: '', kategori_id: '', nominal, catatan: $('#txmNote').value.trim().slice(0, 300), dibuat_oleh: m.oleh, ref: m.edit ? m.edit.ref : '' };
  if (m.tipe === 'transfer') {
    tx.dompet_tujuan_id = $('#txmTo').value;
    if (!tx.dompet_tujuan_id || tx.dompet_tujuan_id === dp) { toast('Dompet tujuan harus berbeda dari dompet asal.', 'warn'); return; }
  } else {
    if (!m.kat) { $('#txmKatWrap').classList.add('invalid'); toast('Pilih kategori transaksi.', 'warn'); return; }
    tx.kategori_id = m.kat;
  }
  if (m.edit) tx.dibuat_pada = m.edit.dibuat_pada;
  DB.simpanTx(tx);
  LS.set('kd_last', { ...LS.get('kd_last', {}), [m.tipe]: dp });
  toast(m.edit ? 'Transaksi diperbarui.' : `${m.tipe === 'transfer' ? 'Transfer' : 'Transaksi'} ${rp(nominal)} tersimpan.`, 'success');
  if (el.dataset.more) {
    $('#txmAmt').value = ''; $('#txmWords').textContent = ''; $('#txmNote').value = ''; $('#txmAmt').focus(); ON['txm-check']();
  } else closeModal();
};
ACT['txm-del'] = async () => {
  const t = TXM.edit;
  if (await confirmBox({ title: 'Hapus Transaksi?', text: `Transaksi <b>${esc(txTitle(t))}</b> sebesar <b>${rp(t.nominal)}</b> akan dihapus permanen dan saldo dompet disesuaikan.`, ok: 'Hapus', danger: true })) {
    DB.hapusTx(t.id); toast('Transaksi dihapus.', 'success');
  }
};
ACT['tx-new'] = (el) => { closeDrawer(); openTxModal({ tipe: el?.dataset?.tipe || 'keluar', dompet: el?.dataset?.dompet || '' }); };
ACT['tx-edit'] = (el) => { closePopover(); const t = S.txMap.get(el.dataset.id); if (t) openTxModal({ tx: t }); };
ACT['tx-del'] = async (el) => {
  closePopover();
  const t = S.txMap.get(el.dataset.id);
  if (!t) return;
  if (await confirmBox({ title: 'Hapus Transaksi?', text: `Transaksi <b>${esc(txTitle(t))}</b> sebesar <b>${rp(t.nominal)}</b> pada ${tgl(t.tanggal, 'long')} akan dihapus permanen dan saldo dompet disesuaikan.`, ok: 'Hapus', danger: true })) {
    DB.hapusTx(t.id); toast('Transaksi dihapus.', 'success');
  }
};
ACT['tx-menu'] = (el) => {
  openPopover(el, `<button class="pop-i" data-act="tx-edit" data-id="${el.dataset.id}">${ic('edit', 'faint')}<span class="grow"><b>Ubah</b></span></button>
    <button class="pop-i" data-act="tx-dup" data-id="${el.dataset.id}">${ic('content_copy', 'faint')}<span class="grow"><b>Duplikat (hari ini)</b></span></button>
    <button class="pop-i" data-act="tx-del" data-id="${el.dataset.id}">${ic('delete', 'neg')}<span class="grow"><b class="neg">Hapus</b></span></button>`, 220);
};
ACT['tx-dup'] = (el) => {
  closePopover();
  const t = S.txMap.get(el.dataset.id);
  if (!t) return;
  DB.simpanTx({ ...t, id: uid('tx'), tanggal: today(), dibuat_oleh: S.me.email, dibuat_pada: '', ref: '' });
  toast('Transaksi diduplikat ke hari ini.', 'success');
};
function txListItem(t) {
  return `<div class="li click" data-act="tx-edit" data-id="${t.id}">${txWell(t)}<div class="grow"><div class="t-title ellipsis">${esc(txTitle(t))}</div>
    <div class="t-sub row" style="gap:6px;flex-wrap:wrap">${relHari(t.tanggal) === 'Hari ini' ? 'Hari ini' : tgl(t.tanggal, 'dm')}${jam(t.dibuat_pada) ? ', ' + jam(t.dibuat_pada) : ''} • <span class="pill" style="height:18px;font-size:10.5px">diinput ${esc(userOf(t.dibuat_oleh).nama)}</span></div></div>
    <div class="right">${txAmt(t)}<div class="t-sub">${txDompet(t)}</div></div></div>`;
}

// ════════════════════════════════════════════════════════════
// 11. HALAMAN TRANSAKSI
// ════════════════════════════════════════════════════════════
const TXF = { q: '', period: 'bulan', dari: '', sampai: '', tipe: '', dompet: '', kat: '', oleh: '', page: 1, per: 15 };
PAGES.transaksi = {
  t: 'Transaksi', r: renderTransaksi,
  prep: (o) => {
    if (o.keep) return;
    if (o.kat !== undefined || o.dompet !== undefined || o.month) {
      Object.assign(TXF, { tipe: '', dompet: o.dompet || '', kat: o.kat || '', oleh: '', page: 1, q: '' });
      if (o.month) { TXF.period = 'custom'; TXF.dari = o.month + '-01'; TXF.sampai = mEnd(o.month); }
    }
  }
};
function txfRange() {
  const T = today(), mk = mKey(T);
  switch (TXF.period) {
    case 'lalu': { const p = mAdd(mk, -1); return [p + '-01', mEnd(p)]; }
    case '30': return [addDays(T, -29), T];
    case 'tahun': return [T.slice(0, 4) + '-01-01', T.slice(0, 4) + '-12-31'];
    case 'semua': return [S.txDari, '9999-12-31'];
    case 'custom': return [TXF.dari || mk + '-01', TXF.sampai || T];
    default: return [mk + '-01', mEnd(mk)];
  }
}
const txSearchText = (t) => `${t.catatan} ${S.kMap.get(t.kategori_id)?.nama || ''} ${S.dMap.get(t.dompet_id)?.nama || ''} ${S.dMap.get(t.dompet_tujuan_id)?.nama || ''} ${userOf(t.dibuat_oleh).nama} ${t.nominal} ${fmtN(t.nominal)}`.toLowerCase();
function txFiltered(ignoreTipe) {
  const [a, b] = txfRange();
  const q = TXF.q.trim().toLowerCase();
  return S.tx.filter((t) => t.tanggal >= a && t.tanggal <= b && (ignoreTipe || !TXF.tipe || t.tipe === TXF.tipe) && (!TXF.dompet || t.dompet_id === TXF.dompet || t.dompet_tujuan_id === TXF.dompet)
    && (!TXF.kat || t.kategori_id === TXF.kat) && (!TXF.oleh || t.dibuat_oleh === TXF.oleh) && (!q || txSearchText(t).includes(q)));
}
function renderTransaksi() {
  const [a] = txfRange();
  if (a < S.txDari) ensureLoaded(a).then(() => { if (CUR === 'transaksi') renderTxResults(); });
  $('#page-transaksi').innerHTML = `
  <div class="page-head"><div class="grow"><div class="crumb"><span class="tag">${ic('receipt_long')}Buku Kas Keluarga</span><span>Sinkronisasi otomatis</span></div>
    <h1 class="page-title">Daftar Transaksi</h1><p class="page-sub">Catatan pemasukan, pengeluaran, dan transfer keluarga ${esc(familyName())} secara real-time.</p></div>
    <div class="head-actions"><button class="btn btn-outline" data-act="tx-new" data-tipe="transfer">${ic('swap_horiz')}Transfer Antar Dompet</button><button class="btn btn-primary" data-act="tx-new">${ic('add')}Catat Transaksi</button></div></div>
  <div class="grid g-3 kpi-row" id="txKpi"></div>
  <div class="card mt-3">
    <div class="filter-card">
      <div class="input-icon grow" style="min-width:220px">${ic('search')}<input class="input sm" type="search" id="txQ" placeholder="Cari catatan, kategori, dompet, nominal…" value="${esc(TXF.q)}" data-in="txf-q" aria-label="Cari transaksi"></div>
      <select class="select sm" style="width:auto" data-ch="txf" data-k="period" aria-label="Periode">
        ${[['bulan', 'Bulan Ini'], ['lalu', 'Bulan Lalu'], ['30', '30 Hari Terakhir'], ['tahun', 'Tahun Ini'], ['semua', 'Semua (12+ bln)'], ['custom', 'Rentang Custom']].map(([v, l]) => `<option value="${v}" ${TXF.period === v ? 'selected' : ''}>${l}</option>`).join('')}</select>
      <span class="date-range" ${TXF.period === 'custom' ? '' : 'hidden'} id="txRange">${ic('calendar_today')}<input type="date" value="${TXF.dari || mKey(today()) + '-01'}" data-ch="txf" data-k="dari" aria-label="Dari"><span class="faint">s/d</span><input type="date" value="${TXF.sampai || today()}" data-ch="txf" data-k="sampai" aria-label="Sampai"></span>
      <select class="select sm" style="width:auto" data-ch="txf" data-k="dompet" aria-label="Dompet">${optList(dompetSemua(), TXF.dompet, 'Semua Dompet')}</select>
      <select class="select sm" style="width:auto" data-ch="txf" data-k="kat" aria-label="Kategori">${optList(S.kategori.slice().sort((x, y) => x.tipe.localeCompare(y.tipe) || x.nama.localeCompare(y.nama)), TXF.kat, 'Semua Kategori')}</select>
      <select class="select sm" style="width:auto" data-ch="txf" data-k="oleh" aria-label="Diinput oleh">${optList(S.users.map((u) => ({ id: u.email, nama: 'Diinput ' + u.nama })), TXF.oleh, 'Semua Penginput')}</select>
    </div>
    <div class="row mt-2" style="flex-wrap:wrap;gap:10px"><div class="chips" id="txTipe"></div><div class="grow"></div><button class="link" data-act="txf-reset">${ic('filter_alt_off')}Reset filter</button></div>
  </div>
  <div class="card mt-3" id="txList"></div>`;
  $('#globalSearch').value = TXF.q;
  renderTxResults();
}
function renderTxResults() {
  const base = txFiltered(true);
  const rows = TXF.tipe ? base.filter((t) => t.tipe === TXF.tipe) : base;
  const tin = sum(base.filter((t) => t.tipe === 'masuk'), (t) => t.nominal);
  const tout = sum(base.filter((t) => t.tipe === 'keluar'), (t) => t.nominal);
  const net = tin - tout;
  const cnt = (tp) => base.filter((t) => !tp || t.tipe === tp).length;
  $('#txKpi').innerHTML = `
    <div class="card kpi"><div class="kpi-top"><span class="kpi-l">Total Pemasukan</span><span class="icon-well sm">${ic('arrow_downward')}</span></div><div class="kpi-v pos money">${rp(tin)}</div><div class="kpi-f">${base.filter((t) => t.tipe === 'masuk').length} transaksi masuk</div></div>
    <div class="card kpi"><div class="kpi-top"><span class="kpi-l">Total Pengeluaran</span><span class="icon-well sm" style="${tint('#EB5757')}">${ic('arrow_upward')}</span></div><div class="kpi-v neg money">${rp(tout)}</div><div class="kpi-f">${base.filter((t) => t.tipe === 'keluar').length} transaksi keluar</div></div>
    <div class="card kpi hl"><div class="kpi-top"><span class="kpi-l">Selisih Bersih (Net Cashflow)</span><span class="pill ${net >= 0 ? 'pill-solid' : 'pill-danger'}">${net >= 0 ? 'Surplus' : 'Defisit'}</span></div><div class="kpi-v money" style="color:${net >= 0 ? 'var(--primary-text)' : 'var(--danger-700)'}">${net >= 0 ? '+' : ''}${rp(net)}</div><div class="kpi-f">${ic('check_circle', 'sm')}Transfer antar dompet tidak dihitung</div></div>`;
  $('#txTipe').innerHTML = [['', 'Semua'], ['masuk', 'Pemasukan'], ['keluar', 'Pengeluaran'], ['transfer', 'Transfer']].map(([v, l]) => `<button class="chip ${TXF.tipe === v ? 'on' : ''}" data-act="txf-tipe" data-v="${v}">${l} <span class="cnt">(${cnt(v)})</span></button>`).join('');
  const page = clamp(TXF.page, 1, Math.max(1, Math.ceil(rows.length / TXF.per)));
  TXF.page = page;
  const slice = rows.slice((page - 1) * TXF.per, page * TXF.per);
  const [a] = txfRange();
  const loading = a < S.txDari;
  let lastDate = '';
  $('#txList').innerHTML = loading ? `<div class="stack">${'<div class="skel" style="height:44px"></div>'.repeat(5)}</div>` : rows.length ? `
    <div class="hide-m"><div class="tbl-wrap"><table class="tbl"><thead><tr><th>Tanggal &amp; Deskripsi</th><th>Kategori</th><th>Dompet</th><th>Oleh</th><th class="r">Jumlah</th><th class="r">Aksi</th></tr></thead><tbody>
    ${slice.map((t) => `<tr><td><div class="row">${txWell(t, 'sm')}<div class="grow" style="min-width:160px"><div class="t-title">${esc(txTitle(t))}</div><div class="t-sub">${tgl(t.tanggal)}${jam(t.dibuat_pada) ? ', ' + jam(t.dibuat_pada) + ' WIB' : ''}${t.ref ? ` • ${ic(t.ref.startsWith('tagihan') ? 'receipt' : 'savings', 'sm')}` : ''}</div></div></div></td>
      <td>${txKatPill(t)}</td><td class="small muted nowrap">${txDompet(t)}</td><td>${whoMini(t.dibuat_oleh)}</td><td class="r">${txAmt(t)}</td>
      <td class="r"><div class="row-actions"><button class="btn-icon" data-act="tx-edit" data-id="${t.id}" aria-label="Ubah">${ic('edit')}</button><button class="btn-icon danger" data-act="tx-del" data-id="${t.id}" aria-label="Hapus">${ic('delete')}</button></div></td></tr>`).join('')}
    </tbody></table></div></div>
    <div class="show-m list" style="display:none">${slice.map((t) => { const head = t.tanggal !== lastDate ? `<div class="list-date">${relHari(t.tanggal) === 'Hari ini' ? 'Hari ini' : tgl(t.tanggal, 'day')}</div>` : ''; lastDate = t.tanggal; return head + txListItem(t); }).join('')}</div>
    ${pager(rows.length, page, TXF.per, 'txf-page')}`
    : emptyBox('search_off', 'Tidak ada transaksi', 'Coba ubah filter atau kata kunci pencarian.', '<button class="btn btn-soft btn-sm" data-act="txf-reset">Reset Filter</button>');
}
const _txQ = debounce(() => { TXF.page = 1; renderTxResults(); }, 200);
ON['txf-q'] = (el) => { TXF.q = el.value; $('#globalSearch').value = el.value; _txQ(); };
CHG.txf = (el) => {
  TXF[el.dataset.k] = el.value; TXF.page = 1;
  if (el.dataset.k === 'period') { $('#txRange').hidden = el.value !== 'custom'; if (el.value === 'custom') { TXF.dari = TXF.dari || mKey(today()) + '-01'; TXF.sampai = TXF.sampai || today(); } }
  const [a] = txfRange();
  if (a < S.txDari) { renderTxResults(); ensureLoaded(a).then(() => CUR === 'transaksi' && renderTxResults()); } else renderTxResults();
};
ACT['txf-tipe'] = (el) => { TXF.tipe = el.dataset.v; TXF.page = 1; renderTxResults(); };
ACT['txf-page'] = (el) => { TXF.page = +el.dataset.p; renderTxResults(); $('#txList').scrollIntoView({ behavior: 'smooth', block: 'start' }); };
ACT['txf-reset'] = () => { Object.assign(TXF, { q: '', period: 'bulan', dari: '', sampai: '', tipe: '', dompet: '', kat: '', oleh: '', page: 1 }); $('#globalSearch').value = ''; renderTransaksi(); };

/** Muat transaksi lama dari server bila rentang melewati data yang dimuat awal. */
let _loading = null;
function ensureLoaded(dari) {
  if (!dari || dari >= S.txDari) return Promise.resolve();
  if (_loading) return _loading;
  const sampai = addDays(S.txDari, -1);
  _loading = api('ambilTransaksiRentang', dari, sampai).then((res) => {
    const { c, r } = res.data;
    let n = 0;
    r.forEach((row) => { const o = {}; c.forEach((k, i) => (o[k] = row[i])); o.nominal = +o.nominal || 0; if (!S.txMap.has(o.id)) { S.tx.push(o); S.txMap.set(o.id, o); n++; } });
    S.txDari = dari; sortTx(); bump();
    if (n) toast(`${fmtN(n)} transaksi lama dimuat dari arsip.`, 'info');
  }).catch((e) => toast('Gagal memuat transaksi lama: ' + e.message, 'error')).finally(() => { _loading = null; });
  return _loading;
}


// ════════════════════════════════════════════════════════════
// 12. HALAMAN DOMPET
// ════════════════════════════════════════════════════════════
const WARNA = [
  ['#0F5C3F', 'Forest Emerald'], ['#2DB37A', 'Emerald Green'], ['#5FD3A0', 'Mint Leaf'], ['#14B8A6', 'Teal Lagoon'], ['#3B82F6', 'Royal Blue'], ['#6366F1', 'Indigo'],
  ['#9B51E0', 'Purple Iris'], ['#EC4899', 'Rose Pink'], ['#EB5757', 'Coral Red'], ['#F2994A', 'Orange Sunset'], ['#F5B13D', 'Golden Sun'], ['#C98516', 'Amber'], ['#64748B', 'Slate Gray'], ['#94A3B8', 'Silver']
];
const warnaNama = (hex) => (WARNA.find((w) => w[0].toLowerCase() === String(hex).toLowerCase()) || [hex, String(hex).toUpperCase()])[1];
const IKON_KAT = ['restaurant', 'local_cafe', 'shopping_cart', 'shopping_bag', 'bolt', 'water_drop', 'wifi', 'phone_iphone', 'school', 'menu_book', 'child_care', 'directions_car', 'two_wheeler', 'local_gas_station', 'medical_services', 'fitness_center', 'movie', 'sports_esports', 'flight_takeoff', 'checkroom', 'home', 'build', 'pets', 'volunteer_activism', 'mosque', 'account_balance', 'credit_card', 'receipt_long', 'payments', 'storefront', 'work', 'redeem', 'card_giftcard', 'attach_money', 'trending_up', 'savings', 'celebration', 'family_restroom', 'spa', 'more_horiz'];
const IKON_TARGET = ['savings', 'shield', 'flight_takeoff', 'home', 'school', 'directions_car', 'mosque', 'celebration', 'child_care', 'medical_services', 'phone_iphone', 'workspace_premium', 'redeem', 'flag'];
function relStamp(st) {
  if (!st) return '-';
  const d = String(st).slice(0, 10), j = jam(st);
  const r = relHari(d);
  return ['Hari ini', 'Kemarin'].includes(r) ? `${r}${j ? ', ' + j + ' WIB' : ''}` : (diffDays(d, today()) < 7 ? r : tgl(d));
}
const txEffect = (t, id) => (t.tipe === 'masuk' && t.dompet_id === id ? t.nominal : t.tipe === 'keluar' && t.dompet_id === id ? -t.nominal : t.tipe === 'transfer' ? (t.dompet_tujuan_id === id ? t.nominal : t.dompet_id === id ? -t.nominal : 0) : 0);

PAGES.dompet = { t: 'Dompet', r: renderDompet };
const TF = { dari: '', ke: '' };
function renderDompet() {
  const aktif = dompetAktif(), total = totalSaldo();
  const arsip = S.dompet.filter((d) => !d.aktif);
  const ch = saldoChange(total);
  const J = { rekening: 0, 'e-wallet': 0, tunai: 0, lainnya: 0 };
  aktif.forEach((d) => (J[d.jenis] = (J[d.jenis] || 0) + d.saldo));
  const pos = Math.max(1, sum(aktif, (d) => Math.max(0, d.saldo)));
  const T = today(), w0 = addDays(T, -6);
  const week = aktif.map((d) => {
    let inn = 0, out = 0;
    S.tx.forEach((t) => { if (t.tanggal < w0 || t.tanggal > T) return; const e = txEffect(t, d.id); if (e > 0) inn += e; else out -= e; });
    return { d, inn, out };
  });
  if (!TF.dari || !S.dMap.get(TF.dari)?.aktif) TF.dari = aktif[0]?.id || '';
  if (!TF.ke || TF.ke === TF.dari || !S.dMap.get(TF.ke)?.aktif) TF.ke = aktif.find((d) => d.id !== TF.dari)?.id || '';

  $('#page-dompet').innerHTML = `
  <div class="page-head"><div class="grow"><div class="crumb"><span class="tag">${ic('account_balance_wallet')}Aset Keluarga</span><span>• Sinkronisasi Otomatis</span></div>
    <h1 class="page-title">Kelola Dompet &amp; Rekening</h1><p class="page-sub">Pantau saldo dan sebaran aset keuangan keluarga di semua rekening, e-wallet, dan uang tunai secara real-time.</p></div>
    <div class="head-actions"><button class="btn btn-outline" data-act="tf-scroll">${ic('swap_horiz')}Transfer Antar Dompet</button><button class="btn btn-primary" data-act="dp-new">${ic('add')}Tambah Dompet Baru</button></div></div>

  <div class="card">
    <div class="row top" style="flex-wrap:wrap;gap:20px">
      <div class="grow" style="min-width:260px">
        <div class="row" style="gap:8px;flex-wrap:wrap"><span class="dot" style="width:8px;height:8px;border-radius:50%;background:var(--accent)"></span><span class="label-xs" style="color:var(--text-2)">Total Saldo Bersih Seluruh Dompet</span>
          <span class="pill ${ch >= 0 ? 'pill-ok' : 'pill-danger'}">${ic(ch >= 0 ? 'trending_up' : 'trending_down')}${ch >= 0 ? '+' : ''}${pct(ch, 1)} vs bln lalu</span></div>
        <div class="row mt-1" style="align-items:baseline;gap:10px;flex-wrap:wrap"><span class="money" style="font:800 34px/1.2 var(--font-head);letter-spacing:-.02em">${rp(total)}</span><span class="small faint">IDR (Rupiah)</span></div>
        <div class="small muted mt-1">Tersebar di <b>${aktif.length} dompet aktif</b> (rekening bank, dompet digital &amp; kas tunai).</div>
      </div>
      <div class="row" style="gap:10px;flex-wrap:wrap">
        ${[['rekening', 'Perbankan'], ['e-wallet', 'E-Wallet'], ['tunai', 'Uang Tunai'], ['lainnya', 'Lainnya']].filter(([k]) => k !== 'lainnya' || J.lainnya).map(([k, l]) => `<div class="type-box"><div class="k">${l}</div><div class="v">${pct(Math.max(0, J[k]) / pos)}</div><div class="s money">${rp(J[k])}</div></div>`).join('')}
      </div>
    </div>
    <div class="sep"></div>
    <div class="row mb-2" style="justify-content:space-between"><span class="semi small">Komposisi Distribusi Saldo</span><span class="xs faint">100% teralokasi</span></div>
    <div class="dist-bar">${aktif.filter((d) => d.saldo > 0).map((d) => `<i style="flex-grow:${d.saldo};background:${d.warna}" title="${esc(d.nama)} ${pct(d.saldo / pos)}"></i>`).join('') || '<i style="flex-grow:1;background:var(--surface-3)"></i>'}</div>
    <div class="dist-legend">${aktif.map((d) => `<span><i style="background:${d.warna}"></i>${esc(d.nama)} <em>${pct(Math.max(0, d.saldo) / pos)} (${rpS(d.saldo)})</em></span>`).join('')}</div>
  </div>

  <div class="section-h"><div class="grow row"><h2>Daftar Dompet &amp; Akun Kas</h2><span class="pill pill-ok">${aktif.length} Aktif</span></div><span class="xs faint hide-m">Klik "Mutasi" untuk melihat rincian transaksi per dompet</span></div>
  <div class="grid g-3">
    ${aktif.map((d) => {
      const p = Math.max(0, d.saldo) / pos;
      return `<div class="card wallet-card">
        <div class="row top">${dpBadge(d)}<div class="grow"><div class="semi" style="font:600 15.5px/1.3 var(--font-head)">${esc(d.nama)}</div><div class="xs faint">${JENIS_DP[d.jenis]?.[0] || d.jenis}${d.keterangan ? ' • ' + esc(d.keterangan) : ''}</div></div><span class="pill pill-ok"><span class="dot"></span>Aktif</span></div>
        <div style="border-top:1px solid var(--divider);padding-top:12px"><div class="label-xs">Saldo Saat Ini</div>
          <div class="row" style="justify-content:space-between;align-items:baseline;gap:8px;flex-wrap:wrap"><span class="saldo money ${d.saldo < 0 ? 'neg' : ''}">${rp(d.saldo)}</span><span class="xs faint">${pct(p)} dari total aset</span></div>
          <div class="mt-1"><div class="progress thin"><i style="width:${(p * 100).toFixed(1)}%;background:${d.warna}"></i></div></div></div>
        <div class="meta-box"><span>Saldo Awal:</span><b class="money">${rp(d.saldo_awal)}</b><span>Terakhir Diupdate:</span><b>${relStamp(d.terakhir)}</b></div>
        <div class="actions"><button class="btn btn-outline btn-sm" data-act="dp-edit" data-id="${d.id}">${ic('edit')}Ubah</button><button class="btn btn-soft btn-sm" data-act="dp-mutasi" data-id="${d.id}">${ic('receipt_long')}Mutasi</button></div>
      </div>`;
    }).join('')}
    <button class="add-card" data-act="dp-new"><span class="plus">${ic('add', 'lg')}</span><b>Tambah Dompet Baru</b><span class="small">Hubungkan rekening baru, dompet digital, atau kas fisik.</span><span class="pill pill-outline">Rekening • E-Wallet • Tunai</span></button>
  </div>

  <div class="grid g-side mt-4">
    <div class="card" id="tfCard">
      <div class="card-h"><span class="icon-well">${ic('swap_horiz')}</span><div class="grow"><div class="card-t" style="font-size:18px">Transfer Antar Dompet</div><div class="card-s">Pindahkan saldo internal antar rekening keluarga</div></div><span class="pill pill-ok">Bebas Biaya Admin</span></div>
      ${aktif.length < 2 ? emptyBox('swap_horiz', 'Butuh minimal 2 dompet aktif', 'Tambahkan dompet lain untuk mulai transfer.') : `
      <div class="transfer-grid">
        <div class="tf-box"><div class="label-xs">Dari Dompet (Sumber)</div><select class="select" id="tfFrom" data-ch="tf" data-k="dari" aria-label="Dompet sumber">${dompetOpts(TF.dari, false)}</select><div class="xs faint">Saldo: <b class="money" id="tfFromSaldo">${rp(S.dMap.get(TF.dari)?.saldo)}</b></div></div>
        <button class="swap-btn" data-act="tf-swap" aria-label="Tukar arah transfer">${ic('swap_horiz')}</button>
        <div class="tf-box"><div class="label-xs">Ke Dompet (Tujuan)</div><select class="select" id="tfTo" data-ch="tf" data-k="ke" aria-label="Dompet tujuan">${dompetOpts(TF.ke, false)}</select><div class="xs faint">Saldo: <b class="money" id="tfToSaldo">${rp(S.dMap.get(TF.ke)?.saldo)}</b></div></div>
      </div>
      <div class="field mt-3"><label for="tfAmt">Nominal Transfer <span class="req">*</span></label>
        <div class="money-input"><span class="cur">Rp</span><input id="tfAmt" inputmode="numeric" autocomplete="off" placeholder="0" data-in="money" data-words="#tfWords" data-after="tf-est">
          <div class="quick"><button type="button" data-act="tf-q" data-v="250000">250rb</button><button type="button" data-act="tf-q" data-v="500000">500rb</button><button type="button" data-act="tf-q" data-v="1000000">1jt</button></div></div>
        <div class="hint" id="tfWords"></div></div>
      <div class="field mt-3"><label for="tfNote">Catatan / Keperluan (Opsional)</label><input class="input" id="tfNote" maxlength="200" placeholder="Contoh: Top up saldo belanja bulanan"></div>
      <div class="field mt-3"><label for="tfDate">Tanggal</label><input type="date" class="input" id="tfDate" value="${today()}"></div>
      <div class="note mt-3">${ic('lightbulb')}<div><b>Catatan:</b> Transfer antar dompet tidak dihitung sebagai pemasukan atau pengeluaran keluarga. Saldo kedua dompet otomatis disesuaikan secara real-time.</div></div>
      <div class="mt-3"><div class="label-xs mb-2">Estimasi saldo setelah transfer</div><div class="est-box" id="tfEst"></div></div>
      <button class="btn btn-primary btn-lg btn-block mt-3" data-act="tf-go">${ic('send')}Proses Transfer Sekarang</button>`}
    </div>
    <div class="stack" style="gap:20px">
      <div class="card">
        <div class="card-h"><div class="grow"><div class="card-t">Aktivitas Dompet Sepekan</div><div class="card-s">Ringkasan arus kas 7 hari terakhir</div></div>${ic('history', 'faint')}</div>
        <div class="list">${week.map(({ d, inn, out }) => `<div class="li click" data-act="dp-mutasi" data-id="${d.id}">${dpBadge(d, true)}<div class="grow"><div class="semi small ellipsis">${esc(d.nama)}</div><div class="xs faint ellipsis">${esc(d.keterangan || JENIS_DP[d.jenis]?.[0] || '')}</div></div>
          <div class="right xs"><div class="${inn ? 'pos semi' : 'faint'} money">+${rp(inn)}</div><div class="${out ? 'neg' : 'faint'} money">-${rp(out)}</div></div></div>`).join('') || '<div class="small faint">Belum ada dompet aktif.</div>'}</div>
        <button class="link mt-2" data-go="laporan">Lihat Laporan Mutasi Lengkap ${ic('arrow_forward')}</button>
      </div>
      <div class="card">
        <div class="card-t mb-2">Kelola Dompet &amp; Pengaturan</div>
        <div class="stack" style="gap:10px">
          <button class="link-row" data-act="dp-urutan"><span class="icon-well sm">${ic('drag_indicator')}</span><span class="grow"><b>Atur Urutan Tampilan Dompet</b><small>Susun rekening favorit di atas</small></span>${ic('chevron_right', 'faint')}</button>
          <button class="link-row" data-act="dp-arsip"><span class="icon-well sm" style="${tint('#F5B13D')}">${ic('archive')}</span><span class="grow"><b>Arsip Rekening Non-Aktif</b><small>${arsip.length} rekening diarsipkan</small></span>${ic('chevron_right', 'faint')}</button>
          <button class="link-row" data-act="dp-rekon"><span class="icon-well sm" style="${tint('#3B82F6')}">${ic('balance')}</span><span class="grow"><b>Rekonsiliasi Saldo Bank</b><small>Cocokkan saldo aplikasi dengan saldo fisik</small></span>${ic('chevron_right', 'faint')}</button>
        </div>
      </div>
    </div>
  </div>`;
  ON['tf-est']();
}
ON['tf-est'] = () => {
  const box = $('#tfEst');
  if (!box) return;
  const n = moneyVal('#tfAmt'), a = S.dMap.get(TF.dari), b = S.dMap.get(TF.ke);
  if (!a || !b) { box.innerHTML = ''; return; }
  box.innerHTML = `<div class="row">${ic('south', 'neg')}<div><div class="xs faint">${esc(a.nama)}</div><b class="money ${a.saldo - n < 0 ? 'neg' : ''}">${rp(a.saldo - n)}</b></div></div>
    <div class="row">${ic('north', 'pos')}<div><div class="xs faint">${esc(b.nama)}</div><b class="money">${rp(b.saldo + n)}</b></div></div>`;
};
CHG.tf = (el) => {
  TF[el.dataset.k] = el.value;
  if (TF.dari === TF.ke) { TF.ke = dompetAktif().find((d) => d.id !== TF.dari)?.id || ''; $('#tfTo').value = TF.ke; }
  $('#tfFromSaldo').textContent = rp(S.dMap.get(TF.dari)?.saldo);
  $('#tfToSaldo').textContent = rp(S.dMap.get(TF.ke)?.saldo);
  ON['tf-est']();
};
ACT['tf-swap'] = () => { [TF.dari, TF.ke] = [TF.ke, TF.dari]; $('#tfFrom').value = TF.dari; $('#tfTo').value = TF.ke; CHG.tf($('#tfFrom')); };
ACT['tf-q'] = (el) => { const i = $('#tfAmt'); i.value = el.dataset.v; ON.money(i); };
ACT['tf-scroll'] = () => { const c = $('#tfCard'); if (c) { c.scrollIntoView({ behavior: 'smooth', block: 'start' }); setTimeout(() => $('#tfAmt')?.focus(), 400); } };
ACT['tf-go'] = () => {
  const n = moneyVal('#tfAmt'), tanggal = $('#tfDate').value;
  if (!(n > 0)) { toast('Isi nominal transfer.', 'warn'); $('#tfAmt').focus(); return; }
  if (!TF.dari || !TF.ke || TF.dari === TF.ke) { toast('Pilih dompet asal dan tujuan yang berbeda.', 'warn'); return; }
  if (!validDate(tanggal)) { toast('Tanggal tidak valid.', 'warn'); return; }
  const a = S.dMap.get(TF.dari), b = S.dMap.get(TF.ke);
  DB.simpanTx({ id: uid('tx'), tanggal, tipe: 'transfer', dompet_id: TF.dari, dompet_tujuan_id: TF.ke, kategori_id: '', nominal: n, catatan: $('#tfNote').value.trim() || `Transfer ke ${b.nama}`, dibuat_oleh: S.me.email });
  toast(`Transfer ${rp(n)} dari ${a.nama} ke ${b.nama} berhasil.`, 'success');
};

// ── Modal Dompet ──
let DPM = null;
function openDompetModal(d) {
  const edit = d ? { ...d } : null;
  DPM = { edit, jenis: edit ? edit.jenis : 'rekening', warna: edit ? edit.warna : WARNA[4][0], img: null, ikon_url: edit ? edit.ikon_url : '' };
  const used = edit ? S.tx.some((t) => t.dompet_id === edit.id || t.dompet_tujuan_id === edit.id) : false;
  openModal({
    title: edit ? 'Ubah Dompet' : 'Tambah Dompet Baru', icon: 'account_balance_wallet',
    body: `
      <div class="field"><label for="dpNama">Nama Dompet / Rekening <span class="req">*</span></label><input class="input" id="dpNama" maxlength="60" placeholder="Contoh: BCA Rekening Utama" value="${esc(edit?.nama || '')}"></div>
      <div class="field mt-3"><span class="lbl">Jenis <span class="req">*</span></span><div class="seg block" id="dpJenis">${Object.entries(JENIS_DP).map(([k, v]) => `<button type="button" data-act="dpm-jenis" data-v="${k}">${ic(v[1])}${v[0]}</button>`).join('')}</div></div>
      <div class="form-grid mt-3">
        <div class="field"><label for="dpSaldo">Saldo Awal <span class="req">*</span></label><div class="money-input"><span class="cur">Rp</span><input id="dpSaldo" inputmode="numeric" data-in="money" data-words="#dpWords" value="${edit ? fmtN(edit.saldo_awal) : ''}" placeholder="0"></div><div class="hint" id="dpWords">${edit && edit.saldo_awal ? terbilang(edit.saldo_awal) : 'Saldo saat dompet mulai dicatat'}</div></div>
        <div class="field"><label for="dpKet">Keterangan</label><input class="input" id="dpKet" maxlength="80" placeholder="Contoh: Rekening operasional" value="${esc(edit?.keterangan || '')}"><div class="hint">${ic('lock', 'sm')} Jangan simpan nomor rekening lengkap / PIN.</div></div>
      </div>
      <div class="field mt-3"><span class="lbl">Warna Penanda</span><div class="swatches" id="dpWarna">${WARNA.map(([h, n]) => `<button type="button" class="swatch" style="background:${h}" data-act="dpm-warna" data-v="${h}" title="${n}" aria-label="${n}"></button>`).join('')}</div></div>
      <div class="field mt-3"><span class="lbl">Ikon Dompet (opsional)</span>
        <div class="row"><span id="dpPrev"></span><label class="btn btn-outline btn-sm">${ic('upload')}Unggah Gambar<input type="file" accept="image/png,image/jpeg,image/webp" hidden data-ch="dpm-img"></label><button type="button" class="btn btn-ghost btn-sm" data-act="dpm-noimg">Hapus</button></div>
        <div class="hint">Diperkecil otomatis (maks. 128×128) lalu disimpan di Google Drive.</div></div>
      ${edit ? `<label class="check mt-3"><span class="switch"><input type="checkbox" id="dpAktif" ${edit.aktif ? 'checked' : ''}><span></span></span><span><b class="small">Dompet aktif</b><br><span class="hint">Nonaktifkan untuk mengarsipkan dompet yang tidak dipakai (riwayat tetap tersimpan).</span></span></label>` : ''}`,
    foot: `${edit ? `<button class="btn btn-danger-soft" data-act="dpm-del" ${used ? 'title="Dompet sudah dipakai transaksi — arsipkan saja"' : ''}>${ic(used ? 'archive' : 'delete')}${used ? 'Arsipkan' : 'Hapus'}</button>` : ''}<div class="grow"></div>
      <button class="btn btn-ghost" data-act="modal-close">Batal</button><button class="btn btn-primary" data-act="dpm-save">${ic('check')}Simpan Dompet</button>`,
    onOpen: dpmRender
  });
}
function dpmRender() {
  $$('#dpJenis button').forEach((b) => b.classList.toggle('on', b.dataset.v === DPM.jenis));
  $$('#dpWarna .swatch').forEach((b) => b.classList.toggle('on', b.dataset.v.toLowerCase() === String(DPM.warna).toLowerCase()));
  const nama = $('#dpNama').value || 'Dompet';
  $('#dpPrev').innerHTML = dpBadge({ nama, jenis: DPM.jenis, warna: DPM.warna, ikon_url: DPM.img || DPM.ikon_url });
}
ACT['dpm-jenis'] = (el) => { DPM.jenis = el.dataset.v; dpmRender(); };
ACT['dpm-warna'] = (el) => { DPM.warna = el.dataset.v; dpmRender(); };
ACT['dpm-noimg'] = () => { DPM.img = null; DPM.ikon_url = ''; dpmRender(); };
CHG['dpm-img'] = async (el) => {
  const f = el.files[0];
  if (!f) return;
  try { DPM.img = await resizeImage(f, 128); dpmRender(); } catch (e) { toast(e.message, 'error'); }
};
ACT['dpm-save'] = () => {
  const nama = $('#dpNama').value.trim();
  if (!nama) { toast('Nama dompet wajib diisi.', 'warn'); $('#dpNama').focus(); return; }
  if (S.dompet.some((d) => d.nama.toLowerCase() === nama.toLowerCase() && d.id !== DPM.edit?.id)) { toast('Nama dompet sudah dipakai.', 'warn'); return; }
  const d = {
    id: DPM.edit ? DPM.edit.id : uid('dp'), nama, jenis: DPM.jenis, saldo_awal: moneyVal('#dpSaldo'), ikon_url: DPM.ikon_url || '', warna: DPM.warna,
    aktif: DPM.edit ? $('#dpAktif').checked : true, keterangan: $('#dpKet').value.trim(), urutan: DPM.edit ? DPM.edit.urutan : S.dompet.length + 1
  };
  const img = DPM.img;
  DB.simpanDompet(d);
  closeModal();
  toast(DPM.edit ? 'Dompet diperbarui.' : `Dompet ${nama} ditambahkan.`, 'success');
  if (img) {
    api('unggahGambar', img, 'dompet').then((res) => { DB.simpanDompet({ ...S.dMap.get(d.id), ikon_url: res.data.url }); toast('Ikon dompet tersimpan di Drive.', 'success'); })
      .catch((e) => toast('Ikon gagal diunggah: ' + e.message, 'error'));
  }
};
ACT['dpm-del'] = async () => {
  const d = DPM.edit;
  const used = S.tx.filter((t) => t.dompet_id === d.id || t.dompet_tujuan_id === d.id).length || (d.saldo !== d.saldo_awal);
  if (used) {
    if (await confirmBox({ title: 'Arsipkan Dompet?', text: `<b>${esc(d.nama)}</b> sudah dipakai transaksi sehingga tidak bisa dihapus. Arsipkan agar tidak tampil di daftar aktif? Riwayatnya tetap aman.`, ok: 'Arsipkan', icon: 'archive' })) {
      DB.simpanDompet({ ...S.dMap.get(d.id), aktif: false }); toast('Dompet diarsipkan.', 'success');
    }
    return;
  }
  if (await confirmBox({ title: 'Hapus Dompet?', text: `Dompet <b>${esc(d.nama)}</b> akan dihapus permanen.`, ok: 'Hapus', danger: true })) { DB.hapusDompet(d.id); toast('Dompet dihapus.', 'success'); }
};
ACT['dp-new'] = () => openDompetModal(null);
ACT['dp-edit'] = (el) => openDompetModal(S.dMap.get(el.dataset.id));

/** Gambar diperkecil di browser sebelum dikirim (hemat kuota & cepat). */
function resizeImage(file, max = 256) {
  return new Promise((ok, fail) => {
    if (!/^image\/(png|jpeg|webp|gif)$/.test(file.type)) { fail(new Error('File harus berupa gambar PNG/JPG/WEBP.')); return; }
    if (file.size > 10 * 1024 * 1024) { fail(new Error('Ukuran gambar maksimal 10 MB.')); return; }
    const url = URL.createObjectURL(file), img = new Image();
    img.onload = () => {
      const s = Math.min(1, max / Math.max(img.width, img.height));
      const w = Math.max(1, Math.round(img.width * s)), h = Math.max(1, Math.round(img.height * s));
      const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
      const ctx = cv.getContext('2d'); ctx.imageSmoothingQuality = 'high'; ctx.drawImage(img, 0, 0, w, h);
      URL.revokeObjectURL(url);
      let out = cv.toDataURL('image/png');
      if (out.length > 300000) out = cv.toDataURL('image/jpeg', 0.86);
      ok(out);
    };
    img.onerror = () => { URL.revokeObjectURL(url); fail(new Error('Gambar tidak dapat dibaca.')); };
    img.src = url;
  });
}

// ── Mutasi, Urutan, Arsip, Rekonsiliasi ──
ACT['dp-mutasi'] = (el) => {
  const d = S.dMap.get(el.dataset.id);
  if (!d) return;
  const list = S.tx.filter((t) => t.dompet_id === d.id || t.dompet_tujuan_id === d.id);
  let bal = d.saldo;
  const rows = list.slice(0, 150).map((t) => { const e = txEffect(t, d.id); const after = bal; bal -= e; return { t, e, after }; });
  const T = today(), d30 = addDays(T, -29);
  const in30 = sum(list.filter((t) => t.tanggal >= d30), (t) => Math.max(0, txEffect(t, d.id)));
  const out30 = sum(list.filter((t) => t.tanggal >= d30), (t) => Math.max(0, -txEffect(t, d.id)));
  openModal({
    title: 'Mutasi ' + d.nama, icon: 'receipt_long', size: 'lg',
    body: `<div class="grid g-3" style="gap:12px">
        <div class="stat-box"><div class="k">${ic('account_balance_wallet')}Saldo Saat Ini</div><div class="v money">${rp(d.saldo)}</div></div>
        <div class="stat-box"><div class="k">${ic('south_west')}Masuk 30 Hari</div><div class="v pos money">+${rp(in30)}</div></div>
        <div class="stat-box"><div class="k">${ic('north_east')}Keluar 30 Hari</div><div class="v neg money">-${rp(out30)}</div></div></div>
      <div class="list mt-3">${rows.length ? rows.map(({ t, e, after }) => `<div class="li click" data-act="tx-edit" data-id="${t.id}">${txWell(t, 'sm')}<div class="grow"><div class="semi small ellipsis">${esc(txTitle(t))}</div><div class="xs faint">${tgl(t.tanggal)} • ${esc(userOf(t.dibuat_oleh).nama)}</div></div>
        <div class="right"><div class="amt ${e >= 0 ? 'in' : 'out'} money" style="font-size:13.5px">${e >= 0 ? '+' : '-'}${rp(Math.abs(e))}</div><div class="xs faint money">Saldo ${rp(after)}</div></div></div>`).join('') : emptyBox('receipt_long', 'Belum ada mutasi', 'Transaksi untuk dompet ini akan muncul di sini.')}</div>
      ${list.length > 150 ? `<div class="hint mt-2">Menampilkan 150 mutasi terbaru dari ${list.length}. Lihat semua di halaman Transaksi.</div>` : ''}`,
    foot: `<button class="btn btn-ghost" data-act="modal-close">Tutup</button><div class="grow"></div><button class="btn btn-outline" data-act="dp-edit" data-id="${d.id}">${ic('edit')}Ubah Dompet</button><button class="btn btn-primary" data-act="tx-new" data-dompet="${d.id}">${ic('add')}Catat di Dompet Ini</button>`
  });
};
ACT['dp-urutan'] = () => {
  let ids = dompetAktif().map((d) => d.id);
  const draw = () => {
    $('#urList').innerHTML = ids.map((id, i) => { const d = S.dMap.get(id); return `<div class="li">${dpBadge(d, true)}<span class="grow semi">${esc(d.nama)}</span><button class="btn-icon" data-act="ur-mv" data-i="${i}" data-d="-1" ${i ? '' : 'disabled'} aria-label="Naik">${ic('arrow_upward')}</button><button class="btn-icon" data-act="ur-mv" data-i="${i}" data-d="1" ${i < ids.length - 1 ? '' : 'disabled'} aria-label="Turun">${ic('arrow_downward')}</button></div>`; }).join('');
  };
  ACT['ur-mv'] = (el) => { const i = +el.dataset.i, j = i + +el.dataset.d; [ids[i], ids[j]] = [ids[j], ids[i]]; draw(); };
  openModal({
    title: 'Atur Urutan Dompet', icon: 'drag_indicator', size: 'sm', body: '<div class="list" id="urList"></div>',
    foot: '<div class="grow"></div><button class="btn btn-ghost" data-act="modal-close">Batal</button><button class="btn btn-primary" id="urSave">Simpan Urutan</button>',
    onOpen: (m) => { draw(); m.querySelector('#urSave').onclick = () => { DB.urutanDompet(ids); closeModal(); toast('Urutan dompet disimpan.', 'success'); }; }
  });
};
ACT['dp-arsip'] = () => {
  const list = S.dompet.filter((d) => !d.aktif);
  openModal({
    title: 'Arsip Rekening Non-Aktif', icon: 'archive', size: 'sm',
    body: list.length ? `<div class="list">${list.map((d) => `<div class="li">${dpBadge(d, true)}<div class="grow"><div class="semi small">${esc(d.nama)}</div><div class="xs faint money">Saldo ${rp(d.saldo)}</div></div><button class="btn btn-soft btn-sm" data-act="dp-aktifkan" data-id="${d.id}">${ic('unarchive')}Aktifkan</button></div>`).join('')}</div>` : emptyBox('archive', 'Tidak ada dompet diarsipkan'),
    foot: '<div class="grow"></div><button class="btn btn-ghost" data-act="modal-close">Tutup</button>'
  });
};
ACT['dp-aktifkan'] = (el) => { DB.simpanDompet({ ...S.dMap.get(el.dataset.id), aktif: true }); toast('Dompet diaktifkan kembali.', 'success'); ACT['dp-arsip'](); };
ACT['dp-rekon'] = () => {
  const list = dompetAktif();
  if (!list.length) return;
  openModal({
    title: 'Rekonsiliasi Saldo', icon: 'balance',
    body: `<p class="small muted mb-3">Cocokkan saldo di aplikasi dengan saldo nyata (m-banking / dompet fisik). Selisihnya dicatat sebagai transaksi penyesuaian.</p>
      <div class="field"><label for="rkDp">Dompet</label><select class="select" id="rkDp" data-ch="rk">${dompetOpts(list[0].id, false)}</select></div>
      <div class="grid g-2 mt-3" style="gap:12px"><div class="stat-box"><div class="k">${ic('smartphone')}Saldo di Aplikasi</div><div class="v money" id="rkApp"></div></div><div class="stat-box"><div class="k">${ic('difference')}Selisih</div><div class="v money" id="rkDiff">Rp 0</div></div></div>
      <div class="field mt-3"><label for="rkReal">Saldo Sebenarnya <span class="req">*</span></label><div class="money-input"><span class="cur">Rp</span><input id="rkReal" inputmode="numeric" data-in="money" data-after="rk" placeholder="0"></div></div>`,
    foot: '<div class="grow"></div><button class="btn btn-ghost" data-act="modal-close">Batal</button><button class="btn btn-primary" data-act="rk-save">Simpan Penyesuaian</button>',
    onOpen: () => ON.rk()
  });
};
ON.rk = () => {
  const d = S.dMap.get($('#rkDp').value);
  $('#rkApp').textContent = rp(d.saldo);
  const real = $('#rkReal').value ? moneyVal('#rkReal') : d.saldo;
  const diff = real - d.saldo;
  $('#rkDiff').innerHTML = `<span class="${diff > 0 ? 'pos' : diff < 0 ? 'neg' : ''}">${diff > 0 ? '+' : ''}${rp(diff)}</span>`;
};
CHG.rk = () => ON.rk();
ACT['rk-save'] = () => {
  const d = S.dMap.get($('#rkDp').value);
  if (!$('#rkReal').value) { toast('Isi saldo sebenarnya.', 'warn'); return; }
  const diff = moneyVal('#rkReal') - d.saldo;
  if (!diff) { toast('Saldo sudah cocok, tidak ada penyesuaian.', 'info'); closeModal(); return; }
  const tipe = diff > 0 ? 'masuk' : 'keluar';
  const kt = tipe === 'masuk' ? 'pemasukan' : 'pengeluaran';
  const kat = S.kMap.has(tipe === 'masuk' ? 'kat_lain_m' : 'kat_lain_k') ? (tipe === 'masuk' ? 'kat_lain_m' : 'kat_lain_k') : katByTipe(kt).slice(-1)[0]?.id;
  if (!kat) { toast('Buat kategori ' + kt + ' terlebih dahulu.', 'warn'); return; }
  DB.simpanTx({ id: uid('tx'), tanggal: today(), tipe, dompet_id: d.id, kategori_id: kat, nominal: Math.abs(diff), catatan: 'Penyesuaian saldo (rekonsiliasi)', dibuat_oleh: S.me.email, ref: 'rekonsiliasi' });
  closeModal(); toast(`Saldo ${d.nama} disesuaikan ${diff > 0 ? '+' : ''}${rp(diff)}.`, 'success');
};

// ════════════════════════════════════════════════════════════
// 13. HALAMAN BUDGET & KATEGORI
// ════════════════════════════════════════════════════════════
const BUD = { mk: '', tipe: 'pengeluaran', page: 1 };
PAGES.budget = { t: 'Budget & Kategori', r: renderBudget, prep: (o) => { if (o.tab === 'kategori') setTimeout(() => $('#katSection')?.scrollIntoView({ behavior: 'smooth' }), 60); } };
function catAvg(id, months = 3) {
  const mk = mKey(today());
  const vals = Array.from({ length: months }, (_, i) => monthAgg(mAdd(mk, -i - 1)).cat[id] || 0);
  return sum(vals, (v) => v) / months;
}
function renderBudget() {
  const T = today(), cur = mKey(T);
  const mk = BUD.mk || cur;
  BUD.mk = mk;
  const bud = budgetMonth(mk);
  const a = monthAgg(mk);
  const lim = sum(bud, (x) => x.limit), used = sum(bud, (x) => x.used), p = lim ? used / lim : 0;
  const prevLim = sum(budgetMonth(mAdd(mk, -1)), (x) => x.limit);
  const daysLeft = mk === cur ? mDays(mk) - +T.slice(8) + 1 : mk > cur ? mDays(mk) : 0;
  const daily = daysLeft ? Math.max(0, lim - used) / daysLeft : 0;
  const dayN = mk === cur ? +T.slice(8) : mDays(mk);
  const proj = mk === cur && dayN ? (used / dayN) * mDays(mk) : used;
  const over = bud.filter((x) => x.p >= 1);
  const budCats = new Set(bud.map((x) => x.k.id));
  const luar = Object.entries(a.cat).filter(([id, v]) => S.kMap.get(id)?.tipe === 'pengeluaran' && !budCats.has(id) && v > 0).map(([id, v]) => ({ k: S.kMap.get(id), v })).sort((x, y) => y.v - x.v);
  const kats = katByTipe(BUD.tipe).slice().sort((x, y) => x.nama.localeCompare(y.nama));
  const per = 8, pages = Math.max(1, Math.ceil(kats.length / per));
  BUD.page = clamp(BUD.page, 1, pages);
  const kSlice = kats.slice((BUD.page - 1) * per, BUD.page * per);
  const status = (x) => x.p >= 1 ? ['Melebihi Batas', 'danger'] : x.p >= 0.8 ? ['Hampir Habis', 'warn'] : ['Aman', 'ok'];

  $('#page-budget').innerHTML = `
  <div class="page-head"><div class="grow"><div class="crumb"><span>Rencana Keuangan</span>${ic('chevron_right', 'sm')}<span class="semi" style="color:var(--text)">Pagu Anggaran Bulanan</span></div>
    <h1 class="page-title">Anggaran Bulanan Keluarga</h1></div>
    <div class="head-actions"><div class="month-nav"><button data-act="bud-mon" data-d="-1" aria-label="Bulan sebelumnya">${ic('chevron_left')}</button><span>${ic('calendar_today', 'sm')}${mLabel(mk)}</span><button data-act="bud-mon" data-d="1" aria-label="Bulan berikutnya">${ic('chevron_right')}</button></div>
      <button class="btn btn-primary" data-act="budget-edit">${ic('tune')}Atur Budget</button></div></div>

  <div class="bud-kpi">
    <div class="card" style="min-width:0">
      <div class="row" style="justify-content:space-between"><div class="card-t">Status Keseluruhan Pagu</div><span class="pill pill-${levelOf(p)}"><span class="dot"></span>${p >= 1 ? 'Melebihi Pagu' : p >= 0.8 ? 'Waspada' : 'Terkendali Baik'}</span></div>
      <div class="row mt-3" style="gap:22px;flex-wrap:wrap"><div class="ring lg">${ring(p, 128, 12)}<span class="rt">${pct(p)}<small>Terpakai</small></span></div>
        <div class="stack" style="gap:10px"><div><div class="xs faint">Terpakai Saat Ini</div><div class="money" style="font:800 22px/1.2 var(--font-head)">${rp(used)}</div></div><div><div class="xs faint">Dari Total Pagu</div><div class="semi money">${rp(lim)}</div></div>
        ${a.out - used > 0 ? `<div class="xs faint">+ ${rp(a.out - used)} di luar pagu</div>` : ''}</div></div>
      <div class="row mt-3" style="justify-content:space-between;border-top:1px solid var(--divider);padding-top:12px;flex-wrap:wrap;gap:8px"><span class="small muted">${ic('timer', 'sm')} Tersisa <b>${daysLeft} hari</b> bulan ini</span><span class="small right"><span class="faint xs">Alokasi Harian Aman</span><br><b class="money">${rp(daily)} / hari</b></span></div>
    </div>
    <div class="card kpi"><div class="kpi-top"><span class="kpi-l">Total Pagu Bulanan</span><span class="icon-well sm">${ic('account_balance')}</span></div><div class="kpi-v money">${rpS(lim)}</div><div class="kpi-f">Ditetapkan untuk ${bud.length} pos</div>
      <div class="kpi-f" style="border-top:1px solid var(--divider);padding-top:8px">${ic('check_circle', 'sm')}${prevLim ? (Math.abs(lim - prevLim) < 1 ? 'Alokasi stabil vs ' + BULAN[+mAdd(mk, -1).slice(5) - 1] : (lim > prevLim ? 'Naik ' : 'Turun ') + rpS(Math.abs(lim - prevLim)) + ' vs bln lalu') : 'Bulan lalu belum ada pagu'}</div></div>
    <div class="card kpi"><div class="kpi-top"><span class="kpi-l">Total Pengeluaran</span><span class="icon-well sm" style="${tint('#F5B13D')}">${ic('shopping_bag')}</span></div><div class="kpi-v money">${rpS(a.out)}</div><div class="kpi-f"><span class="pill pill-${levelOf(lim ? a.out / lim : 0)}">${lim ? pct(a.out / lim, 1) : '-'}</span>dari pagu</div>
      <div class="kpi-f" style="border-top:1px solid var(--divider);padding-top:8px">${ic(over.length ? 'error' : 'check_circle', 'sm')}${over.length ? over.length + ' kategori melebihi batas' : 'Semua kategori dalam batas'}</div></div>
    <div class="card kpi"><div class="kpi-top"><span class="kpi-l">Sisa Kuota Aman</span><span class="icon-well sm">${ic('savings')}</span></div><div class="kpi-v money" style="color:var(--primary-text)">${rpS(Math.max(0, lim - used))}</div><div class="kpi-f">Siap disimpan ke dana darurat</div>
      <div class="kpi-f" style="border-top:1px solid var(--divider);padding-top:8px">${ic(lim - proj >= 0 ? 'trending_up' : 'trending_down', 'sm')}${lim ? (lim - proj >= 0 ? 'Proyeksi surplus ' : 'Proyeksi defisit ') + rpS(Math.abs(lim - proj)) : '—'}</div></div>
  </div>

  <div class="card pad-0 mt-3">
    <div class="row" style="padding:20px 22px;border-bottom:1px solid var(--border);flex-wrap:wrap"><div class="grow"><div class="card-t">Rincian Alokasi Per Kategori</div><div class="card-s">Pantau pemakaian harian agar tidak melampaui toleransi pagu bulanan</div></div>
      <div class="legend-inline"><span class="faint">Legenda:</span><span><i style="background:var(--accent)"></i>&lt;80% Aman</span><span><i style="background:var(--gold)"></i>80–99% Siaga</span><span><i style="background:var(--danger)"></i>≥100% Melebihi</span></div></div>
    ${bud.length ? bud.map((x) => {
      const [sl, sc] = status(x);
      const harian = daysLeft ? Math.max(0, x.sisa) / daysLeft : 0;
      return `<div class="bud-cat ${x.p >= 1 ? 'over' : ''}">
        <div class="row">${katWell(x.k)}<div class="grow"><div class="semi" style="font-family:var(--font-head);${x.p >= 1 ? 'color:var(--danger-700)' : ''}">${esc(x.k.nama)}</div><div class="xs faint">Pagu: <b class="money" style="color:var(--text)">${rp(x.limit)}</b> • ${x.n} transaksi</div></div></div>
        <div><div class="bar-top"><span>Terpakai: <b class="money" style="color:var(--text)">${rp(x.used)}</b></span><b style="color:${x.p >= 1 ? 'var(--danger-700)' : x.p >= 0.8 ? 'var(--gold-700)' : 'var(--primary-text)'}">${pct(x.p, 1)}</b></div>${progress(x.p)}
          <div class="bar-bot"><span>${x.sisa >= 0 ? `Sisa kuota: <b class="money" style="color:var(--text)">${rp(x.sisa)}</b>` : `<span class="neg">Defisit: <b class="money">${rp(x.sisa)}</b> (melampaui pagu)</span>`}</span><span class="${x.p >= 1 ? 'neg' : ''}">${x.p >= 1 ? 'Disarankan pause pos ini' : x.p >= 0.8 ? `Batas wajar harian ${rpS(harian)}` : 'Pengeluaran terencana aman'}</span></div></div>
        <div class="bc-status"><span class="pill pill-${sc}"><span class="dot"></span>${sl} (${pct(x.p)})</span></div>
        <div class="row-actions bc-act"><button class="btn-icon" data-go="transaksi" data-kat="${x.k.id}" data-month="${mk}" aria-label="Riwayat transaksi" title="Riwayat transaksi">${ic('history')}</button><button class="btn-icon" data-act="budget-edit" data-kat="${x.k.id}" aria-label="Ubah pagu" title="Ubah pagu">${ic('edit')}</button></div>
      </div>`;
    }).join('') : `<div style="padding:10px 22px 22px">${emptyBox('donut_large', 'Pagu ' + mLabel(mk) + ' belum diatur', 'Tetapkan batas belanja per kategori; peringatan otomatis muncul saat mencapai 80% dan 100%.',
      `<div class="row" style="justify-content:center;flex-wrap:wrap"><button class="btn btn-primary btn-sm" data-act="budget-edit">${ic('tune')}Atur Pagu</button>${budgetMonth(mAdd(mk, -1)).length ? `<button class="btn btn-outline btn-sm" data-act="budget-copy">${ic('content_copy')}Salin dari ${mLabel(mAdd(mk, -1))}</button>` : ''}</div>`)}</div>`}
    ${luar.length ? `<div style="padding:14px 22px;border-top:1px solid var(--border);background:var(--surface-2);border-radius:0 0 var(--r-lg) var(--r-lg)"><div class="small"><b>Pengeluaran di luar pagu:</b> ${luar.slice(0, 4).map((x) => `${esc(x.k.nama)} <span class="money">${rpS(x.v)}</span>`).join(' • ')} <button class="link" data-act="budget-edit">Tambahkan ke pagu ${ic('arrow_forward')}</button></div></div>` : ''}
  </div>

  <div id="katSection" class="section-h" style="margin-top:42px"><div class="grow"><div class="crumb"><span>Master Data Keuangan</span>${ic('chevron_right', 'sm')}<span class="semi" style="color:var(--text)">Manajemen Kategori</span></div><h2 style="font-size:24px">Katalog Kategori Transaksi</h2><div class="small muted">Kelola pos pencatatan bersama untuk suami dan istri dengan label visual terstandarisasi</div></div>
    <div class="seg"><button class="${BUD.tipe === 'pengeluaran' ? 'on' : ''}" data-act="kat-tab" data-v="pengeluaran">Pengeluaran (${katByTipe('pengeluaran').length})</button><button class="${BUD.tipe === 'pemasukan' ? 'on' : ''}" data-act="kat-tab" data-v="pemasukan">Pemasukan (${katByTipe('pemasukan').length})</button></div>
    <button class="btn btn-primary" data-act="kat-new" data-tipe="${BUD.tipe}">${ic('add')}Tambah Kategori</button></div>
  <div class="card">
    <div class="tbl-wrap"><table class="tbl"><thead><tr><th>Kategori &amp; Ikon</th><th>Tipe Arus</th><th>Penanda Warna</th><th class="c">Frekuensi Bulan Ini</th><th class="r">Rata-rata 3 Bulan</th><th class="r">Aksi</th></tr></thead><tbody>
      ${kSlice.map((k) => `<tr><td><div class="row">${katWell(k, 'sm')}<div><div class="t-title">${esc(k.nama)}</div><div class="t-sub">${esc(k.keterangan || '')}</div></div></div></td>
        <td><span class="pill ${k.tipe === 'pengeluaran' ? 'pill-danger' : 'pill-ok'}">${k.tipe === 'pengeluaran' ? 'Pengeluaran' : 'Pemasukan'}</span></td>
        <td><span class="row" style="gap:8px"><span style="width:12px;height:12px;border-radius:50%;background:${k.warna};display:inline-block"></span><span class="small">${esc(warnaNama(k.warna))}</span></span></td>
        <td class="c small">${monthAgg(mKey(today())).catN[k.id] || 0} transaksi</td>
        <td class="r small semi money">${rp(catAvg(k.id))} / bln</td>
        <td class="r"><div class="row-actions"><button class="btn-icon" data-act="kat-edit" data-id="${k.id}" aria-label="Ubah">${ic('edit')}</button><button class="btn-icon danger" data-act="kat-del" data-id="${k.id}" aria-label="Hapus">${ic('delete')}</button></div></td></tr>`).join('') || `<tr><td colspan="6">${emptyBox('category', 'Belum ada kategori')}</td></tr>`}
    </tbody></table></div>
    ${pager(kats.length, BUD.page, per, 'kat-page')}
  </div>`;
}
ACT['bud-mon'] = (el) => { BUD.mk = mAdd(BUD.mk, +el.dataset.d); renderBudget(); };
ACT['kat-tab'] = (el) => { BUD.tipe = el.dataset.v; BUD.page = 1; renderBudget(); };
ACT['kat-page'] = (el) => { BUD.page = +el.dataset.p; renderBudget(); };
ACT['budget-copy'] = () => {
  const mk = BUD.mk || mKey(today()), src = budgetMonth(mAdd(mk, -1));
  DB.simpanBudget(src.map((x) => ({ id: 'bg_' + mk.replace('-', '') + '_' + x.k.id, bulan: mk, kategori_id: x.k.id, batas_nominal: x.limit })));
  toast(`${src.length} pagu disalin ke ${mLabel(mk)}.`, 'success');
};
ACT['budget-edit'] = (el) => {
  const mk = CUR === 'budget' && BUD.mk ? BUD.mk : mKey(today());
  const cats = katByTipe('pengeluaran');
  const cur = new Map(S.budget.filter((b) => b.bulan === mk).map((b) => [b.kategori_id, b]));
  const prev = new Map(S.budget.filter((b) => b.bulan === mAdd(mk, -1)).map((b) => [b.kategori_id, b]));
  openModal({
    title: 'Atur Pagu — ' + mLabel(mk), icon: 'tune', size: 'lg',
    body: `<div class="row mb-3" style="flex-wrap:wrap"><div class="grow small muted">Isi batas belanja untuk kategori yang ingin dipantau. Kosongkan untuk menghapus pagu.</div>${prev.size ? `<button class="btn btn-outline btn-sm" data-act="bm-copy">${ic('content_copy')}Salin ${mLabel(mAdd(mk, -1), true)}</button>` : ''}</div>
      <div class="list" id="bmList">${cats.map((k) => `<div class="li" data-kat="${k.id}">${katWell(k, 'sm')}<div class="grow"><div class="semi small">${esc(k.nama)}</div><div class="xs faint">Rata-rata 3 bln: <span class="money">${rpS(catAvg(k.id))}</span>${prev.get(k.id) ? ` • bln lalu <span class="money">${rpS(prev.get(k.id).batas_nominal)}</span>` : ''}</div></div>
        <div class="money-input" style="width:170px"><span class="cur">Rp</span><input inputmode="numeric" data-in="money" data-after="bm-total" data-kat="${k.id}" placeholder="0" value="${cur.get(k.id) ? fmtN(cur.get(k.id).batas_nominal) : ''}" aria-label="Pagu ${esc(k.nama)}"></div></div>`).join('')}</div>`,
    foot: `<div class="grow small">Total pagu: <b class="money" id="bmTotal"></b></div><button class="btn btn-ghost" data-act="modal-close">Batal</button><button class="btn btn-primary" data-act="bm-save" data-mk="${mk}">${ic('check')}Simpan Pagu</button>`,
    onOpen: (m) => {
      ON['bm-total']();
      ACT['bm-copy'] = () => { $$('#bmList input').forEach((i) => { const p = prev.get(i.dataset.kat); if (p && !i.value) { i.value = fmtN(p.batas_nominal); } }); ON['bm-total'](); };
      if (el?.dataset?.kat) { const i = m.querySelector(`input[data-kat="${el.dataset.kat}"]`); if (i) setTimeout(() => { i.scrollIntoView({ block: 'center' }); i.focus(); }, 60); }
    }
  });
};
ON['bm-total'] = () => { const t = sum($$('#bmList input'), (i) => parseNum(i.value)); const el = $('#bmTotal'); if (el) el.textContent = rp(t); };
ACT['bm-save'] = (el) => {
  const mk = el.dataset.mk;
  const ada = new Map(S.budget.filter((b) => b.bulan === mk).map((b) => [b.kategori_id, b]));
  const simpan = [], hapus = [];
  $$('#bmList input').forEach((i) => {
    const v = parseNum(i.value), k = i.dataset.kat, b = ada.get(k);
    if (v > 0 && (!b || b.batas_nominal !== v)) simpan.push({ id: b ? b.id : 'bg_' + mk.replace('-', '') + '_' + k, bulan: mk, kategori_id: k, batas_nominal: v });
    if (!v && b) hapus.push(b.id);
  });
  if (simpan.length || hapus.length) DB.simpanBudget(simpan, hapus);
  closeModal();
  toast(simpan.length || hapus.length ? `Pagu ${mLabel(mk)} disimpan.` : 'Tidak ada perubahan.', simpan.length || hapus.length ? 'success' : 'info');
};

// ── Modal Kategori ──
let KM = null;
function openKatModal(k, tipe, onDone) {
  const edit = k ? { ...k } : null;
  const used = edit ? S.tx.some((t) => t.kategori_id === edit.id) : false;
  KM = { edit, tipe: edit ? edit.tipe : tipe || 'pengeluaran', ikon: edit ? edit.ikon : 'category', warna: edit ? edit.warna : WARNA[9][0], onDone, used };
  openModal({
    title: edit ? 'Ubah Kategori' : 'Tambah Kategori', icon: 'category',
    body: `<div class="field"><label for="kmNama">Nama Kategori <span class="req">*</span></label><input class="input" id="kmNama" maxlength="50" value="${esc(edit?.nama || '')}" placeholder="Contoh: Belanja Dapur"></div>
      <div class="field mt-3"><span class="lbl">Tipe Arus</span><div class="seg block" id="kmTipe"><button type="button" data-act="km-tipe" data-v="pengeluaran" ${used ? 'disabled' : ''}>${ic('arrow_upward')}Pengeluaran</button><button type="button" data-act="km-tipe" data-v="pemasukan" ${used ? 'disabled' : ''}>${ic('arrow_downward')}Pemasukan</button></div>${used ? '<div class="hint">Tipe tidak bisa diubah karena kategori sudah dipakai transaksi.</div>' : ''}</div>
      <div class="field mt-3"><label for="kmKet">Keterangan</label><input class="input" id="kmKet" maxlength="80" value="${esc(edit?.keterangan || '')}" placeholder="Contoh: Kebutuhan pokok harian"></div>
      <div class="field mt-3"><span class="lbl">Ikon</span><div class="icon-pick" id="kmIkon">${IKON_KAT.map((n) => `<button type="button" data-act="km-ikon" data-v="${n}" aria-label="${n}">${ic(n)}</button>`).join('')}</div></div>
      <div class="field mt-3"><span class="lbl">Penanda Warna</span><div class="swatches" id="kmWarna">${WARNA.map(([h, n]) => `<button type="button" class="swatch" style="background:${h}" data-act="km-warna" data-v="${h}" title="${n}" aria-label="${n}"></button>`).join('')}</div></div>
      <div class="row mt-3"><span class="small faint">Pratinjau:</span><span id="kmPrev"></span></div>`,
    foot: `${edit ? `<button class="btn btn-danger-soft" data-act="kat-del" data-id="${edit.id}">${ic('delete')}Hapus</button>` : ''}<div class="grow"></div><button class="btn btn-ghost" data-act="modal-close">Batal</button><button class="btn btn-primary" data-act="km-save">${ic('check')}Simpan Kategori</button>`,
    onOpen: kmRender
  });
}
function kmRender() {
  $$('#kmTipe button').forEach((b) => b.classList.toggle('on', b.dataset.v === KM.tipe));
  $$('#kmIkon button').forEach((b) => b.classList.toggle('on', b.dataset.v === KM.ikon));
  $$('#kmWarna .swatch').forEach((b) => b.classList.toggle('on', b.dataset.v.toLowerCase() === String(KM.warna).toLowerCase()));
  $('#kmPrev').innerHTML = `<span class="row" style="gap:8px">${katWell({ ikon: KM.ikon, warna: KM.warna }, 'sm')}${katPill({ nama: $('#kmNama').value || 'Kategori', warna: KM.warna })}</span>`;
}
ACT['km-tipe'] = (el) => { KM.tipe = el.dataset.v; kmRender(); };
ACT['km-ikon'] = (el) => { KM.ikon = el.dataset.v; kmRender(); };
ACT['km-warna'] = (el) => { KM.warna = el.dataset.v; kmRender(); };
ACT['km-save'] = () => {
  const nama = $('#kmNama').value.trim();
  if (!nama) { toast('Nama kategori wajib diisi.', 'warn'); $('#kmNama').focus(); return; }
  if (S.kategori.some((k) => k.nama.toLowerCase() === nama.toLowerCase() && k.tipe === KM.tipe && k.id !== KM.edit?.id)) { toast('Kategori dengan nama ini sudah ada.', 'warn'); return; }
  const k = { id: KM.edit ? KM.edit.id : uid('kat'), nama, tipe: KM.tipe, ikon: KM.ikon, warna: KM.warna, keterangan: $('#kmKet').value.trim() };
  DB.simpanKategori(k);
  const done = KM.onDone;
  closeModal();
  toast(KM.edit ? 'Kategori diperbarui.' : `Kategori ${nama} ditambahkan.`, 'success');
  if (done) done(k);
};
ACT['kat-new'] = (el) => {
  if (el.dataset.from === 'txm' && TXM) {
    const snap = { amt: $('#txmAmt').value, note: $('#txmNote').value, tgl: $('#txmTgl').value, dp: $('#txmDp').value, tipe: TXM.tipe, oleh: TXM.oleh, edit: TXM.edit };
    openKatModal(null, el.dataset.tipe, (k) => {
      if (snap.edit) { openTxModal({ tx: { ...snap.edit, kategori_id: k.id } }); }
      else openTxModal({ tipe: snap.tipe, dompet: snap.dp, tanggal: snap.tgl, catatan: snap.note });
      TXM.kat = k.id; TXM.oleh = snap.oleh; txmRender();
      if (snap.amt) { $('#txmAmt').value = snap.amt; ON.money($('#txmAmt')); }
    });
    return;
  }
  openKatModal(null, el.dataset.tipe || BUD.tipe);
};
ACT['kat-edit'] = (el) => openKatModal(S.kMap.get(el.dataset.id));
ACT['kat-del'] = async (el) => {
  const k = S.kMap.get(el.dataset.id);
  if (!k) return;
  const n = S.tx.filter((t) => t.kategori_id === k.id).length + S.budget.filter((b) => b.kategori_id === k.id).length + S.tagihan.filter((g) => g.kategori_id === k.id).length;
  if (n || (S.meta.totalTx > S.tx.length && k.id.startsWith('kat_'))) {
    toast(n ? `Kategori ${k.nama} masih dipakai ${n} data (transaksi/budget/tagihan). Ubah datanya dulu.` : 'Kategori bawaan mungkin dipakai transaksi lama, tidak bisa dihapus.', 'warn', 5000);
    return;
  }
  if (await confirmBox({ title: 'Hapus Kategori?', text: `Kategori <b>${esc(k.nama)}</b> akan dihapus permanen.`, ok: 'Hapus', danger: true })) { DB.hapusKategori(k.id); toast('Kategori dihapus.', 'success'); }
};


// ════════════════════════════════════════════════════════════
// 14. HALAMAN TARGET TABUNGAN
// ════════════════════════════════════════════════════════════
PAGES.tabungan = { t: 'Tabungan', r: renderTabungan };
const bulanAntara = (a, b) => { const [y1, m1] = a.split('-').map(Number), [y2, m2] = b.split('-').map(Number); return (y2 - y1) * 12 + (m2 - m1); };
function targetStatus(t) {
  const p = targetPct(t);
  if (t.status === 'dicairkan') return ['Sudah Dicairkan', 'muted', 'task_alt'];
  if (t.status === 'tercapai' || p >= 1) return ['Tercapai 100%', 'warn', 'emoji_events'];
  if (p >= 0.8) return ['Hampir Tercapai', 'warn', ''];
  if (t.tenggat && diffDays(today(), t.tenggat) < 0) return ['Lewat Tenggat', 'danger', ''];
  return [p >= 0.5 ? 'On Track' : 'Berjalan', 'ok', ''];
}
function surplusRata() {
  const mk = mKey(today());
  const v = [1, 2, 3].map((i) => { const a = monthAgg(mAdd(mk, -i)); return a.in - a.out; }).filter((x, i, arr) => arr.length);
  return sum(v, (x) => x) / 3;
}
function renderTabungan() {
  const T = today();
  const list = S.target.slice().sort((a, b) => ({ aktif: 0, tercapai: 1, dicairkan: 2 }[a.status] ?? 0) - ({ aktif: 0, tercapai: 1, dicairkan: 2 }[b.status] ?? 0) || String(a.tenggat || '9').localeCompare(String(b.tenggat || '9')));
  const tampil = list.filter((t) => t.status !== 'dicairkan');
  const cair = list.filter((t) => t.status === 'dicairkan');
  const aktif = tampil.filter((t) => t.status === 'aktif');
  const totT = sum(tampil, (t) => t.target_nominal), totK = sum(tampil, (t) => Math.min(t.terkumpul, t.target_nominal));
  const sisa = sum(aktif, (t) => Math.max(0, t.target_nominal - t.terkumpul));
  const butuh = sum(aktif.filter((t) => t.tenggat), (t) => Math.max(0, t.target_nominal - t.terkumpul) / Math.max(1, bulanAntara(T, t.tenggat)));
  const sur = surplusRata();
  const nBulan = sur > 0 && sisa > 0 ? Math.ceil(sisa / sur) : 0;
  const onTrack = sisa === 0 || (sur > 0 && sur >= butuh);

  $('#page-tabungan').innerHTML = `
  <div class="page-head"><div class="grow"><div class="crumb"><span class="tag">${ic('savings')}Finansial Masa Depan</span><span>• ${aktif.length} pos aktif</span></div><h1 class="page-title">Target Tabungan Keluarga</h1><p class="page-sub">Pantau progres tabungan bersama &amp; pos dana masa depan keluarga.</p></div>
    <div class="card flat row" style="padding:14px 18px;gap:12px"><span class="icon-well sm">${ic('data_usage')}</span><div><div class="xs faint">Total Terkumpul</div><div class="semi money"><b style="font:800 18px var(--font-head)">${rp(totK)}</b> <span class="xs faint">/ ${rpS(totT)} (${pct(totT ? totK / totT : 0, 1)})</span></div></div></div></div>

  <div class="sim-banner">
    <span class="icon-well">${ic('auto_awesome')}</span>
    <div class="grow" style="min-width:240px"><b>Simulasi Capaian Otomatis</b> <span class="pill ${onTrack ? 'pill-glass' : 'pill-warn'}" style="margin-left:6px">${onTrack ? 'On Track' : 'Perlu Ditingkatkan'}</span>
      <p>${sisa === 0 ? 'Semua target aktif sudah tercapai. Saatnya membuat target baru!' : sur > 0
        ? `Berdasarkan rata-rata surplus bulanan <b class="money">${rp(sur)}/bln</b>, seluruh target keluarga (sisa <b class="money">${rpS(sisa)}</b>) dapat tercapai dalam <b>${nBulan} bulan</b> (${mLabel(mAdd(mKey(T), nBulan))}).${butuh ? ` Agar semua tepat tenggat, setoran yang dibutuhkan <b class="money">${rp(butuh)}/bln</b>.` : ''}`
        : `Sisa target <b class="money">${rpS(sisa)}</b>. Rata-rata 3 bulan terakhir belum surplus — ${butuh ? `butuh setoran <b class="money">${rp(butuh)}/bln</b> agar tepat tenggat.` : 'tingkatkan pemasukan atau tekan pengeluaran.'}`}</p></div>
    <button class="btn" data-act="target-setor-pick">${ic('add_card')}Setor Tabungan</button>
  </div>

  <div class="grid g-4 mt-3">
    ${tampil.map((t) => {
      const p = targetPct(t), [sl, sc, si] = targetStatus(t), done = t.status === 'tercapai' || p >= 1;
      const d = t.dompet_id ? S.dMap.get(t.dompet_id) : null;
      const sisaBln = t.tenggat ? bulanAntara(T, t.tenggat) : null;
      const warnDl = t.tenggat && !done && sisaBln <= 2;
      return `<div class="card goal-card ${done ? 'done' : ''}">
        <div class="row" style="justify-content:space-between"><span class="icon-well" style="${tint(t.warna || '#0F5C3F')}">${ic(t.ikon || 'savings')}</span><span class="pill pill-${sc}">${si ? ic(si) : '<span class="dot"></span>'}${sl}</span></div>
        <div><div style="font:600 16px/1.3 var(--font-head)">${esc(t.nama)}</div><div class="xs faint row" style="gap:4px;margin-top:3px">${ic(done ? 'check_circle' : 'account_balance', 'sm')}${done ? 'Siap dibayarkan' : esc(d ? d.nama : 'Tanpa dompet khusus')}</div></div>
        <div><div class="row xs faint" style="justify-content:space-between"><span>${done ? 'Lengkap' : 'Terkumpul'}</span><span class="semi">${pct(p)}</span></div><b class="amt-big money nowrap" style="display:block;font-size:19px;color:${done ? 'var(--primary-text)' : 'var(--text)'}">${rp(t.terkumpul)}</b>
          <div class="mt-1">${progress(p, '', done ? 'primary' : p >= 0.8 ? 'warn' : 'primary')}</div>
          <div class="row xs faint mt-1" style="justify-content:space-between"><span class="money">Target: ${rpS(t.target_nominal)}</span><span class="money">${done ? 'Tuntas' : 'Sisa: ' + rpS(Math.max(0, t.target_nominal - t.terkumpul))}</span></div></div>
        ${done ? `<div class="deadline" style="background:var(--primary-50);color:var(--primary-text)">${ic('verified')}Tersimpan di ${esc(d ? d.nama : 'kas keluarga')}</div>`
          : `<div class="deadline ${warnDl ? 'warn' : ''}">${ic(warnDl ? 'hourglass_top' : 'event')}<span>${t.tenggat ? `Tenggat: <b>${mLabel(mKey(t.tenggat), true)}</b> (${sisaBln < 0 ? 'lewat ' + -sisaBln + ' bln' : 'sisa ' + sisaBln + ' bln'})` : 'Tanpa tenggat'}</span></div>`}
        <div class="foot">${done ? `<button class="btn btn-gold btn-sm grow" data-act="target-cair" data-id="${t.id}">${ic('redeem')}Siap Dicairkan</button>`
          : `<button class="btn ${p >= 0.8 ? 'btn-primary' : 'btn-soft'} btn-sm grow" data-act="target-setor" data-id="${t.id}">${ic('add')}${p >= 0.8 ? 'Setor Cepat' : 'Setor'}</button>`}
          <button class="btn-icon" data-act="target-edit" data-id="${t.id}" aria-label="Ubah target">${ic('edit')}</button></div>
      </div>`;
    }).join('')}
  </div>
  ${!tampil.length ? emptyBox('flag', 'Belum ada target tabungan', 'Rencanakan dana darurat, pendidikan anak, umrah, qurban, atau liburan keluarga.') : ''}
  <div class="cta-dashed mt-3"><span class="icon-well" style="background:var(--surface)">${ic('add', 'lg')}</span><div class="grow" style="min-width:200px"><b style="font:700 16px var(--font-head)">Buat Target Tabungan Baru</b><div class="small muted">Rencanakan dana pendidikan anak, uang muka rumah, atau proteksi keluarga lainnya.</div></div><button class="btn btn-primary" data-act="target-new">${ic('add_circle')}Tambah Pos Tabungan</button></div>
  ${cair.length ? `<div class="section-h"><h2 style="font-size:17px">Riwayat Target Dicairkan</h2></div><div class="card"><div class="list">${cair.map((t) => `<div class="li"><span class="icon-well sm" style="${tint(t.warna)}">${ic(t.ikon || 'savings')}</span><div class="grow"><div class="semi small">${esc(t.nama)}</div><div class="xs faint money">${rp(t.terkumpul)} dari ${rp(t.target_nominal)}</div></div><span class="pill">Dicairkan</span><button class="btn-icon" data-act="target-edit" data-id="${t.id}" aria-label="Ubah">${ic('edit')}</button></div>`).join('')}</div></div>` : ''}`;
}

// ── Modal target ──
let TGM = null;
function openTargetModal(t) {
  const edit = t ? { ...t } : null;
  TGM = { edit, ikon: edit ? edit.ikon || 'savings' : 'savings', warna: edit ? edit.warna || '#0F5C3F' : '#0F5C3F' };
  openModal({
    title: edit ? 'Ubah Target Tabungan' : 'Target Tabungan Baru', icon: 'flag',
    body: `<div class="field"><label for="tgNama">Nama Target <span class="req">*</span></label><input class="input" id="tgNama" maxlength="60" value="${esc(edit?.nama || '')}" placeholder="Contoh: Dana Pendidikan Anak"></div>
      <div class="form-grid mt-3">
        <div class="field"><label for="tgTarget">Nominal Target <span class="req">*</span></label><div class="money-input"><span class="cur">Rp</span><input id="tgTarget" inputmode="numeric" data-in="money" data-words="#tgW" value="${edit ? fmtN(edit.target_nominal) : ''}" placeholder="0"></div><div class="hint" id="tgW">${edit ? terbilang(edit.target_nominal) : ''}</div></div>
        <div class="field"><label for="tgKumpul">${edit ? 'Terkumpul Saat Ini' : 'Dana Awal (opsional)'}</label><div class="money-input"><span class="cur">Rp</span><input id="tgKumpul" inputmode="numeric" data-in="money" value="${edit ? fmtN(edit.terkumpul) : ''}" placeholder="0"></div></div>
        <div class="field"><label for="tgDompet">Disimpan di Dompet</label><select class="select" id="tgDompet">${dompetOpts(edit?.dompet_id || '', false, '— Tanpa dompet khusus —')}</select></div>
        <div class="field"><label for="tgTenggat">Tenggat</label><input type="date" class="input" id="tgTenggat" value="${edit?.tenggat || ''}" min="${today()}"></div>
      </div>
      <div class="field mt-3"><span class="lbl">Ikon</span><div class="icon-pick" id="tgIkon">${IKON_TARGET.map((n) => `<button type="button" data-act="tgm-ikon" data-v="${n}" aria-label="${n}">${ic(n)}</button>`).join('')}</div></div>
      <div class="field mt-3"><span class="lbl">Warna</span><div class="swatches" id="tgWarna">${WARNA.map(([h, n]) => `<button type="button" class="swatch" style="background:${h}" data-act="tgm-warna" data-v="${h}" title="${n}" aria-label="${n}"></button>`).join('')}</div></div>`,
    foot: `${edit ? `<button class="btn btn-danger-soft" data-act="tgm-del">${ic('delete')}Hapus</button>` : ''}<div class="grow"></div><button class="btn btn-ghost" data-act="modal-close">Batal</button><button class="btn btn-primary" data-act="tgm-save">${ic('check')}Simpan Target</button>`,
    onOpen: tgmRender
  });
}
function tgmRender() {
  $$('#tgIkon button').forEach((b) => b.classList.toggle('on', b.dataset.v === TGM.ikon));
  $$('#tgWarna .swatch').forEach((b) => b.classList.toggle('on', b.dataset.v.toLowerCase() === TGM.warna.toLowerCase()));
}
ACT['tgm-ikon'] = (el) => { TGM.ikon = el.dataset.v; tgmRender(); };
ACT['tgm-warna'] = (el) => { TGM.warna = el.dataset.v; tgmRender(); };
ACT['tgm-save'] = () => {
  const nama = $('#tgNama').value.trim(), target = moneyVal('#tgTarget'), kumpul = moneyVal('#tgKumpul'), tenggat = $('#tgTenggat').value;
  if (!nama) { toast('Nama target wajib diisi.', 'warn'); return; }
  if (!(target > 0)) { toast('Nominal target wajib angka positif.', 'warn'); return; }
  if (tenggat && !validDate(tenggat)) { toast('Tanggal tenggat tidak valid.', 'warn'); return; }
  const e = TGM.edit;
  let status = e ? e.status : 'aktif';
  if (status !== 'dicairkan') status = kumpul >= target ? 'tercapai' : 'aktif';
  DB.simpanTarget({ id: e ? e.id : uid('tg'), nama, target_nominal: target, terkumpul: kumpul, dompet_id: $('#tgDompet').value, tenggat, status, ikon: TGM.ikon, warna: TGM.warna });
  closeModal(); toast(e ? 'Target diperbarui.' : `Target ${nama} dibuat.`, 'success');
};
ACT['tgm-del'] = async () => {
  const t = TGM.edit;
  if (await confirmBox({ title: 'Hapus Target?', text: `Target <b>${esc(t.nama)}</b> akan dihapus. Transaksi transfer yang pernah dibuat tetap tersimpan.`, ok: 'Hapus', danger: true })) { DB.hapusTarget(t.id); toast('Target dihapus.', 'success'); }
};
ACT['target-new'] = () => openTargetModal(null);
ACT['target-edit'] = (el) => openTargetModal(S.target.find((t) => t.id === el.dataset.id));
ACT['target-cair'] = async (el) => {
  const t = S.target.find((x) => x.id === el.dataset.id);
  if (await confirmBox({ title: 'Cairkan Dana?', text: `Tandai target <b>${esc(t.nama)}</b> (${rp(t.terkumpul)}) sebagai sudah dicairkan/dipakai? Target akan dipindah ke riwayat.`, ok: 'Ya, Cairkan', icon: 'redeem' })) {
    DB.simpanTarget({ ...t, status: 'dicairkan' }); toast('Target ditandai sudah dicairkan.', 'success');
  }
};
ACT['target-setor-pick'] = () => {
  const list = S.target.filter((t) => t.status === 'aktif');
  if (!list.length) { toast('Belum ada target aktif.', 'info'); return; }
  if (list.length === 1) { openSetor(list[0]); return; }
  openModal({
    title: 'Pilih Target Tabungan', icon: 'savings', size: 'sm',
    body: `<div class="list">${list.map((t) => `<button class="li click" style="width:100%;text-align:left" data-act="target-setor" data-id="${t.id}"><span class="icon-well sm" style="${tint(t.warna)}">${ic(t.ikon || 'savings')}</span><span class="grow"><b class="small">${esc(t.nama)}</b><br><span class="xs faint">${pct(targetPct(t))} • sisa ${rpS(t.target_nominal - t.terkumpul)}</span></span>${ic('chevron_right', 'faint')}</button>`).join('')}</div>`
  });
};
ACT['target-setor'] = (el) => openSetor(S.target.find((t) => t.id === el.dataset.id));
let STM = null;
function openSetor(t) {
  if (!t) return;
  STM = { t, mode: 'setor' };
  const sisa = Math.max(0, t.target_nominal - t.terkumpul);
  const src = dompetAktif().find((d) => d.id !== t.dompet_id);
  openModal({
    title: 'Setor ke ' + t.nama, icon: 'savings',
    body: `<div class="mini-box"><div class="row" style="justify-content:space-between"><span class="small">Terkumpul <b class="money">${rp(t.terkumpul)}</b></span><span class="small faint money">Target ${rp(t.target_nominal)}</span></div><div class="mt-1">${progress(targetPct(t), '', 'primary')}</div></div>
      <div class="seg block mt-3" id="stSeg"><button type="button" class="on" data-act="st-mode" data-v="setor">${ic('add')}Setor</button><button type="button" data-act="st-mode" data-v="tarik">${ic('remove')}Tarik Dana</button></div>
      <div class="field mt-3"><label for="stAmt">Nominal <span class="req">*</span></label><div class="money-input"><span class="cur">Rp</span><input id="stAmt" inputmode="numeric" data-in="money" data-words="#stW" placeholder="0" autofocus>
        <div class="quick"><button type="button" data-act="st-q" data-v="500000">500rb</button><button type="button" data-act="st-q" data-v="1000000">1jt</button>${sisa ? `<button type="button" data-act="st-q" data-v="${sisa}">Lunasi</button>` : ''}</div></div><div class="hint" id="stW"></div></div>
      ${t.dompet_id && src ? `<label class="check mt-3"><input type="checkbox" id="stMove" checked><span><b class="small" id="stMoveLbl">Pindahkan dana dari dompet lain ke ${esc(S.dMap.get(t.dompet_id)?.nama || '')}</b><br><span class="hint">Dicatat sebagai transfer antar dompet (bukan pengeluaran).</span></span></label>
      <div class="form-grid mt-2"><div class="field"><label for="stSrc" id="stSrcLbl">Dari Dompet</label><select class="select" id="stSrc">${dompetOpts(src.id, true, '', dompetAktif().filter((d) => d.id !== t.dompet_id))}</select></div><div class="field"><label for="stTgl">Tanggal</label><input type="date" class="input" id="stTgl" value="${today()}"></div></div>`
      : `<div class="hint mt-2">${ic('info', 'sm')} Setoran dicatat ke progres target. ${t.dompet_id ? '' : 'Atur "Disimpan di Dompet" pada target untuk mencatat perpindahan dana otomatis.'}</div>`}`,
    foot: '<div class="grow"></div><button class="btn btn-ghost" data-act="modal-close">Batal</button><button class="btn btn-primary" data-act="st-save">' + ic('check') + 'Simpan</button>'
  });
}
ACT['st-mode'] = (el) => {
  STM.mode = el.dataset.v;
  $$('#stSeg button').forEach((b) => b.classList.toggle('on', b.dataset.v === STM.mode));
  const tn = S.dMap.get(STM.t.dompet_id)?.nama || '';
  if ($('#stMoveLbl')) $('#stMoveLbl').textContent = STM.mode === 'setor' ? `Pindahkan dana dari dompet lain ke ${tn}` : `Pindahkan dana dari ${tn} ke dompet lain`;
  if ($('#stSrcLbl')) $('#stSrcLbl').textContent = STM.mode === 'setor' ? 'Dari Dompet' : 'Ke Dompet';
};
ACT['st-q'] = (el) => { const i = $('#stAmt'); i.value = el.dataset.v; ON.money(i); };
ACT['st-save'] = () => {
  const t = STM.t, n = moneyVal('#stAmt');
  if (!(n > 0)) { toast('Isi nominal setoran.', 'warn'); return; }
  if (STM.mode === 'tarik' && n > t.terkumpul) { toast('Penarikan melebihi dana terkumpul.', 'warn'); return; }
  let tx = null;
  if ($('#stMove')?.checked) {
    const other = $('#stSrc').value, tanggal = $('#stTgl').value || today();
    tx = STM.mode === 'setor'
      ? { id: uid('tx'), tanggal, dompet_id: other, dompet_tujuan_id: t.dompet_id, nominal: n, catatan: 'Setor tabungan: ' + t.nama, dibuat_oleh: S.me.email, kategori_id: '' }
      : { id: uid('tx'), tanggal, dompet_id: t.dompet_id, dompet_tujuan_id: other, nominal: n, catatan: 'Tarik tabungan: ' + t.nama, dibuat_oleh: S.me.email, kategori_id: '' };
  }
  DB.setorTarget(t.id, STM.mode === 'setor' ? n : -n, tx);
  closeModal();
  const nt = S.target.find((x) => x.id === t.id);
  toast(STM.mode === 'setor' ? (nt.status === 'tercapai' ? `Selamat! Target ${t.nama} tercapai!` : `Setoran ${rp(n)} dicatat (${pct(targetPct(nt))}).`) : `Penarikan ${rp(n)} dicatat.`, 'success');
};

// ════════════════════════════════════════════════════════════
// 15. HALAMAN TAGIHAN & PENGINGAT
// ════════════════════════════════════════════════════════════
const TG = { f: 'semua' };
PAGES.tagihan = { t: 'Tagihan', r: renderTagihan, prep: (o) => { if (o.f) TG.f = o.f; } };
const REM = [[30, 'H-30'], [14, 'H-14'], [7, 'H-7'], [3, 'H-3'], [1, 'H-1'], [0, 'Hari-H']];
const remList = (g) => String(g.hari_pengingat || '').split(',').filter((x) => x !== '').map(Number);
function recipients() {
  const raw = String(S.settings.email_pengingat || '').split(/[,;\s]+/).filter((x) => RE_EMAIL.test(x));
  return raw.length ? raw : S.users.map((u) => u.email);
}
function lastPayTx(g) { return S.tx.find((t) => t.ref === 'tagihan:' + g.id); }
function renderTagihan() {
  const T = today(), mk = mKey(T);
  const all = S.tagihan.slice().sort((a, b) => (a.status === 'lunas') - (b.status === 'lunas') || a.tanggal_jatuh_tempo.localeCompare(b.tanggal_jatuh_tempo));
  const isLunas = (g) => g.status === 'lunas' || paidThisMonth(g);
  const F = {
    semua: () => true,
    akan_datang: (g) => billStatus(g) === 'akan_datang',
    jatuh_tempo: (g) => billStatus(g) === 'jatuh_tempo',
    terlambat: (g) => billStatus(g) === 'terlambat',
    lunas: isLunas
  };
  const rows = all.filter(F[TG.f] || F.semua);
  const rec = recipients();
  const hariDef = String(S.settings.hari_pengingat_default || '7,3,1,0').split(',').map(Number).sort((a, b) => b - a);
  const jamP = String(S.settings.jam_pengingat ?? 7).padStart(2, '0');
  const pending = all.filter((g) => g.status !== 'lunas' && g.tanggal_jatuh_tempo <= mEnd(mk));
  const days = mDays(mk);
  const pos = (d) => ((+d.slice(8) - 1) / (days - 1)) * 100;
  const nodes = [];
  all.forEach((g) => {
    if (g.status !== 'lunas' && mKey(g.tanggal_jatuh_tempo) === mk) nodes.push({ d: g.tanggal_jatuh_tempo, g, st: billStatus(g) });
    if (g.terakhir_dibayar && mKey(g.terakhir_dibayar) === mk) nodes.push({ d: g.terakhir_dibayar, g, st: 'lunas' });
  });
  const nodeColor = { lunas: 'var(--primary)', terlambat: 'var(--danger)', jatuh_tempo: 'var(--danger)', akan_datang: 'var(--info)' };
  const nodeIcon = { lunas: 'check', terlambat: 'priority_high', jatuh_tempo: 'priority_high', akan_datang: 'schedule' };

  $('#page-tagihan').innerHTML = `
  <div class="page-head"><div class="grow"><div class="crumb"><span class="tag red">${ic('alarm')}Jatuh Tempo &amp; Otomasi</span><span>• Pengingat Email Aktif</span></div><h1 class="page-title">Tagihan Rutin Keluarga</h1><p class="page-sub">Pastikan pengeluaran wajib keluarga terlunasi tepat waktu.</p></div>
    <div class="head-actions"><button class="btn btn-primary" data-act="bill-new">${ic('add_task')}Tambah Tagihan Baru</button></div></div>

  <div class="email-banner"><span class="icon-well">${ic('mail')}</span><div class="grow" style="min-width:220px"><b class="small">Pengingat email otomatis diaktifkan untuk ${rec.map((e) => `<a>${esc(e)}</a>`).join(' dan ')}</b>
    <div class="xs muted mt-1">Notifikasi dikirim pada ${hariDef.map((h) => (h ? 'H-' + h : 'Hari-H')).join(', ')} sebelum batas tempo (atur per tagihan) pukul ${jamP}:00 WIB.</div></div>
    <button class="link" data-go="pengaturan" data-tab="notifikasi">Ubah Jadwal ${ic('tune')}</button></div>

  <div class="chips mt-3">${[['semua', 'Semua'], ['akan_datang', 'Akan Datang'], ['jatuh_tempo', 'Jatuh Tempo'], ['terlambat', 'Terlambat'], ['lunas', 'Lunas']].map(([k, l]) => {
    const n = all.filter(F[k]).length;
    return `<button class="chip ${TG.f === k ? 'on' : ''}" data-act="tg-f" data-v="${k}">${l} <span class="cnt">(${n})</span>${(k === 'jatuh_tempo' || k === 'terlambat') && n ? '<span class="dot" style="width:7px;height:7px;border-radius:50%;background:var(--danger);display:inline-block"></span>' : ''}</button>`;
  }).join('')}</div>

  <div class="card mt-3">
    <div class="card-h"><span class="ms" style="color:var(--text-2)">calendar_month</span><div class="grow"><div class="card-t">Garis Waktu Jatuh Tempo • ${mLabel(mk)}</div></div>
      <div class="legend-inline hide-m"><span><i style="background:var(--primary)"></i>Lunas</span><span><i style="background:var(--danger)"></i>Mendesak</span><span><i style="background:var(--info)"></i>Akan Datang</span></div></div>
    <div class="timeline"><div class="track"><div class="line"></div>
      ${nodes.some((n) => +n.d.slice(8) <= 4) ? '' : `<div class="tick" style="left:0%;align-items:flex-start;transform:none"><span>01 ${BLN[+mk.slice(5) - 1]}</span></div>`}
      ${nodes.some((n) => +n.d.slice(8) >= days - 3) ? '' : `<div class="tick" style="left:100%;align-items:flex-end;transform:translateX(-100%)"><span>${days} ${BLN[+mk.slice(5) - 1]}</span></div>`}
      <div class="today" style="left:${pos(T)}%" title="Hari ini"></div>
      ${nodes.map((n, i) => `<div class="tick" style="left:${pos(n.d)}%;top:${i % 2 ? 0 : 0}px"><b>${+n.d.slice(8)} ${BLN[+n.d.slice(5, 7) - 1]}</b><span class="node" style="background:${nodeColor[n.st]}">${ic(nodeIcon[n.st])}</span><span class="tl-lbl" style="${n.st === 'lunas' ? tint('#2DB37A') : n.st === 'akan_datang' ? tint('#3B82F6') : tint('#EB5757')};margin-top:${i % 2 ? 22 : 0}px" title="${esc(n.g.nama)}">${esc(short(n.g.nama, 16))}${n.st === 'jatuh_tempo' && diffDays(T, n.d) === 1 ? ' (Besok!)' : ''}</span></div>`).join('')}
    </div></div>
  </div>

  <div class="card mt-3">
    ${rows.length ? `<div class="hide-m"><div class="tbl-wrap"><table class="tbl"><thead><tr><th>Nama &amp; Rincian Tagihan</th><th class="r">Nominal</th><th>Periode</th><th>Jatuh Tempo</th><th>Pengingat Email</th><th>Status</th><th class="r">Aksi</th></tr></thead><tbody>
      ${rows.map((g) => billRow(g, T)).join('')}</tbody></table></div></div>
      <div class="show-m list" style="display:none">${rows.map((g) => billCardM(g, T)).join('')}</div>`
      : emptyBox('receipt', TG.f === 'semua' ? 'Belum ada tagihan' : 'Tidak ada tagihan dengan status ini', 'Tambahkan tagihan rutin (SPP, listrik, cicilan, pajak) agar pengingat email aktif.', '<button class="btn btn-primary btn-sm" data-act="bill-new">Tambah Tagihan</button>')}
    <div class="row mt-3" style="border-top:1px solid var(--border);padding-top:14px;flex-wrap:wrap"><span class="small grow" style="min-width:240px">${ic('mail', 'sm faint')} Total kewajiban tertunda bulan ini: <b class="money">${rp(sum(pending, (g) => g.nominal))}</b> (${pending.length} tagihan)</span>
      <button class="link" data-go="pengaturan" data-tab="notifikasi">Kelola Pengaturan Pengingat Email ${ic('arrow_forward')}</button></div>
  </div>`;
}
function remChips(g, T) {
  const L = remList(g), d = diffDays(T, g.tanggal_jatuh_tempo);
  const base = [7, 3, 1, 0].concat(L.filter((x) => ![7, 3, 1, 0].includes(x))).sort((a, b) => b - a);
  return `<div class="rem-chips">${base.map((h) => `<span class="${L.includes(h) ? 'on' : ''} ${L.includes(h) && h === d ? 'today' : ''}" title="${h ? 'H-' + h : 'Hari-H'}">${h ? 'H-' + h : 'H'}</span>`).join('')}</div>`;
}
function billRow(g, T) {
  const st = billStatus(g), [sl, sc] = BILL_ST[st], k = S.kMap.get(g.kategori_id);
  const d = diffDays(T, g.tanggal_jatuh_tempo);
  const lunasBulanIni = st !== 'lunas' && paidThisMonth(g);
  const ptx = lastPayTx(g);
  const pd = ptx ? S.dMap.get(ptx.dompet_id) : null;
  if (st === 'lunas') {
    return `<tr class="paid-row"><td><div class="row">${katWell(k, 'sm')}<div><div class="t-title">${esc(g.nama)}</div><div class="t-sub">${ic('check_circle', 'sm')} Dibayar ${tgl(g.terakhir_dibayar, 'dm')}${pd ? ' via ' + esc(pd.nama) : ''}</div></div></div></td><td class="r"><span class="amt money" style="text-decoration:line-through">${rp(g.nominal)}</span></td><td><span class="pill">${PERIODE[g.periode]}</span></td><td class="small">${tgl(g.tanggal_jatuh_tempo)}</td><td><span class="xs faint">Tuntas</span></td><td><span class="pill pill-ok">${ic('check')}Lunas</span></td>
      <td class="r"><div class="row-actions"><button class="btn-icon" data-act="bill-edit" data-id="${g.id}" aria-label="Ubah">${ic('edit')}</button><button class="btn-icon danger" data-act="bill-del" data-id="${g.id}" aria-label="Hapus">${ic('delete')}</button></div></td></tr>`;
  }
  return `<tr><td><div class="row">${katWell(k, 'sm')}<div style="min-width:150px"><div class="t-title">${esc(g.nama)}</div><div class="t-sub">${esc(g.catatan || (k ? k.nama : ''))}</div>${lunasBulanIni ? `<div class="t-sub" style="color:var(--primary-text)">${ic('check_circle', 'sm')} Periode lalu lunas ${tgl(g.terakhir_dibayar, 'dm')}</div>` : ''}</div></div></td>
    <td class="r"><span class="amt money">${rp(g.nominal)}</span></td><td><span class="pill">${PERIODE[g.periode]}${g.periode === 'custom' ? ' ' + g.interval_hari + 'h' : ''}</span></td>
    <td><div class="semi small" style="color:${d <= 1 ? 'var(--danger-700)' : 'var(--text)'}">${tgl(g.tanggal_jatuh_tempo)}</div><div class="xs ${d <= 1 ? 'neg' : 'faint'}">${d < 0 ? 'Terlambat ' + -d + ' hari' : d === 0 ? 'Hari ini!' : d === 1 ? '● Besok!' : 'Sisa ' + d + ' hari'}</div></td>
    <td>${remChips(g, T)}</td><td><span class="pill pill-${sc}"><span class="dot"></span>${sl}</span></td>
    <td class="r"><div class="row-actions" style="align-items:center"><button class="btn ${d <= 1 ? 'btn-primary' : 'btn-soft'} btn-sm" data-act="bill-pay" data-id="${g.id}">${ic('done_all')}Tandai Lunas</button><button class="btn-icon" data-act="bill-edit" data-id="${g.id}" aria-label="Ubah">${ic('edit')}</button><button class="btn-icon danger" data-act="bill-del" data-id="${g.id}" aria-label="Hapus">${ic('delete')}</button></div></td></tr>`;
}
function billCardM(g, T) {
  const st = billStatus(g), [sl, sc] = BILL_ST[st], k = S.kMap.get(g.kategori_id), [txt, cls] = dueText(g);
  return `<div class="li" style="flex-wrap:wrap">${katWell(k, 'sm')}<div class="grow"><div class="t-title">${esc(g.nama)}</div><div class="due xs ${cls === 'red' ? 'neg' : 'faint'}">${esc(txt)} • ${PERIODE[g.periode]}</div></div><div class="right"><b class="money">${rp(g.nominal)}</b><div><span class="pill pill-${sc}" style="height:20px">${sl}</span></div></div>
    <div class="row" style="width:100%;justify-content:flex-end;gap:6px">${st !== 'lunas' ? `<button class="btn btn-soft btn-sm" data-act="bill-pay" data-id="${g.id}">${ic('done_all')}Tandai Lunas</button>` : ''}<button class="btn-icon" data-act="bill-edit" data-id="${g.id}" aria-label="Ubah">${ic('edit')}</button><button class="btn-icon danger" data-act="bill-del" data-id="${g.id}" aria-label="Hapus">${ic('delete')}</button></div></div>`;
}
ACT['tg-f'] = (el) => { TG.f = el.dataset.v; renderTagihan(); };
ACT['bill-new'] = () => goPage('tagihan-form', { id: '' });
ACT['bill-edit'] = (el) => goPage('tagihan-form', { id: el.dataset.id });
ACT['bill-del'] = async (el) => {
  const g = S.tagihan.find((x) => x.id === el.dataset.id);
  if (!g) return;
  if (await confirmBox({ title: 'Hapus Tagihan?', text: `Tagihan <b>${esc(g.nama)}</b> dan pengingat emailnya akan dihapus. Transaksi pembayaran sebelumnya tetap tersimpan.`, ok: 'Hapus', danger: true })) { DB.hapusTagihan(g.id); toast('Tagihan dihapus.', 'success'); }
};
ACT['bill-quick'] = () => {
  const list = unpaidBills();
  if (!list.length) { toast('Semua tagihan sudah lunas.', 'info'); return; }
  if (list.length === 1) { openBayar(list[0]); return; }
  openModal({
    title: 'Bayar Tagihan', icon: 'receipt', size: 'sm',
    body: `<div class="stack" style="gap:10px">${list.map((g) => { const [txt, cls] = dueText(g); return `<button class="bill-item" data-act="bill-pay" data-id="${g.id}">${katWell(S.kMap.get(g.kategori_id), 'sm')}<span class="grow"><span class="semi" style="display:block">${esc(g.nama)}</span><span class="due ${cls}">${esc(txt)}</span></span><b class="money">${rp(g.nominal)}</b></button>`; }).join('')}</div>`
  });
};
ACT['bill-pay'] = (el) => { closePopover(); const g = S.tagihan.find((x) => x.id === el.dataset.id); if (g) openBayar(g); };
function openBayar(g) {
  const k = S.kMap.get(g.kategori_id);
  const rec = g.periode !== 'sekali';
  const nd = nextDue(g);
  openModal({
    title: 'Tandai Lunas', icon: 'done_all',
    body: `<div class="mini-box row">${katWell(k)}<div class="grow"><div class="semi">${esc(g.nama)}</div><div class="xs faint">${PERIODE[g.periode]} • jatuh tempo ${tgl(g.tanggal_jatuh_tempo, 'long')}</div></div><b class="money" style="font-size:17px">${rp(g.nominal)}</b></div>
      <div class="form-grid mt-3">
        <div class="field"><label for="byTgl">Tanggal Bayar</label><input type="date" class="input" id="byTgl" value="${today()}"></div>
        <div class="field"><label for="byDp">Dibayar dari Dompet</label><select class="select" id="byDp">${dompetOpts(g.dompet_id && S.dMap.get(g.dompet_id)?.aktif ? g.dompet_id : '', true)}</select></div>
        <div class="field full"><label for="byAmt">Nominal Dibayar</label><div class="money-input"><span class="cur">Rp</span><input id="byAmt" inputmode="numeric" data-in="money" value="${fmtN(g.nominal)}"></div></div>
      </div>
      <label class="check mt-3"><input type="checkbox" id="byTx" checked><span><b class="small">Catat otomatis sebagai transaksi pengeluaran</b><br><span class="hint">Kategori: ${esc(k ? k.nama : 'Tagihan & Utilitas')} — saldo dompet langsung berkurang.</span></span></label>
      ${rec ? `<label class="check mt-2"><input type="checkbox" id="byNext" ${g.auto_lanjut !== false ? 'checked' : ''}><span><b class="small">Jadwalkan jatuh tempo berikutnya: ${tgl(nd, 'long')}</b><br><span class="hint">Pengingat email aktif kembali untuk periode berikutnya.</span></span></label>` : ''}`,
    foot: `<div class="grow"></div><button class="btn btn-ghost" data-act="modal-close">Batal</button><button class="btn btn-primary" data-act="by-save" data-id="${g.id}">${ic('check_circle')}Konfirmasi Lunas</button>`
  });
}
ACT['by-save'] = (el) => {
  const g = S.tagihan.find((x) => x.id === el.dataset.id);
  const tanggal = $('#byTgl').value, nominal = moneyVal('#byAmt'), dp = $('#byDp').value, buatTx = $('#byTx').checked;
  if (!validDate(tanggal)) { toast('Tanggal bayar tidak valid.', 'warn'); return; }
  if (buatTx && (!(nominal > 0) || !dp)) { toast('Isi nominal dan dompet pembayaran.', 'warn'); return; }
  const jt = DB.bayarTagihan(g, { tanggal, dompet_id: dp, nominal, buatTx, lanjut: $('#byNext') ? $('#byNext').checked : false });
  closeModal();
  toast(`${g.nama} lunas.${jt ? ' Jatuh tempo berikutnya ' + tgl(jt, 'long') + '.' : ''}`, 'success', 4500);
};

// ════════════════════════════════════════════════════════════
// 16. FORMULIR TAGIHAN (tambah / ubah) + draft lokal
// ════════════════════════════════════════════════════════════
const DRAFT_TG = 'kd_draft_tagihan';
let TGF = null;
PAGES['tagihan-form'] = {
  t: 'Formulir Tagihan', r: renderTagihanForm, noAuto: true,
  prep: (o) => {
    const g = o.id ? S.tagihan.find((x) => x.id === o.id) : null;
    if (g) TGF = { ...g, rem: remList(g), edit: true };
    else {
      const draft = LS.get(DRAFT_TG, null);
      const def = String(S.settings.hari_pengingat_default || '7,3,1,0').split(',').filter(Boolean).map(Number);
      TGF = draft && draft.nama !== undefined ? { ...draft, edit: false, restored: true } : { id: uid('tgh'), nama: '', kategori_id: S.kMap.has('kat_tagihan') ? 'kat_tagihan' : katByTipe('pengeluaran')[0]?.id || '', dompet_id: dompetAktif()[0]?.id || '', nominal: 0, catatan: '', periode: 'bulanan', tanggal_jatuh_tempo: addDays(today(), 7), interval_hari: 30, auto_lanjut: true, rem: def, edit: false };
    }
  }
};
function siklusText(f) {
  if (!validDate(f.tanggal_jatuh_tempo)) return 'Pilih tanggal jatuh tempo.';
  const d = +f.tanggal_jatuh_tempo.slice(8);
  if (f.periode === 'sekali') return `Dibayar sekali pada ${tgl(f.tanggal_jatuh_tempo, 'long')}.`;
  if (f.periode === 'bulanan') return `Ulangi setiap <b>1 bulan</b> pada tanggal <b>${d}</b> setiap bulannya.`;
  if (f.periode === 'tahunan') return `Ulangi setiap <b>1 tahun</b> pada <b>${d} ${BULAN[+f.tanggal_jatuh_tempo.slice(5, 7) - 1]}</b>.`;
  return `Ulangi setiap <b>${f.interval_hari || '?'} hari</b> sejak ${tgl(f.tanggal_jatuh_tempo)}.`;
}
function renderTagihanForm() {
  const f = TGF;
  if (!f) { goPage('tagihan'); return; }
  const rec = recipients();
  const recLbl = (e) => { const u = S.users.find((x) => x.email === e); return u ? u.label || u.nama : ''; };
  $('#page-tagihan-form').innerHTML = `
  <div class="row mb-3" style="flex-wrap:wrap"><button class="link" data-act="tgf-back">${ic('arrow_back')}Kembali ke Daftar Tagihan</button><span class="pill pill-ok">${f.edit ? 'Ubah Tagihan' : 'Formulir Tagihan Baru'}</span><div class="grow"></div><span class="xs faint row" style="gap:4px">${ic('verified_user', 'sm')}Sinkronisasi otomatis anggota keluarga</span></div>
  <h1 class="page-title">${f.edit ? 'Ubah Tagihan Rutin' : 'Tambah Tagihan Rutin Rumah Tangga'}</h1>
  <p class="page-sub mb-3">Catat dan atur pengingat berkala untuk tagihan utilitas, cicilan, dan kewajiban keluarga agar tidak terkena denda keterlambatan.</p>
  ${f.restored ? `<div class="note info mb-3">${ic('history')}<div class="grow">Draft formulir sebelumnya dipulihkan.</div><button class="link" data-act="tgf-clear">Mulai baru</button></div>` : ''}

  <div class="card step-card"><div class="step-h"><span class="step-no">01</span><div><div class="card-t">Informasi Tagihan</div><div class="card-s">Detail identitas kewajiban tagihan dan alokasi dompet sumber</div></div></div>
    <div class="field"><label for="gfNama">Nama Tagihan <span class="req">*</span></label><input class="input" id="gfNama" maxlength="80" data-in="tgf" value="${esc(f.nama)}" placeholder="Contoh: Pajak Kendaraan Bermotor"></div>
    <div class="form-grid mt-3">
      <div class="field"><label for="gfKat">Kategori Tagihan <span class="req">*</span></label><select class="select" id="gfKat" data-ch="tgf">${optList(katByTipe('pengeluaran'), f.kategori_id)}</select></div>
      <div class="field"><label for="gfDp">Dompet Pembayaran Default <span class="req">*</span></label><select class="select" id="gfDp" data-ch="tgf">${dompetOpts(f.dompet_id, true)}</select></div>
      <div class="field full"><label for="gfAmt">Nominal Tagihan <span class="req">*</span></label><div class="money-input big"><span class="cur">Rp</span><input id="gfAmt" inputmode="numeric" data-in="money" data-after="tgf" data-words="#gfWords2" value="${f.nominal ? fmtN(f.nominal) : ''}" placeholder="0"></div>
        <div class="terbilang">${ic('record_voice_over', 'sm')}<span>Terbilang:</span><i id="gfWords2">${f.nominal ? terbilang(f.nominal) : '-'}</i></div></div>
      <div class="field full"><label for="gfNote">Catatan / Nomor Pelanggan &amp; Referensi</label><textarea class="textarea" id="gfNote" maxlength="300" data-in="tgf" placeholder="Contoh: ID Pelanggan 5382 9102 4410 a.n. Budi">${esc(f.catatan || '')}</textarea></div>
    </div></div>

  <div class="card step-card mt-3"><div class="step-h"><span class="step-no">02</span><div><div class="card-t">Jadwal &amp; Frekuensi</div><div class="card-s">Tentukan interval pembayaran serta siklus berulang tagihan</div></div></div>
    <span class="lbl">Frekuensi Pembayaran <span class="req">*</span></span>
    <div class="grid g-4 mt-1" style="gap:12px" id="gfPer">${[['sekali', 'Sekali Bayar', 'Tidak berulang secara berkala'], ['bulanan', 'Bulanan', 'Jatuh tempo setiap bulan'], ['tahunan', 'Tahunan', 'Satu kali setiap tahun'], ['custom', 'Custom / Fleksibel', 'Mingguan / siklus khusus']].map(([v, l, s]) => `<button type="button" class="radio-card ${f.periode === v ? 'on' : ''}" data-act="tgf-per" data-v="${v}"><b>${l}</b><small>${s}</small></button>`).join('')}</div>
    <div class="field mt-3" id="gfIntWrap" ${f.periode === 'custom' ? '' : 'hidden'}><label for="gfInt">Ulangi setiap (hari) <span class="req">*</span></label><div class="row"><input type="number" class="input" id="gfInt" min="1" max="3660" style="max-width:140px" data-in="tgf" value="${f.interval_hari || 30}"><div class="chips">${[7, 14, 30, 90].map((n) => `<button type="button" class="chip" data-act="tgf-int" data-v="${n}">${n === 7 ? 'Mingguan' : n + ' hari'}</button>`).join('')}</div></div></div>
    <div class="form-grid mt-3">
      <div class="field"><label for="gfDue">Tanggal Jatuh Tempo ${f.edit ? '' : 'Pertama'} <span class="req">*</span></label><input type="date" class="input" id="gfDue" data-ch="tgf" value="${f.tanggal_jatuh_tempo}"><div class="hint" id="gfDueHint">${validDate(f.tanggal_jatuh_tempo) ? ic('info', 'sm') + ' Terjadwal: ' + tgl(f.tanggal_jatuh_tempo, 'day') : ''}</div></div>
      <div class="field"><span class="lbl">Aturan Siklus Pengulangan</span><div class="mini-box row small" id="gfSiklus">${ic('autorenew', 'faint')}<span class="sk-txt">${siklusText(f)}</span></div></div>
    </div>
    <label class="check mt-3 mini-box" id="gfAutoWrap" ${f.periode === 'sekali' ? 'hidden' : ''}><input type="checkbox" id="gfAuto" data-ch="tgf" ${f.auto_lanjut !== false ? 'checked' : ''}><span><b class="small">Otomatis jadwalkan jatuh tempo berikutnya setelah ditandai lunas</b><br><span class="hint" id="gfAutoHint"></span></span></label>
  </div>

  <div class="card step-card mt-3"><div class="step-h"><span class="step-no">03</span><div><div class="card-t">Pengaturan Pengingat Email</div><div class="card-s">Integrasi pemberitahuan otomatis ke kotak masuk email bersama</div></div></div>
    <div class="note info">${ic('mail')}<div>Email pengingat otomatis dikirim ke alamat terdaftar sebelum tanggal jatuh tempo untuk menjamin transparansi pengelolaan uang rumah tangga.</div></div>
    <div class="field mt-3"><span class="lbl">Penerima Notifikasi Tagihan</span><div class="chips">${rec.map((e) => `<span class="chip on-soft">${ic('check_circle')}${esc(e)}${recLbl(e) ? ` <span class="pill" style="height:18px">${esc(recLbl(e))}</span>` : ''}</span>`).join('')}<button type="button" class="chip" data-go="pengaturan" data-tab="notifikasi" style="border-style:dashed">${ic('add')}Kelola Penerima</button></div></div>
    <div class="field mt-3"><span class="lbl">Jadwal Kirim Pengingat</span><div class="hint">Pilih rentang waktu pengiriman email peringatan sebelum jatuh tempo:</div>
      <div class="chips mt-1" id="gfRem">${REM.map(([h, l]) => `<button type="button" class="chip toggle ${f.rem.includes(h) ? 'on' : ''}" data-act="tgf-rem" data-v="${h}" aria-pressed="${f.rem.includes(h)}">${f.rem.includes(h) ? ic('check') : ''}${l}${h === 7 ? ' (1 Minggu)' : h === 1 ? ' (Besok)' : h === 0 ? ` (Pukul ${String(S.settings.jam_pengingat ?? 7).padStart(2, '0')}.00 WIB)` : ''}</button>`).join('')}</div></div>
    <div class="mini-box mt-3 row" style="flex-wrap:wrap"><span class="ms faint">mail</span><div class="grow" style="min-width:220px"><div class="label-xs">Pratinjau Subjek Email</div><div class="semi small mt-1" id="gfSubj"></div><div class="xs faint mt-1">Pengirim: akun Google pemilik aplikasi • melalui Gmail</div></div>
      <button type="button" class="btn btn-outline btn-sm" data-act="tgf-test" data-loading="Mengirim…">${ic('send')}Kirim Email Uji Coba</button></div>
  </div>

  <div class="sticky-bar"><span class="status"><i></i><span id="gfDraftTxt">${f.edit ? 'Mengubah tagihan yang sudah ada' : 'Draft formulir tersimpan secara lokal'}</span></span>
    <button class="btn btn-outline" data-act="tgf-back">Batal</button><button class="btn btn-primary" data-act="tgf-save">${ic('check_circle')}Simpan Tagihan &amp; Aktifkan Pengingat</button></div>`;
  tgfDerived();
}
function tgfRead() {
  const f = TGF;
  f.nama = $('#gfNama').value; f.kategori_id = $('#gfKat').value; f.dompet_id = $('#gfDp').value; f.nominal = moneyVal('#gfAmt');
  f.catatan = $('#gfNote').value; f.tanggal_jatuh_tempo = $('#gfDue').value; f.interval_hari = +$('#gfInt').value || 30; f.auto_lanjut = $('#gfAuto').checked;
}
function tgfDerived() {
  const f = TGF;
  $('#gfSiklus .sk-txt').innerHTML = siklusText(f);
  $('#gfDueHint').innerHTML = validDate(f.tanggal_jatuh_tempo) ? ic('info', 'sm') + ' Terjadwal: ' + tgl(f.tanggal_jatuh_tempo, 'day') : '';
  const nd = nextDue(f);
  $('#gfAutoHint').textContent = nd ? `Sistem langsung membuat jadwal baru (${tgl(nd, 'long')}) setelah pembayaran dikonfirmasi oleh salah satu anggota keluarga.` : '';
  const h = f.rem.length ? Math.min(...f.rem.filter((x) => x > 0).concat(f.rem.includes(0) ? [0] : [])) : 3;
  const hh = f.rem.filter((x) => x > 0).sort((a, b) => a - b)[0] ?? 0;
  $('#gfSubj').textContent = `[Pengingat] Tagihan ${f.nama || 'Nama Tagihan'} (${rp(f.nominal)}) ${hh ? `jatuh tempo dalam ${hh} hari!` : 'jatuh tempo HARI INI'}`;
  void h;
}
const _tgfSave = debounce(() => { if (TGF && !TGF.edit) LS.set(DRAFT_TG, { ...TGF, restored: false }); }, 400);
ON.tgf = () => { if (!TGF) return; tgfRead(); tgfDerived(); _tgfSave(); };
CHG.tgf = ON.tgf;
ACT['tgf-per'] = (el) => {
  tgfRead(); TGF.periode = el.dataset.v;
  $$('#gfPer .radio-card').forEach((b) => b.classList.toggle('on', b.dataset.v === TGF.periode));
  $('#gfIntWrap').hidden = TGF.periode !== 'custom'; $('#gfAutoWrap').hidden = TGF.periode === 'sekali';
  tgfDerived(); _tgfSave();
};
ACT['tgf-int'] = (el) => { $('#gfInt').value = el.dataset.v; ON.tgf(); };
ACT['tgf-rem'] = (el) => {
  const v = +el.dataset.v, f = TGF;
  f.rem = f.rem.includes(v) ? f.rem.filter((x) => x !== v) : f.rem.concat(v);
  const on = f.rem.includes(v);
  el.classList.toggle('on', on); el.setAttribute('aria-pressed', on);
  const l = REM.find((r) => r[0] === v)[1];
  el.innerHTML = `${on ? ic('check') : ''}${l}${v === 7 ? ' (1 Minggu)' : v === 1 ? ' (Besok)' : v === 0 ? ` (Pukul ${String(S.settings.jam_pengingat ?? 7).padStart(2, '0')}.00 WIB)` : ''}`;
  tgfDerived(); _tgfSave();
};
ACT['tgf-back'] = () => goPage('tagihan');
ACT['tgf-clear'] = () => { LS.del(DRAFT_TG); PAGES['tagihan-form'].prep({}); TGF.restored = false; renderTagihanForm(); };
ACT['tgf-test'] = (el) => {
  tgfRead();
  const f = TGF;
  const hh = f.rem.filter((x) => x > 0).sort((a, b) => a - b)[0] ?? 0;
  withBtn(el, () => api('kirimEmailUji', { nama: f.nama || 'Tagihan Contoh', nominal: f.nominal, tanggal: f.tanggal_jatuh_tempo, hari: hh, catatan: f.catatan })
    .then((res) => toast(res.message + (res.data.sisaKuota != null ? ` (sisa kuota hari ini: ${res.data.sisaKuota})` : ''), 'success', 5000))
    .catch((e) => toast('Email gagal: ' + e.message, 'error', 6000)));
};
ACT['tgf-save'] = () => {
  tgfRead();
  const f = TGF;
  if (!f.nama.trim()) { toast('Nama tagihan wajib diisi.', 'warn'); $('#gfNama').focus(); return; }
  if (!(f.nominal > 0)) { toast('Nominal tagihan wajib angka positif.', 'warn'); $('#gfAmt').focus(); return; }
  if (!validDate(f.tanggal_jatuh_tempo)) { toast('Tanggal jatuh tempo tidak valid.', 'warn'); return; }
  if (f.periode === 'custom' && !(f.interval_hari >= 1)) { toast('Isi interval hari untuk periode custom.', 'warn'); return; }
  if (!f.kategori_id || !f.dompet_id) { toast('Pilih kategori dan dompet pembayaran.', 'warn'); return; }
  const lama = S.tagihan.find((x) => x.id === f.id);
  DB.simpanTagihan({
    id: f.id, nama: f.nama.trim(), nominal: f.nominal, periode: f.periode, tanggal_jatuh_tempo: f.tanggal_jatuh_tempo, hari_pengingat: f.rem.slice().sort((a, b) => b - a).join(','),
    kategori_id: f.kategori_id, dompet_id: f.dompet_id, status: lama && lama.status === 'lunas' && f.tanggal_jatuh_tempo <= (lama.tanggal_jatuh_tempo || '') ? 'lunas' : 'aktif',
    catatan: f.catatan.trim(), interval_hari: f.periode === 'custom' ? f.interval_hari : '', auto_lanjut: f.auto_lanjut, terakhir_dibayar: lama ? lama.terakhir_dibayar : ''
  });
  LS.del(DRAFT_TG);
  toast(f.edit ? 'Tagihan diperbarui.' : `Tagihan ${f.nama} disimpan & pengingat aktif.`, 'success');
  TGF = null;
  goPage('tagihan');
};


// ════════════════════════════════════════════════════════════
// 17. HALAMAN LAPORAN & ANALITIK
// ════════════════════════════════════════════════════════════
const R = { preset: 'bulan', dari: '', sampai: '', dompet: '', kat: '', q: '', oleh: '', page: 1, per: 10, loading: false };
PAGES.laporan = { t: 'Laporan', r: renderLaporan };
function rPreset(p) {
  const T = today(), mk = mKey(T);
  if (p === 'lalu') { const m = mAdd(mk, -1); return [m + '-01', mEnd(m)]; }
  if (p === 'tahun') return [T.slice(0, 4) + '-01-01', T];
  if (p === 'custom') return [R.dari || mk + '-01', R.sampai || T];
  return [mk + '-01', T];
}
function laporanData() {
  const [a, b] = [R.dari, R.sampai];
  const okF = (t) => (!R.dompet || t.dompet_id === R.dompet || t.dompet_tujuan_id === R.dompet) && (!R.kat || t.kategori_id === R.kat);
  const rows = S.tx.filter((t) => t.tanggal >= a && t.tanggal <= b && okF(t));
  const masuk = rows.filter((t) => t.tipe === 'masuk'), keluar = rows.filter((t) => t.tipe === 'keluar');
  const tin = sum(masuk, (t) => t.nominal), tout = sum(keluar, (t) => t.nominal), net = tin - tout;
  const L = diffDays(a, b) + 1;
  const hariBerjalan = Math.max(1, diffDays(a, b < today() ? b : today()) + 1);
  // periode pembanding: bulan/tahun sebelumnya pada rentang tanggal yang sama, atau periode sepanjang L sebelum a
  let pa, pb;
  if (R.preset === 'tahun') { pa = addMonths(a, -12); pb = addMonths(b, -12); }
  else if (a.slice(8) === '01' && mKey(a) === mKey(b)) { pa = addMonths(a, -1); pb = b === mEnd(mKey(b)) ? mEnd(mKey(pa)) : addMonths(b, -1); }
  else { pa = addDays(a, -L); pb = addDays(a, -1); }
  const prevOk = pa >= S.txDari;
  const prev = S.tx.filter((t) => t.tanggal >= pa && t.tanggal <= pb && okF(t));
  const pin = sum(prev.filter((t) => t.tipe === 'masuk'), (t) => t.nominal), pout = sum(prev.filter((t) => t.tipe === 'keluar'), (t) => t.nominal);
  const cat = {}, dom = {};
  keluar.forEach((t) => { cat[t.kategori_id] = (cat[t.kategori_id] || 0) + t.nominal; dom[t.dompet_id] = (dom[t.dompet_id] || 0) + t.nominal; });
  // seri bulanan: Januari tahun akhir periode s/d bulan akhir (maks 12 bulan, dalam data yang dimuat)
  const endM = mKey(b);
  let startM = b.slice(0, 4) + '-01';
  if (mKey(a) < startM) startM = mKey(a);
  if (startM < mKey(S.txDari)) startM = mKey(S.txDari);
  const months = [];
  for (let m = startM; m <= endM && months.length < 24; m = mAdd(m, 1)) months.push(m);
  if (months.length > 12) months.splice(0, months.length - 12);
  const mIn = {}, mOut = {};
  S.tx.forEach((t) => { if (!okF(t)) return; const m = mKey(t.tanggal); if (t.tipe === 'masuk') mIn[m] = (mIn[m] || 0) + t.nominal; else if (t.tipe === 'keluar') mOut[m] = (mOut[m] || 0) + t.nominal; });
  // tren saldo akhir bulan (seluruh dompet)
  const totNow = sum(S.dompet, (d) => d.saldo);
  const netByM = {};
  S.tx.forEach((t) => { if (t.tipe === 'transfer') return; const m = mKey(t.tanggal); netByM[m] = (netByM[m] || 0) + (t.tipe === 'masuk' ? t.nominal : -t.nominal); });
  const nowM = mKey(today());
  const saldoAkhir = months.map((m) => { let s = totNow; Object.entries(netByM).forEach(([k, v]) => { if (k > m && k <= nowM) s -= v; }); return s; });
  return { a, b, pa, rows, masuk, keluar, tin, tout, net, hariBerjalan, prevOk, pin, pout, cat, dom, months, mIn, mOut, saldoAkhir, L };
}
function renderLaporan() {
  if (!R.dari) [R.dari, R.sampai] = rPreset(R.preset);
  const need = addMonths(R.dari, -12) < addDays(R.dari, -(diffDays(R.dari, R.sampai) + 1)) ? addMonths(R.dari, -12) : addDays(R.dari, -(diffDays(R.dari, R.sampai) + 1));
  if (need < S.txDari && !R.loading) {
    R.loading = true;
    ensureLoaded(need).finally(() => { R.loading = false; if (CUR === 'laporan') renderLaporan(); });
  }
  const D = laporanData();
  const T = today(), mk = mKey(T);
  const chg = (cur, prev) => (D.prevOk && prev ? (cur - prev) / prev : null);
  const cIn = chg(D.tin, D.pin), cOut = chg(D.tout, D.pout);
  const avgDay = D.tout / D.hariBerjalan;
  const catItems = donutItems(D.cat, 4);
  const domRows = Object.entries(D.dom).map(([id, v]) => ({ d: S.dMap.get(id), v })).filter((x) => x.d).sort((x, y) => y.v - x.v);
  const ytd = D.saldoAkhir.length > 1 && D.saldoAkhir[0] ? (D.saldoAkhir[D.saldoAkhir.length - 1] - D.saldoAkhir[0]) / Math.abs(D.saldoAkhir[0]) : 0;
  const surplusMonths = D.months.filter((m) => (D.mIn[m] || 0) > (D.mOut[m] || 0)).length;
  const presets = [['bulan', `Bulan Ini (${mLabel(mk, true)})`], ['lalu', `Bulan Lalu (${mLabel(mAdd(mk, -1), true)})`], ['tahun', `Tahun Ini (${T.slice(0, 4)})`], ['custom', 'Rentang Tanggal Custom']];

  $('#page-laporan').innerHTML = `
  <div class="page-head"><div class="grow"><div class="crumb"><span class="tag">${ic('insights')}Periode Analisis Finansial</span><span>• ${tgl(D.a)} – ${tgl(D.b)}</span></div><h1 class="page-title">Laporan &amp; Analitik Keuangan</h1><p class="page-sub">Rekapitulasi pemasukan, pengeluaran, serta arus kas keluarga secara terperinci.</p></div>
    <div class="head-actions"><button class="btn btn-outline" data-act="lap-xlsx" data-loading="Menyiapkan…">${ic('table_view')}Ekspor Excel (.xlsx)</button><button class="btn btn-primary" data-act="lap-print">${ic('print')}Cetak / Simpan PDF</button></div></div>

  <div class="card">
    <div>
      <div class="chips mb-2">${presets.map(([k, l]) => `<button class="chip ${R.preset === k ? 'on-soft' : ''}" data-act="lap-preset" data-v="${k}">${R.preset === k ? ic('check') : ''}${l}</button>`).join('')}</div>
      <div class="filter-card">
        <span class="date-range">${ic('calendar_today')}<input type="date" id="lapDari" value="${D.a}" aria-label="Dari tanggal"><span class="faint">sampai</span><input type="date" id="lapSampai" value="${D.b}" aria-label="Sampai tanggal"></span>
        <select class="select sm" style="width:auto" id="lapDp" aria-label="Dompet">${optList(dompetSemua(), R.dompet, `Semua Dompet (${S.dompet.length})`)}</select>
        <select class="select sm" style="width:auto" id="lapKat" aria-label="Kategori">${optList(S.kategori.slice().sort((x, y) => x.nama.localeCompare(y.nama)), R.kat, 'Semua Kategori')}</select>
        <button class="btn btn-primary btn-sm" data-act="lap-apply">${ic('tune')}Terapkan Filter</button>
      </div>
    </div>
  </div>

  <div class="grid g-4 mt-3">
    <div class="card kpi"><div class="kpi-top"><span class="kpi-l">Total Pemasukan</span><span class="icon-well sm">${ic('arrow_downward')}</span></div><div class="kpi-v money">${rp(D.tin)}</div>
      <div class="kpi-f">${cIn == null ? '<span>—</span>' : `<span class="pill ${cIn >= 0 ? 'pill-ok' : 'pill-danger'}">${ic(cIn >= 0 ? 'trending_up' : 'trending_down')}${cIn >= 0 ? '+' : ''}${pct(cIn, 1)}</span>`} vs periode lalu</div></div>
    <div class="card kpi"><div class="kpi-top"><span class="kpi-l">Total Pengeluaran</span><span class="icon-well sm" style="${tint('#F5B13D')}">${ic('arrow_upward')}</span></div><div class="kpi-v money">${rp(D.tout)}</div>
      <div class="kpi-f">${cOut == null ? '<span>—</span>' : `<span class="pill ${cOut <= 0 ? 'pill-ok' : 'pill-danger'}">${ic(cOut <= 0 ? 'trending_down' : 'trending_up')}${cOut > 0 ? '+' : ''}${pct(cOut, 1)} ${cOut <= 0 ? 'hemat' : ''}</span>`} vs periode lalu</div></div>
    <div class="card kpi hl"><div class="kpi-top"><span class="kpi-l">Selisih Bersih (Net Cashflow)</span><span class="pill ${D.net >= 0 ? 'pill-solid' : 'pill-danger'}">${D.net >= 0 ? (D.tin && D.net / D.tin >= 0.3 ? 'Surplus Kuat' : 'Surplus') : 'Defisit'}</span></div><div class="kpi-v money" style="color:${D.net >= 0 ? 'var(--primary-text)' : 'var(--danger-700)'}">${D.net >= 0 ? '+' : ''}${rp(D.net)}</div>
      <div class="kpi-f">${ic(D.net >= 0 ? 'check_circle' : 'error', 'sm')}${D.net >= 0 ? 'Arus kas sehat & produktif' : 'Pengeluaran melebihi pemasukan'}</div></div>
    <div class="card kpi"><div class="kpi-top"><span class="kpi-l">Rata-rata Belanja Harian</span><span class="icon-well sm" style="background:var(--surface-3);color:var(--text-2)">${ic('calendar_month')}</span></div><div class="kpi-v money">${rp(avgDay)}</div>
      <div class="kpi-f">/ hari <span class="pill">${D.hariBerjalan} hari berjalan</span></div></div>
  </div>

  <div class="grid g-2 mt-3">
    <div class="card"><div class="card-h"><div class="grow"><div class="card-t">Perbandingan Arus Kas Bulanan</div><div class="card-s">Pemasukan vs Pengeluaran (${D.months.length ? mLabel(D.months[0]) + ' – ' + mLabel(D.months[D.months.length - 1]) : '-'})</div></div>
      <div class="chart-legend"><span><i style="background:var(--primary)"></i>Pemasukan</span><span><i style="background:var(--accent);opacity:.55"></i>Pengeluaran</span></div></div>
      <div class="chart-box lg"><canvas id="chLapBar" role="img" aria-label="Grafik pemasukan dan pengeluaran per bulan"></canvas></div>
      <div class="row small mt-2" style="justify-content:space-between;border-top:1px solid var(--divider);padding-top:10px;flex-wrap:wrap"><span class="faint">Rata-rata selisih: <span class="money">${rpS(D.months.length ? sum(D.months, (m) => (D.mIn[m] || 0) - (D.mOut[m] || 0)) / D.months.length : 0)}/bln</span></span><b>${surplusMonths} dari ${D.months.length} bulan surplus</b></div></div>
    <div class="card"><div class="card-h"><div class="grow"><div class="card-t">Komposisi Pengeluaran</div><div class="card-s">Distribusi kategori utama pada periode terpilih</div></div><span class="pill pill-ok money">${rp(D.tout)} Total</span></div>
      <div class="donut-wrap" style="min-height:240px"><div class="donut" style="width:190px;height:190px"><canvas id="chLapDonut" role="img" aria-label="Grafik komposisi pengeluaran"></canvas><div class="donut-center"><small>PENGELUARAN</small><b>100%</b><small>Terbagi rapi</small></div></div><div class="legend">${legendHtml(catItems)}</div></div></div>
  </div>

  <div class="grid g-2 mt-3">
    <div class="card"><div class="card-h"><div class="grow"><div class="card-t">Tren Akumulasi Saldo Keluarga</div><div class="card-s">Saldo seluruh dompet di akhir setiap bulan</div></div><span class="pill ${ytd >= 0 ? 'pill-ok' : 'pill-danger'}">${ic(ytd >= 0 ? 'trending_up' : 'trending_down')}${ytd >= 0 ? '+' : ''}${pct(ytd, 1)}</span></div>
      <div class="chart-box"><canvas id="chLapLine" role="img" aria-label="Grafik tren saldo keluarga"></canvas></div></div>
    <div class="card"><div class="card-h"><div class="grow"><div class="card-t">Pengeluaran Berdasarkan Dompet</div><div class="card-s">Kanal pengeluaran harian dan tagihan</div></div><span class="xs faint">${domRows.length} kanal aktif</span></div>
      <div class="bar-list">${domRows.length ? domRows.map((x) => `<div class="bl"><div class="row"><span class="row grow" style="gap:8px">${ic(JENIS_DP[x.d.jenis]?.[1] || 'wallet', 'sm faint')}<span class="semi ellipsis">${esc(x.d.nama)}</span></span><span class="faint xs">${pct(x.v / (D.tout || 1), 1)}</span><b class="money">${rp(x.v)}</b></div><div class="progress"><i style="width:${((x.v / (domRows[0].v || 1)) * 100).toFixed(1)}%;background:${x.d.warna}"></i></div></div>`).join('') : '<div class="small faint">Belum ada pengeluaran.</div>'}</div></div>
  </div>

  <div class="card mt-3" id="lapLog"></div>`;
  renderLapLog();

  whenChart(() => {
    const c = CC();
    mkChart('chLapBar', {
      type: 'bar',
      data: { labels: D.months.map((m) => BLN[+m.slice(5) - 1]), datasets: [
        { label: 'Pemasukan', data: D.months.map((m) => D.mIn[m] || 0), backgroundColor: c.primary, borderRadius: 5, barPercentage: 0.7, categoryPercentage: 0.7 },
        { label: 'Pengeluaran', data: D.months.map((m) => D.mOut[m] || 0), backgroundColor: rgba(c.accent, 0.5), borderRadius: 5, barPercentage: 0.7, categoryPercentage: 0.7 }] },
      options: { scales: { x: xAxis(c), y: yAxis(c) }, plugins: { tooltip: { ...TIP(), mode: 'index', intersect: false, callbacks: { label: (x) => ` ${x.dataset.label}: ${rp(x.parsed.y)}` } } } }
    });
    drawDonut('chLapDonut', catItems);
    mkChart('chLapLine', {
      type: 'line',
      data: { labels: D.months.map((m) => BLN[+m.slice(5) - 1] + (D.months.length > 12 ? ' ' + m.slice(2, 4) : '')), datasets: [{ label: 'Saldo', data: D.saldoAkhir, borderColor: c.primary, borderWidth: 2.5, tension: 0.38, pointRadius: (x) => (x.dataIndex === D.saldoAkhir.length - 1 ? 5 : 3), pointBackgroundColor: (x) => (x.dataIndex === D.saldoAkhir.length - 1 ? c.accent : c.primary), fill: true,
        backgroundColor: (ctx) => { const { chart } = ctx; const { ctx: g, chartArea } = chart; if (!chartArea) return rgba(c.accent, 0.12); const gr = g.createLinearGradient(0, chartArea.top, 0, chartArea.bottom); gr.addColorStop(0, rgba(c.accent, 0.25)); gr.addColorStop(1, rgba(c.accent, 0)); return gr; } }] },
      options: { scales: { x: xAxis(c), y: { ...yAxis(c), beginAtZero: false } }, plugins: { tooltip: { ...TIP(), callbacks: { label: (x) => ` Saldo akhir bulan: ${rp(x.parsed.y)}` } } } }
    });
  });
}
function renderLapLog() {
  const D = laporanData();
  const q = R.q.trim().toLowerCase();
  const rows = D.rows.filter((t) => (!R.oleh || t.dibuat_oleh === R.oleh) && (!q || txSearchText(t).includes(q)));
  const page = clamp(R.page, 1, Math.max(1, Math.ceil(rows.length / R.per)));
  R.page = page;
  const sl = rows.slice((page - 1) * R.per, page * R.per);
  const tipePill = (t) => (t.tipe === 'masuk' ? '<span class="pill pill-ok">Pemasukan</span>' : t.tipe === 'keluar' ? '<span class="pill pill-warn">Pengeluaran</span>' : '<span class="pill pill-info">Transfer</span>');
  $('#lapLog').innerHTML = `
    <div class="card-h" style="flex-wrap:wrap"><div class="grow"><div class="card-t">Log Detail Transaksi</div><div class="card-s">Data pembukuan terverifikasi rentang ${tgl(D.a, 'num')} – ${tgl(D.b, 'num')}</div></div>
      <div class="input-icon" style="min-width:200px">${ic('search')}<input class="input sm" type="search" placeholder="Cari dalam tabel…" value="${esc(R.q)}" data-in="lap-q" aria-label="Cari dalam tabel"></div>
      <div class="seg">${[['', 'Semua']].concat(S.users.map((u) => [u.email, 'Diinput ' + u.nama])).map(([v, l]) => `<button class="${R.oleh === v ? 'on' : ''}" data-act="lap-oleh" data-v="${esc(v)}">${esc(l)}</button>`).join('')}</div></div>
    ${rows.length ? `<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Tanggal</th><th>Deskripsi</th><th>Kategori</th><th>Dompet</th><th>Penginput</th><th>Tipe</th><th class="r">Nominal</th></tr></thead><tbody>
      ${sl.map((t) => `<tr class="click" data-act="tx-edit" data-id="${t.id}" style="cursor:pointer"><td class="small nowrap">${tgl(t.tanggal)}</td><td><span class="t-title">${esc(txTitle(t))}</span></td><td>${txKatPill(t)}</td><td class="small muted nowrap">${txDompet(t)}</td><td>${whoMini(t.dibuat_oleh)}</td><td>${tipePill(t)}</td><td class="r">${txAmt(t)}</td></tr>`).join('')}
      </tbody><tfoot><tr><td colspan="4">${ic('summarize', 'sm')} Total ringkasan: <b>${fmtN(rows.length)} transaksi</b></td><td colspan="3" class="r small">Pengeluaran: <b class="money">${rp(sum(rows.filter((t) => t.tipe === 'keluar'), (t) => t.nominal))}</b> &nbsp; Pemasukan: <b class="money pos">+${rp(sum(rows.filter((t) => t.tipe === 'masuk'), (t) => t.nominal))}</b></td></tr></tfoot></table></div>
      ${pager(rows.length, page, R.per, 'lap-page')}` : emptyBox('search_off', 'Tidak ada transaksi pada filter ini')}`;
}
const _lapQ = debounce(() => { R.page = 1; renderLapLog(); }, 200);
ON['lap-q'] = (el) => { R.q = el.value; _lapQ(); };
ACT['lap-oleh'] = (el) => { R.oleh = el.dataset.v; R.page = 1; renderLapLog(); };
ACT['lap-page'] = (el) => { R.page = +el.dataset.p; renderLapLog(); };
ACT['lap-preset'] = (el) => { R.preset = el.dataset.v; if (R.preset !== 'custom') [R.dari, R.sampai] = rPreset(R.preset); R.page = 1; renderLaporan(); };
ACT['lap-apply'] = () => {
  const a = $('#lapDari').value, b = $('#lapSampai').value;
  if (!validDate(a) || !validDate(b) || a > b) { toast('Rentang tanggal tidak valid.', 'warn'); return; }
  const [pa, pb] = rPreset(R.preset);
  if (a !== pa || b !== pb) R.preset = 'custom';
  R.dari = a; R.sampai = b; R.dompet = $('#lapDp').value; R.kat = $('#lapKat').value; R.page = 1;
  renderLaporan();
  toast('Filter laporan diterapkan.', 'success', 1800);
};

// ── Ekspor Excel ──
function lapLabelFilter() {
  return [R.dompet ? 'Dompet: ' + (S.dMap.get(R.dompet)?.nama || '') : 'Semua dompet', R.kat ? 'Kategori: ' + (S.kMap.get(R.kat)?.nama || '') : 'Semua kategori'].join(' • ');
}
function lapSheets() {
  const D = laporanData();
  const nama = S.settings.nama_aplikasi || 'Keuangan Digital';
  const cats = Object.entries(D.cat).map(([id, v]) => [S.kMap.get(id)?.nama || 'Tanpa kategori', v, D.tout ? v / D.tout : 0]).sort((a, b) => b[1] - a[1]);
  const ringkasan = [
    [`${nama} — Laporan Rekapitulasi Arus Kas`], [`Keluarga ${familyName()}`], [],
    ['Periode', `${tgl(D.a, 'long')} s/d ${tgl(D.b, 'long')}`], ['Filter', lapLabelFilter()], ['Dicetak', `${tgl(today(), 'long')} ${nowHMS().slice(0, 5)} WIB oleh ${S.me.nama}`], [],
    ['RINGKASAN', 'Nominal (Rp)'], ['Total Pemasukan', D.tin], ['Total Pengeluaran', D.tout], ['Selisih Bersih', D.net], ['Rata-rata Belanja Harian', Math.round(D.tout / D.hariBerjalan)], ['Jumlah Transaksi', D.rows.length], [],
    ['PENGELUARAN PER KATEGORI', 'Nominal (Rp)', 'Porsi'], ...cats, [],
    ['POSISI SALDO DOMPET (saat ini)', 'Saldo (Rp)'], ...dompetAktif().map((d) => [d.nama, d.saldo]), ['Total Saldo', totalSaldo()]
  ];
  const transaksi = [['Tanggal', 'Tipe', 'Deskripsi', 'Kategori', 'Dompet', 'Dompet Tujuan', 'Nominal (Rp)', 'Diinput Oleh', 'Waktu Input'],
    ...D.rows.slice().reverse().map((t) => [t.tanggal, t.tipe === 'masuk' ? 'Pemasukan' : t.tipe === 'keluar' ? 'Pengeluaran' : 'Transfer', txTitle(t), t.tipe === 'transfer' ? 'Transfer Dompet' : S.kMap.get(t.kategori_id)?.nama || '', S.dMap.get(t.dompet_id)?.nama || '', S.dMap.get(t.dompet_tujuan_id)?.nama || '', t.tipe === 'keluar' ? -t.nominal : t.nominal, userOf(t.dibuat_oleh).nama, t.dibuat_pada])];
  const bulanan = [['Bulan', 'Pemasukan (Rp)', 'Pengeluaran (Rp)', 'Selisih (Rp)', 'Saldo Akhir Bulan (Rp)'], ...D.months.map((m, i) => [mLabel(m), D.mIn[m] || 0, D.mOut[m] || 0, (D.mIn[m] || 0) - (D.mOut[m] || 0), D.saldoAkhir[i]])];
  return { D, ringkasan, transaksi, bulanan, file: `Laporan_Keuangan_${D.a}_sd_${D.b}.xlsx` };
}
ACT['lap-xlsx'] = (btn) => withBtn(btn, async () => {
  const X = lapSheets();
  try {
    await loadScript(XLSX_URL);
    const wb = XLSX.utils.book_new();
    const add = (aoa, name, cols, pctCol) => {
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = cols.map((w) => ({ wch: w }));
      Object.keys(ws).forEach((k) => { if (k[0] === '!') return; const c = ws[k]; if (c.t === 'n') c.z = pctCol && XLSX.utils.decode_cell(k).c === pctCol && c.v <= 1 ? '0.0%' : '#,##0'; });
      XLSX.utils.book_append_sheet(wb, ws, name);
    };
    add(X.ringkasan, 'Ringkasan', [36, 26, 12], 2);
    add(X.transaksi, 'Transaksi', [12, 13, 40, 22, 22, 22, 16, 14, 20]);
    add(X.bulanan, 'Bulanan', [18, 18, 18, 18, 22]);
    XLSX.writeFile(wb, X.file, { compression: true });
    toast('File Excel diunduh: ' + X.file, 'success', 4000);
  } catch (e) {
    toast('Unduhan langsung tidak tersedia, membuat Google Sheets…', 'info');
    try {
      const res = await api('eksporKeSpreadsheet', X.file.replace('.xlsx', ''), [{ nama: 'Ringkasan', rows: X.ringkasan }, { nama: 'Transaksi', rows: X.transaksi }, { nama: 'Bulanan', rows: X.bulanan }]);
      openModal({ title: 'Laporan Siap', icon: 'table_view', size: 'sm', body: `<p class="small muted">Laporan dibuat sebagai Google Sheets di folder <b>Ekspor</b>. Buka lalu pilih <b>File › Download › Microsoft Excel (.xlsx)</b>.</p>`, foot: `<div class="grow"></div><a class="btn btn-primary" href="${esc(res.data.url)}" target="_blank" rel="noopener">${ic('open_in_new')}Buka Google Sheets</a>` });
    } catch (e2) { toast('Ekspor gagal: ' + e2.message, 'error', 6000); }
  }
});

// ════════════════════════════════════════════════════════════
// 18. PRATINJAU CETAK & PDF
// ════════════════════════════════════════════════════════════
const DOC_CSS = `
.doc{width:794px;min-height:1123px;background:#fff;color:#13231B;padding:48px 52px;box-sizing:border-box;font-family:Inter,Arial,Helvetica,sans-serif;font-size:11px;line-height:1.45;box-shadow:0 10px 40px rgba(0,0,0,.18);margin:0 auto}
.doc table{border-collapse:collapse;width:100%}.doc td,.doc th{vertical-align:top}
.d-head td{vertical-align:middle}.d-logo{width:62px}.d-logo .lg{width:50px;height:50px;border-radius:12px;background:#E7F4EC;color:#0F5C3F;text-align:center;line-height:50px;font-weight:800;font-size:20px;overflow:hidden}.d-logo img{width:50px;height:50px;border-radius:12px;object-fit:cover}
.d-title{font-family:'Plus Jakarta Sans',Arial,sans-serif;font-size:22px;font-weight:800;letter-spacing:-.02em}.d-badge{display:inline-block;font-size:9px;font-weight:700;letter-spacing:.08em;background:#E7F4EC;color:#0F5C3F;padding:3px 7px;border-radius:5px;vertical-align:middle;margin-left:6px}
.d-sub{color:#0F5C3F;font-weight:600;font-size:12px;margin-top:2px}.d-sub2{color:#8BA296;font-size:10.5px}
.d-meta{width:250px}.d-meta table{border:1px solid #E2ECE6;background:#F7FAF8}.d-meta td{padding:4px 10px;font-size:10px;color:#587265}.d-meta td b{color:#13231B}.mono{font-family:ui-monospace,Menlo,Consolas,monospace}
.d-rule{border-top:2px solid #0F5C3F;margin:18px 0 16px}
.d-sec{font-weight:700;font-size:11px;letter-spacing:.04em;margin:16px 0 8px}.d-sec .d-r{float:right;font-weight:400;color:#8BA296;letter-spacing:0;font-size:10px}
.d-cards td{width:33.33%;padding:0 5px}.d-cards td:first-child{padding-left:0}.d-cards td:last-child{padding-right:0}
.d-card{border:1px solid #E2ECE6;border-radius:10px;padding:12px 14px;min-height:78px}.d-card .k{font-size:9.5px;font-weight:700;letter-spacing:.05em;color:#587265}.d-card .v{font-family:'Plus Jakarta Sans',Arial,sans-serif;font-size:21px;font-weight:800;margin:6px 0 4px}.d-card .s{font-size:9.5px;color:#587265}
.d-card .in{color:#0F5C3F}.d-card .out{color:#C62828}.d-card.ok{background:#E7F4EC;border-color:#BFE3CF}.d-card.bad{background:#FDF2F2;border-color:#F5C2C2}.d-card.warn{background:#FFF8F7}
.d-two td{width:50%;padding:0 6px}.d-two td:first-child{padding-left:0}.d-two td:last-child{padding-right:0}
.d-box{border:1px solid #E2ECE6;border-radius:10px;padding:12px 14px;margin-top:12px}.d-bt{font-size:10px;font-weight:700;letter-spacing:.04em;margin-bottom:8px}.d-bt span{float:right;font-weight:400;color:#8BA296;letter-spacing:0}
.d-row{width:100%}.d-row td{padding:3px 0;font-size:10.5px}.d-row td.r{text-align:right;white-space:nowrap;font-weight:600}.d-row td.r em{font-style:normal;color:#8BA296;font-weight:400;font-size:9.5px}
.d-bar{height:5px;background:#EEF4F0;border-radius:4px;margin:2px 0 5px}.d-bar i{display:block;height:5px;border-radius:4px;background:#0F5C3F}
.d-dot{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:6px}
.d-tot td{border-top:1px solid #E2ECE6;padding-top:7px!important;font-weight:700}
.d-tbl{border:1px solid #E2ECE6;margin-top:4px}.d-tbl th{background:#F7FAF8;font-size:9.5px;text-align:left;padding:8px;border-bottom:1px solid #E2ECE6;font-weight:700}.d-tbl td{padding:7px 8px;border-bottom:1px solid #F0F4F1;font-size:10px}
.d-tbl .r{text-align:right;white-space:nowrap}.d-tbl .in{color:#0F5C3F;font-weight:700}.d-tbl .out{color:#C62828;font-weight:700}.d-tbl .tr{color:#1D4ED8;font-weight:700}.d-tbl tfoot td{background:#F7FAF8;font-weight:700}
.d-tag{display:inline-block;font-size:8.5px;font-weight:700;padding:2px 6px;border-radius:4px;background:#E7F4EC;color:#0F5C3F}
.d-sign{margin-top:34px;border-top:1px solid #E2ECE6}.d-sign td{width:50%;text-align:center;padding-top:16px;color:#587265;font-size:10px}
.d-sig{font-family:'Brush Script MT','Segoe Script',cursive;font-style:italic;color:#0F5C3F;font-size:22px;margin:16px 0 4px;font-weight:700}.d-line{width:170px;border-top:1px solid #13231B;margin:0 auto 4px}.d-sign b{color:#13231B;font-size:11px}
.d-foot{margin-top:26px;border-top:1px solid #E2ECE6;padding-top:10px;font-size:9px;color:#8BA296}.d-foot .r{float:right}
`;
function docHtml(D, forPdf) {
  const set = S.settings, nama = set.nama_aplikasi || 'Keuangan Digital';
  const cats = Object.entries(D.cat).map(([id, v]) => ({ n: S.kMap.get(id)?.nama || 'Tanpa kategori', v })).sort((a, b) => b.v - a.v);
  const top = cats.slice(0, 5); const lain = sum(cats.slice(5), (x) => x.v); if (lain) top.push({ n: 'Lain-lain', v: lain });
  const rate = D.tin ? D.net / D.tin : 0;
  const no = `KD/REP/${D.b.slice(0, 7).replace('-', '')}/${String(Date.now()).slice(-3)}`;
  const u0 = S.users[0] || S.me, u1 = S.users[1];
  const rows = D.rows.slice().reverse();
  const logo = set.logo_url && !forPdf ? `<img src="${esc(set.logo_url)}" alt="" referrerpolicy="no-referrer">` : `<div class="lg">${esc(nama.charAt(0))}</div>`;
  return `<div class="doc">
    <table class="d-head"><tr><td class="d-logo">${logo}</td>
      <td><div class="d-title">${esc(nama)}<span class="d-badge">LAPORAN</span></div><div class="d-sub">Laporan Rekapitulasi Arus Kas Keluarga ${esc(familyName())}</div><div class="d-sub2">${esc(set.subjudul || 'Buku Kas Rumah Tangga')} • ${esc(lapLabelFilter())}</div></td>
      <td class="d-meta"><table><tr><td>Periode Data:</td><td><b>${tgl(D.a)} – ${tgl(D.b)}</b></td></tr><tr><td>Waktu Cetak:</td><td>${tgl(today())}, ${nowHMS().slice(0, 5)} WIB</td></tr><tr><td>Nomor Dokumen:</td><td class="mono"><b>${no}</b></td></tr></table></td></tr></table>
    <div class="d-rule"></div>
    <div class="d-sec">I. RINGKASAN ARUS KAS MASUK &amp; KELUAR <span class="d-r">Mata Uang: IDR (Rupiah Indonesia)</span></div>
    <table class="d-cards"><tr>
      <td><div class="d-card"><div class="k">TOTAL PEMASUKAN</div><div class="v in">${rp(D.tin)}</div><div class="s">${D.masuk.length} transaksi pemasukan</div></div></td>
      <td><div class="d-card warn"><div class="k">TOTAL PENGELUARAN</div><div class="v out">${rp(D.tout)}</div><div class="s">${D.tin ? pct(D.tout / D.tin, 1) + ' dari total penerimaan' : D.keluar.length + ' transaksi pengeluaran'}</div></div></td>
      <td><div class="d-card ${D.net >= 0 ? 'ok' : 'bad'}"><div class="k">${D.net >= 0 ? 'SURPLUS' : 'DEFISIT'} BERSIH (NETT)</div><div class="v ${D.net >= 0 ? 'in' : 'out'}">${D.net >= 0 ? '+' : ''}${rp(D.net)}</div><div class="s">Tingkat tabungan: ${pct(rate, 1)} ${rate >= 0.3 ? '(Sangat Sehat)' : rate >= 0.1 ? '(Sehat)' : rate >= 0 ? '(Tipis)' : '(Defisit)'}</div></div></td></tr></table>
    <table class="d-two"><tr>
      <td><div class="d-box"><div class="d-bt">PENGELUARAN PER KATEGORI UTAMA <span>Periode terpilih</span></div>
        ${top.length ? top.map((c, i) => `<table class="d-row"><tr><td>${i + 1}. ${esc(c.n)}</td><td class="r">${rp(c.v)} <em>(${pct(c.v / (D.tout || 1), 1)})</em></td></tr></table><div class="d-bar"><i style="width:${((c.v / (top[0].v || 1)) * 100).toFixed(0)}%;background:${['#0F5C3F', '#167A54', '#2DB37A', '#5FD3A0', '#94A3B8', '#BAC6BF'][i] || '#BAC6BF'}"></i></div>`).join('') : '<div style="color:#8BA296">Tidak ada pengeluaran.</div>'}</div></td>
      <td><div class="d-box"><div class="d-bt">POSISI SALDO KAS &amp; BANK AKTIF <span>Per ${tgl(today())}</span></div>
        <table class="d-row">${dompetAktif().map((d) => `<tr><td><span class="d-dot" style="background:${d.warna}"></span>${esc(d.nama)}</td><td class="r">${rp(d.saldo)}</td></tr>`).join('')}
        <tr class="d-tot"><td>Total Likuiditas Keluarga:</td><td class="r" style="color:#0F5C3F;font-size:12px">${rp(totalSaldo())}</td></tr></table></div></td></tr></table>
    <div class="d-sec">II. REKAP TRANSAKSI PERIODE <span class="d-r">${rows.length} transaksi tercatat</span></div>
    <table class="d-tbl"><thead><tr><th style="width:68px">Tanggal</th><th>Deskripsi Transaksi</th><th>Kategori</th><th>Dompet</th><th>Dicatat Oleh</th><th class="r">Jumlah (Rp)</th></tr></thead><tbody>
      ${rows.map((t) => `<tr><td class="mono">${tgl(t.tanggal, 'num')}</td><td>${esc(txTitle(t))}</td><td>${esc(t.tipe === 'transfer' ? 'Transfer Dompet' : S.kMap.get(t.kategori_id)?.nama || '-')}</td><td>${t.tipe === 'transfer' ? esc((S.dMap.get(t.dompet_id)?.nama || '?') + ' → ' + (S.dMap.get(t.dompet_tujuan_id)?.nama || '?')) : esc(S.dMap.get(t.dompet_id)?.nama || '-')}</td><td>${esc(userOf(t.dibuat_oleh).nama)}${userOf(t.dibuat_oleh).label ? ' (' + esc(userOf(t.dibuat_oleh).label) + ')' : ''}</td>
        <td class="r ${t.tipe === 'masuk' ? 'in' : t.tipe === 'keluar' ? 'out' : 'tr'}">${t.tipe === 'masuk' ? '+' : t.tipe === 'keluar' ? '-' : ''}${fmtN(t.nominal)}</td></tr>`).join('') || '<tr><td colspan="6" style="text-align:center;color:#8BA296;padding:16px">Tidak ada transaksi pada periode ini.</td></tr>'}
    </tbody><tfoot><tr><td colspan="5">Total Pemasukan / Pengeluaran / Selisih</td><td class="r">+${fmtN(D.tin)} / -${fmtN(D.tout)} / ${D.net >= 0 ? '+' : ''}${fmtN(D.net)}</td></tr></tfoot></table>
    <table class="d-sign"><tr>
      <td>Mengetahui &amp; Menyetujui${u0.label ? ' (' + esc(u0.label) + ')' : ''},<br><span style="color:#8BA296">Kepala Keluarga</span><div class="d-sig">${esc(u0.nama)}</div><div class="d-line"></div><b>${esc(u0.nama)}</b></td>
      <td>${u1 ? `Pengelola Kas Rumah Tangga${u1.label ? ' (' + esc(u1.label) + ')' : ''},<br><span style="color:#8BA296">Bendahara Keuangan Keluarga</span><div class="d-sig">${esc(u1.nama)}</div><div class="d-line"></div><b>${esc(u1.nama)}</b>` : ''}</td></tr></table>
    <div class="d-foot"><span class="r">Dicetak oleh ${esc(S.me.nama)} • Halaman 1</span>Dokumen ini dicetak otomatis dari ${esc(nama)}. Data bersumber dari Google Sheets keluarga.</div>
  </div>`;
}
let PV = { zoom: 1, D: null };
function ensureDocCss() {
  if ($('#docCss')) return;
  const s = document.createElement('style'); s.id = 'docCss'; s.textContent = DOC_CSS; document.head.appendChild(s);
}
ACT['lap-print'] = () => {
  const D = laporanData();
  PV.D = D;
  ensureDocCss();
  $('#pvPage').innerHTML = docHtml(D);
  $('#pvSub').textContent = `Laporan Rekapitulasi Arus Kas • ${tgl(D.a)} – ${tgl(D.b)}`;
  const pages = Math.max(1, Math.ceil((D.rows.length + 14) / 32));
  $('#pvPages').textContent = `A4 Portrait • ${pages} Halaman`;
  PV.zoom = Math.min(1, (window.innerWidth - 24) / 794);
  pvZoom();
  document.body.classList.add('mode-print');
  $('#printView').scrollTop = 0;
};
function pvZoom() {
  const p = $('#pvPage');
  p.style.transform = `scale(${PV.zoom})`;
  p.style.height = PV.zoom < 1 ? `${p.firstElementChild.offsetHeight * PV.zoom}px` : '';
  $('#pvZoom').textContent = Math.round(PV.zoom * 100) + '%';
}
ACT['pv-zoom'] = (el) => { PV.zoom = clamp(PV.zoom + +el.dataset.z * 0.1, 0.4, 1.6); pvZoom(); };
function closePrint() { document.body.classList.remove('mode-print'); }
ACT['print-close'] = closePrint;
ACT['pv-print'] = () => { const z = PV.zoom; PV.zoom = 1; pvZoom(); setTimeout(() => { window.print(); PV.zoom = z; pvZoom(); }, 60); };
ACT['pv-pdf'] = (btn) => withBtn(btn, async () => {
  const D = PV.D || laporanData();
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>@page{size:A4;margin:10mm}body{margin:0}${DOC_CSS}.doc{box-shadow:none;width:auto;min-height:0;padding:0}</style></head><body>${docHtml(D, true)}</body></html>`;
  try {
    const res = await api('buatPdfLaporan', html, `Laporan-Keuangan-${D.a}-sd-${D.b}`);
    downloadBlob(b64Blob(res.data.base64, 'application/pdf'), res.data.nama);
    toast('PDF diunduh: ' + res.data.nama, 'success', 4000);
  } catch (e) {
    toast('PDF server tidak tersedia — pilih "Simpan sebagai PDF" di dialog cetak.', 'info', 5000);
    ACT['pv-print']();
  }
});


// ════════════════════════════════════════════════════════════
// 19. TEMA DINAMIS (warna dari Pengaturan → seluruh aplikasi)
// ════════════════════════════════════════════════════════════
const HEX = /^#[0-9A-Fa-f]{6}$/;
function applyTheme(set) {
  const p = HEX.test(set?.warna_utama || '') ? set.warna_utama : '#0F5C3F';
  const a = HEX.test(set?.warna_aksen || '') ? set.warna_aksen : '#2DB37A';
  const st = document.documentElement.style, dark = isDark();
  const v = {
    '--primary': p, '--primary-rgb': hexRgb(p).join(', '), '--primary-600': mix(p, '#000000', 0.18), '--primary-700': mix(p, '#000000', 0.34),
    '--primary-glow': mix(p, a, 0.3), '--accent': a, '--accent-rgb': hexRgb(a).join(', '),
    '--on-primary': lum(p) > 0.45 ? '#13231B' : '#FFFFFF',
    '--primary-50': dark ? rgba(a, 0.14) : mix(a, '#FFFFFF', 0.87),
    '--primary-25': dark ? rgba(a, 0.07) : mix(a, '#FFFFFF', 0.94),
    '--accent-50': dark ? rgba(a, 0.14) : mix(a, '#FFFFFF', 0.87),
    '--primary-text': dark ? mix(a, '#FFFFFF', 0.3) : (lum(p) > 0.3 ? mix(p, '#000000', 0.45) : p)
  };
  Object.entries(v).forEach(([k, val]) => st.setProperty(k, val));
  const m = document.querySelector('meta[name="theme-color"]');
  if (m) m.setAttribute('content', p);
}

// ════════════════════════════════════════════════════════════
// 20. HALAMAN PENGATURAN
// ════════════════════════════════════════════════════════════
const SETF = { draft: null, dirty: false, info: null, tab: 'identitas' };
const QUICK_P = ['#0F5C3F', '#1E40AF', '#6D28D9', '#C2410C', '#1F2937'];
const QUICK_A = ['#2DB37A', '#0284C7', '#F59E0B', '#F43F5E', '#10B981'];
PAGES.pengaturan = {
  t: 'Pengaturan', r: renderPengaturan, noAuto: true,
  prep: (o) => {
    SETF.draft = { ...S.settings, emails: recipients() };
    SETF.dirty = false;
    SETF.tab = o.tab || 'identitas';
    api('infoSistem').then((res) => { SETF.info = res.data; if (CUR === 'pengaturan') renderSetInfo(); }).catch(() => {});
    if (o.tab) setTimeout(() => $('#set-' + o.tab)?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 80);
  }
};
function revertLiveTheme() { if (SETF.dirty) { applyTheme(S.settings); SETF.dirty = false; } }
function setDirty() {
  SETF.dirty = true;
  const bar = $('#setBar');
  if (bar) bar.hidden = false;
}
function integritas() {
  const masalah = [];
  const yatim = S.tx.filter((t) => !S.dMap.has(t.dompet_id) || (t.tipe === 'transfer' && !S.dMap.has(t.dompet_tujuan_id)) || (t.tipe !== 'transfer' && !S.kMap.has(t.kategori_id))).length;
  if (yatim) masalah.push(`${yatim} transaksi merujuk dompet/kategori yang tidak ada`);
  const ids = new Set(); let dup = 0;
  S.tx.forEach((t) => { if (ids.has(t.id)) dup++; ids.add(t.id); });
  if (dup) masalah.push(`${dup} ID transaksi ganda`);
  return masalah;
}
function renderPengaturan() {
  const d = SETF.draft || { ...S.settings, emails: recipients() };
  SETF.draft = d;
  const jam = +d.jam_pengingat || 0;
  const defRem = String(d.hari_pengingat_default || '7,3,1,0').split(',').filter(Boolean).map(Number);
  const masalah = integritas();
  const totalBaris = Object.values(S.meta.baris || {}).reduce((s, n) => s + n, 0);
  $('#page-pengaturan').innerHTML = `
  <div class="page-head"><div class="grow"><div class="crumb"><span class="tag">${ic('tune')}Preferensi Rumah Tangga</span></div><h1 class="page-title">Pengaturan Sistem</h1><p class="page-sub">Kelola identitas aplikasi keluarga, integrasi pengingat, hak akses akun, dan cadangan data.</p></div>
    <span class="pill pill-ok pill-lg"><span class="dot"></span>Sinkronisasi Cloud: Aktif</span></div>
  <div class="tabs" role="tablist">${[['identitas', 'palette', 'Identitas Aplikasi'], ['notifikasi', 'notifications_active', 'Notifikasi & Jadwal'], ['pengguna', 'group', 'Pengguna Terdaftar'], ['data', 'database', 'Data & Cadangan']].map(([k, i, l]) => `<button class="${SETF.tab === k ? 'on' : ''}" data-act="set-tab" data-v="${k}">${ic(i)}${l}${k === 'pengguna' ? ` <span class="pill">${S.users.length}</span>` : ''}</button>`).join('')}</div>

  <div class="grid g-side" id="set-identitas">
    <div class="card">
      <div class="card-h"><span class="icon-well">${ic('badge')}</span><div class="grow"><div class="card-t">Identitas Aplikasi</div><div class="card-s">Kustomisasi nama, logo, serta palet warna antarmuka</div></div><span class="pill pill-ok">Berlaku di semua perangkat</span></div>
      <div class="form-grid">
        <div class="field"><label for="setNama">Nama Aplikasi Rumah Tangga</label><div class="input-icon"><span class="ms" style="left:auto;right:13px">edit</span><input class="input" id="setNama" maxlength="40" value="${esc(d.nama_aplikasi || '')}" data-in="set-f" data-k="nama_aplikasi" style="padding-left:14px"></div><div class="hint">Tampil di sidebar, dokumen laporan, dan email pengingat.</div></div>
        <div class="field"><label for="setSub">Subjudul / Nama Keluarga</label><input class="input" id="setSub" maxlength="60" value="${esc(d.subjudul || '')}" data-in="set-f" data-k="subjudul" placeholder="Keluarga ${esc(familyName())}"></div>
      </div>
      <div class="field mt-3"><span class="lbl">Logo Aplikasi</span>
        <div class="logo-drop"><button class="logo-prev" data-act="logo-preview" aria-label="Pratinjau logo">${d.logo_url ? `<img src="${esc(d.logo_url)}" alt="Logo" referrerpolicy="no-referrer">` : ic('eco')}</button>
          <div class="grow" style="min-width:200px"><div class="row" style="flex-wrap:wrap;gap:8px"><label class="btn btn-soft btn-sm" id="logoBtn">${ic('upload')}Ganti Gambar<input type="file" accept="image/png,image/jpeg,image/webp" hidden data-ch="logo-file"></label>${d.logo_url ? `<button class="btn btn-danger-soft btn-sm" data-act="logo-del">${ic('delete')}Hapus Logo</button>` : ''}</div>
            <div class="hint mt-1">PNG/JPG otomatis diperkecil maks. 256×256 px lalu disimpan di folder Google Drive keluarga.</div></div></div></div>
      <div class="field mt-3"><span class="lbl">Penyesuaian Warna Tema</span>
        <div class="grid g-2" style="gap:14px">
          ${[['warna_utama', 'Warna Utama (Primary Surface)', QUICK_P], ['warna_aksen', 'Warna Aksen (Indikator & Highlight)', QUICK_A]].map(([k, l, q]) => `<div class="mini-box"><div class="small semi mb-2">${l}</div>
            <div class="color-field"><span class="sw" style="background:${esc(d[k])}"><input type="color" value="${esc(d[k])}" data-in="set-color" data-k="${k}" aria-label="${l}"></span><input class="input" value="${esc(String(d[k]).toUpperCase())}" maxlength="7" data-in="set-hex" data-k="${k}" aria-label="Kode warna ${l}"></div>
            <div class="xs faint mt-2 mb-2">Pilihan Cepat:</div><div class="swatches">${q.map((h) => `<button type="button" class="swatch ${String(d[k]).toUpperCase() === h ? 'on' : ''}" style="background:${h}" data-act="set-quick" data-k="${k}" data-v="${h}" aria-label="Warna ${h}"></button>`).join('')}</div></div>`).join('')}
        </div></div>
    </div>
    <div class="card">
      <div class="card-h"><div class="grow"><div class="card-t">${ic('preview')}Mini Live Preview</div><div class="card-s">Representasi antarmuka dengan tema terpilih</div></div><span class="pill pill-solid">Interaktif</span></div>
      <div class="preview-app"><div class="ph"><span class="brand-logo" style="width:26px;height:26px;border-radius:8px">${d.logo_url ? `<img src="${esc(d.logo_url)}" alt="" referrerpolicy="no-referrer">` : ic('eco', 'sm')}</span>${esc(d.nama_aplikasi || 'Keuangan Digital')}<span class="grow"></span><span class="xs faint">${esc(familyName())}</span></div>
        <div class="phero"><div class="xs" style="opacity:.8">Total Saldo Keluarga</div><div class="money" style="font:800 20px/1.3 var(--font-head)">${rp(totalSaldo())}</div><div class="row xs" style="justify-content:space-between;margin-top:8px"><span>${dompetAktif().length} Rekening Aktif</span><span class="pill" style="background:var(--accent);color:#fff;height:20px">+${pct(Math.max(0, health().rate))} tabungan</span></div></div>
        <div class="row small mt-2" style="justify-content:space-between"><span class="faint">Komponen Aksen:</span><span class="pill pill-ok">Lunas / On Track</span></div>
        <button class="btn btn-primary btn-block btn-sm mt-2" type="button">${ic('add_circle')}Catat Transaksi Cepat</button></div>
      <div class="note ok mt-3">${ic('palette')}<div class="small">Warna diterapkan konsisten pada grafik belanja, pill status, dan tombol aksi utama — langsung terlihat di seluruh aplikasi.</div></div>
    </div>
  </div>

  <div class="card mt-3" id="set-notifikasi">
    <div class="card-h"><span class="icon-well">${ic('notifications_active')}</span><div class="grow"><div class="card-t">Notifikasi &amp; Jadwal Pengingat</div><div class="card-s">Penerima email pengingat tagihan dan jam pengiriman harian</div></div></div>
    <div class="grid g-2">
      <div class="field"><span class="lbl">Alamat Email Penerima Pengingat</span>
        <div class="email-chips" id="setEmails">${d.emails.map((e, i) => `<span class="ec">${ic('account_circle', 'sm')}${esc(e)}<button type="button" data-act="set-email-del" data-i="${i}" aria-label="Hapus ${esc(e)}">${ic('close', 'sm')}</button></span>`).join('')}<input id="setEmailIn" type="email" placeholder="+ Tambah email…" data-enter="set-email-add" aria-label="Tambah email"></div>
        <div class="hint">Tekan Enter untuk menambah. Pengingat tagihan (H-x &amp; hari-H) dikirim ke semua alamat ini.</div></div>
      <div class="field"><label for="setJam">Waktu Pengiriman Pengingat Harian</label>
        <select class="select" id="setJam" data-ch="set-jam">${Array.from({ length: 17 }, (_, i) => i + 5).map((h) => `<option value="${h}" ${h === jam ? 'selected' : ''}>Pukul ${String(h).padStart(2, '0')}:00 WIB${h === 7 ? ' (Pagi — sebelum aktivitas)' : ''}</option>`).join('')}</select>
        <div class="note ok mt-2" id="setQuota">${ic('verified_user')}<div class="small">Memeriksa kuota email Gmail…</div></div></div>
    </div>
    <div class="field mt-3"><span class="lbl">Jadwal Default untuk Tagihan Baru</span><div class="chips">${REM.map(([h, l]) => `<button type="button" class="chip toggle ${defRem.includes(h) ? 'on' : ''}" data-act="set-rem" data-v="${h}">${defRem.includes(h) ? ic('check') : ''}${l}</button>`).join('')}</div></div>
    <div class="row mt-3" style="flex-wrap:wrap"><button class="btn btn-outline btn-sm" data-act="set-test" data-loading="Mengirim…">${ic('send')}Kirim Email Uji Coba</button><span class="hint">Mengirim contoh pengingat ke penerima yang <b>sudah tersimpan</b>.</span></div>
  </div>

  <div class="card mt-3" id="set-pengguna">
    <div class="card-h"><span class="icon-well">${ic('shield_person')}</span><div class="grow"><div class="card-t">Pengguna Terdaftar (Akses Google Whitelist)</div><div class="card-s">Otoritas akun co-management pasangan rumah tangga</div></div><span class="pill pill-warn">${ic('lock')}Maksimum 2 Akun Google</span></div>
    <div class="grid g-2">${S.users.map((u, i) => `<div class="user-card">${avatar(u.email, 'lg')}<div class="grow" style="min-width:0"><div class="semi">${esc(u.nama)}${u.email === S.me.email ? '<span class="online" title="Anda"></span>' : ''}</div><div class="small faint ellipsis">${esc(u.email)}</div>
      <div class="row mt-1" style="gap:8px;flex-wrap:wrap"><span class="pill pill-ok">Pemilik (Hak Akses Penuh)</span><span class="xs faint">${esc(u.label || (i ? 'Istri' : 'Suami'))}</span></div></div><button class="btn-icon" data-act="users-edit" aria-label="Ubah pengguna">${ic('more_vert')}</button></div>`).join('')}
      ${S.users.length < 2 ? `<button class="add-card" style="min-height:96px" data-act="users-edit">${ic('person_add', 'lg')}<b>Tambah Pasangan</b></button>` : ''}</div>
    <div class="note danger mt-3">${ic('gpp_maybe')}<div><b class="small">Catatan Keamanan Berlapis</b><div class="small">Aplikasi ini privat dan hanya untuk akun di atas. Akun Google lain langsung ditolak di halaman "Akses Tidak Diizinkan". Setiap transaksi mencatat siapa penginputnya.</div></div></div>
  </div>

  <div class="card mt-3" id="set-data">
    <div class="card-h" style="flex-wrap:wrap"><span class="icon-well">${ic('database')}</span><div class="grow"><div class="card-t">Data &amp; Cadangan Spreadsheet</div><div class="card-s">Penyimpanan cloud Google Sheets dan arsip riwayat transaksi</div></div>
      <button class="btn btn-outline btn-sm" data-act="set-backup" data-loading="Membuat salinan…">${ic('backup')}Buat Salinan Cadangan Sekarang</button></div>
    <div class="grid g-3" style="gap:14px">
      <div class="stat-box"><div class="k">${ic('table_view')}Status Database</div><div class="v">${esc(S.settings.nama_aplikasi || 'Keuangan Digital')}</div><div class="s">Terhubung ke Google Sheets (${fmtN(totalBaris)} baris data; ${fmtN(S.meta.totalTx || S.tx.length)} transaksi). ${fmtN(S.tx.length)} transaksi termuat di perangkat ini.</div></div>
      <div class="stat-box"><div class="k">${ic('history')}Cadangan Terakhir</div><div class="v" id="setBackupAt">${esc(S.meta.cadanganTerakhir || 'Belum ada')}</div><div class="s">Cadangan bulanan otomatis ke folder Drive <span class="mono">/Keuangan Digital/Cadangan/</span> (12 terakhir disimpan).</div></div>
      <div class="stat-box"><div class="k">${ic(masalah.length ? 'report' : 'verified')}Integritas Data</div><div class="v ${masalah.length ? 'neg' : ''}">${masalah.length ? 'Perlu Dicek' : '100% Sehat &amp; Tersinkron'}</div><div class="s">${masalah.length ? esc(masalah.join('; ')) : 'Tidak ada transaksi yatim atau ID ganda yang terdeteksi.'}</div></div>
    </div>
    <div class="row mt-3" style="flex-wrap:wrap;gap:8px">
      ${S.meta.ssUrl ? `<a class="btn btn-soft btn-sm" href="${esc(S.meta.ssUrl)}" target="_blank" rel="noopener">${ic('open_in_new')}Buka Spreadsheet</a>` : ''}
      ${S.meta.folderUrl ? `<a class="btn btn-soft btn-sm" href="${esc(S.meta.folderUrl)}" target="_blank" rel="noopener">${ic('folder_open')}Buka Folder Drive</a>` : ''}
      <button class="btn btn-ghost btn-sm" data-act="reload">${ic('sync')}Muat Ulang Data</button>
      <button class="btn btn-ghost btn-sm" data-act="set-cache" data-loading="Membersihkan…">${ic('cleaning_services')}Bersihkan Cache Server</button>
      ${S.meta.adaContoh ? `<button class="btn btn-danger-soft btn-sm" data-act="set-demo" data-loading="Menghapus…">${ic('delete_sweep')}Hapus Data Contoh</button>` : ''}
    </div>
    ${S.meta.adaContoh ? `<div class="note mt-3">${ic('info')}<div class="small">Aplikasi masih berisi <b>data contoh</b>. Setelah mencoba, klik "Hapus Data Contoh", lalu tambahkan dompet asli beserta saldo awalnya.</div></div>` : ''}
    <div class="note info mt-3">${ic('install_mobile')}<div class="small"><b>Tips:</b> buka aplikasi di Chrome HP → menu ⋮ → <b>Tambahkan ke layar utama</b> agar terasa seperti aplikasi.</div></div>
  </div>

  <div class="sticky-bar" id="setBar" ${SETF.dirty ? '' : 'hidden'}><span class="status">${ic('info', 'sm')}Ada perubahan yang belum disimpan.</span>
    <button class="btn btn-outline" data-act="set-reset">${ic('restart_alt')}Reset ke Default</button><button class="btn btn-ghost" data-act="set-cancel">Batal</button><button class="btn btn-primary" data-act="set-save">${ic('save')}Simpan Perubahan Pengaturan</button></div>`;
  renderSetInfo();
}
function renderSetInfo() {
  const i = SETF.info, q = $('#setQuota');
  if (!q) return;
  if (!i) return;
  q.innerHTML = `${ic(i.kuotaEmail > 10 ? 'verified_user' : 'warning')}<div class="small">Gmail (MailApp) aktif${i.kuotaEmail != null ? `, sisa kuota hari ini <b>${i.kuotaEmail}</b> penerima` : ''}. ${i.pemicu ? `Pemicu otomatis dibuat oleh ${esc(i.pemicu.oleh)}.` : '<b>Pemicu belum dibuat — jalankan setupAppEnvironment().</b>'}</div>`;
  q.className = 'note mt-2 ' + (i.pemicu ? 'ok' : '');
  if (i.cadanganTerakhir && $('#setBackupAt')) $('#setBackupAt').textContent = i.cadanganTerakhir;
}
ACT['set-tab'] = (el) => { SETF.tab = el.dataset.v; $$('.tabs button').forEach((b) => b.classList.toggle('on', b === el)); $('#set-' + el.dataset.v)?.scrollIntoView({ behavior: 'smooth', block: 'start' }); };
ON['set-f'] = (el) => { SETF.draft[el.dataset.k] = el.value; setDirty(); };
function setColor(k, v) {
  if (!HEX.test(v)) return;
  SETF.draft[k] = v.toUpperCase();
  applyTheme(SETF.draft);
  setDirty();
}
ON['set-color'] = (el) => { setColor(el.dataset.k, el.value); const box = el.closest('.color-field'); box.querySelector('.input').value = el.value.toUpperCase(); box.querySelector('.sw').style.background = el.value; };
ON['set-hex'] = (el) => { let v = el.value.trim(); if (!v.startsWith('#')) v = '#' + v; if (HEX.test(v)) { setColor(el.dataset.k, v); const box = el.closest('.color-field'); box.querySelector('.sw').style.background = v; box.querySelector('input[type=color]').value = v; } };
ACT['set-quick'] = (el) => { setColor(el.dataset.k, el.dataset.v); renderPengaturan(); };
CHG['set-jam'] = (el) => { SETF.draft.jam_pengingat = +el.value; setDirty(); };
ACT['set-rem'] = (el) => {
  const v = +el.dataset.v;
  let L = String(SETF.draft.hari_pengingat_default || '7,3,1,0').split(',').filter(Boolean).map(Number);
  L = L.includes(v) ? L.filter((x) => x !== v) : L.concat(v);
  SETF.draft.hari_pengingat_default = L.sort((a, b) => b - a).join(',');
  setDirty(); renderPengaturan();
};
ACT['set-email-add'] = () => {
  const i = $('#setEmailIn'), v = i.value.trim().toLowerCase().replace(/[,;]$/, '');
  if (!v) return;
  if (!RE_EMAIL.test(v)) { toast('Format email tidak valid.', 'warn'); return; }
  if (SETF.draft.emails.includes(v)) { i.value = ''; return; }
  if (SETF.draft.emails.length >= 10) { toast('Maksimal 10 email.', 'warn'); return; }
  SETF.draft.emails.push(v); setDirty(); renderPengaturan(); $('#setEmailIn').focus();
};
ACT['set-email-del'] = (el) => { SETF.draft.emails.splice(+el.dataset.i, 1); setDirty(); renderPengaturan(); };
CHG['logo-file'] = async (el) => {
  const f = el.files[0];
  if (!f) return;
  const btn = $('#logoBtn');
  try {
    const data = await resizeImage(f, 256);
    SETF.draft.logo_url = data; renderPengaturan(); setDirty();
    await withBtn($('#logoBtn'), async () => {
      const res = await api('unggahGambar', data, 'logo');
      SETF.draft.logo_url = res.data.url;
      toast('Logo diunggah ke Google Drive. Klik "Simpan Perubahan" untuk menerapkan.', 'success', 4500);
    });
    renderPengaturan(); setDirty();
  } catch (e) { toast('Logo gagal diunggah: ' + e.message, 'error', 5000); SETF.draft.logo_url = S.settings.logo_url || ''; renderPengaturan(); }
  void btn;
};
ACT['logo-del'] = () => { SETF.draft.logo_url = ''; setDirty(); renderPengaturan(); };
ACT['logo-preview'] = () => {
  const url = SETF.draft?.logo_url || S.settings.logo_url;
  if (!url) { toast('Belum ada logo.', 'info'); return; }
  openModal({
    title: 'Pratinjau Logo', icon: 'image', size: 'sm',
    body: `<div style="display:grid;place-items:center;padding:12px;background:var(--surface-2);border-radius:14px"><img src="${esc(url)}" alt="Logo aplikasi" style="max-width:256px;max-height:256px;border-radius:16px" referrerpolicy="no-referrer"></div>`,
    foot: `<button class="btn btn-ghost" data-act="modal-close">${ic('close')}Tutup</button><div class="grow"></div>${url.startsWith('data:') ? `<a class="btn btn-primary" href="${esc(url)}" download="logo.png">${ic('download')}Unduh</a>` : `<a class="btn btn-primary" href="${esc(url)}" target="_blank" rel="noopener" download>${ic('download')}Unduh</a>`}`
  });
};
ACT['set-reset'] = () => {
  Object.assign(SETF.draft, { nama_aplikasi: 'Keuangan Digital', warna_utama: '#0F5C3F', warna_aksen: '#2DB37A', jam_pengingat: 7, hari_pengingat_default: '7,3,1,0' });
  applyTheme(SETF.draft); setDirty(); renderPengaturan(); toast('Nilai default dimuat. Klik Simpan untuk menerapkan.', 'info');
};
ACT['set-cancel'] = () => { SETF.draft = { ...S.settings, emails: recipients() }; SETF.dirty = false; applyTheme(S.settings); renderPengaturan(); };
ACT['set-save'] = () => {
  const d = SETF.draft;
  if (!String(d.nama_aplikasi || '').trim()) { toast('Nama aplikasi wajib diisi.', 'warn'); return; }
  if (String(d.logo_url || '').startsWith('data:')) { toast('Tunggu logo selesai diunggah.', 'warn'); return; }
  if (!d.emails.length) { toast('Minimal satu email penerima pengingat.', 'warn'); return; }
  DB.simpanPengaturan({ nama_aplikasi: d.nama_aplikasi.trim(), subjudul: String(d.subjudul || '').trim(), logo_url: d.logo_url || '', warna_utama: d.warna_utama, warna_aksen: d.warna_aksen, email_pengingat: d.emails.join(','), jam_pengingat: +d.jam_pengingat, hari_pengingat_default: d.hari_pengingat_default });
  SETF.dirty = false;
  SETF.draft = { ...S.settings, emails: recipients() };
  renderPengaturan();
  toast('Pengaturan disimpan & diterapkan ke seluruh aplikasi.', 'success');
};
ACT['set-test'] = (el) => withBtn(el, () => api('kirimEmailUji', { nama: 'Contoh: Listrik PLN', nominal: 650000, tanggal: addDays(today(), 3), hari: 3 })
  .then((r) => toast(r.message, 'success', 5000)).catch((e) => toast('Email gagal: ' + e.message, 'error', 6000)));
ACT['set-backup'] = (el) => withBtn(el, () => api('buatCadanganSekarang').then((r) => { S.meta.cadanganTerakhir = r.data.waktu; if ($('#setBackupAt')) $('#setBackupAt').textContent = r.data.waktu; toast(r.message, 'success', 5000); }).catch((e) => toast('Cadangan gagal: ' + e.message, 'error', 6000)));
ACT['set-cache'] = (el) => withBtn(el, () => api('bersihkanCache').then(() => { toast('Cache server dibersihkan.', 'success'); refreshData({ silent: true }); }).catch((e) => toast(e.message, 'error')));
ACT['set-demo'] = async (el) => {
  if (!(await confirmBox({ title: 'Hapus Data Contoh?', text: 'Semua transaksi, budget, target, dan tagihan <b>contoh</b> (berawalan demo_) akan dihapus permanen. Data yang Anda input sendiri tetap aman.', ok: 'Hapus Data Contoh', danger: true }))) return;
  const btn = $('[data-act="set-demo"]');
  withBtn(btn, () => new Promise((ok) => {
    const go = () => api('hapusDataContoh').then((r) => { toast(r.message, 'success', 6000); refreshData({ silent: true, done: () => { SETF.draft = null; if (CUR === 'pengaturan') renderPengaturan(); } }); }).catch((e) => toast(e.message, 'error')).finally(ok);
    if (SYNC.outbox.length) { flushOutbox(); setTimeout(go, 2500); } else go();
  }));
  void el;
};
ACT['users-edit'] = () => {
  const list = S.users.length ? S.users.map((u) => ({ ...u })) : [{ ...S.me }];
  while (list.length < 2) list.push({ email: '', nama: '', label: list.length ? 'Istri' : 'Suami' });
  openModal({
    title: 'Pengguna Terdaftar', icon: 'group',
    body: `<div class="note mb-3">${ic('warning')}<div class="small">Hanya 2 akun Google ini yang dapat membuka aplikasi. Akun baru otomatis diberi akses Editor ke spreadsheet &amp; folder Drive. Email Anda sendiri wajib tetap terdaftar.</div></div>
      ${list.map((u, i) => `<div class="mini-box ${i ? 'mt-2' : ''}"><div class="semi small mb-2">Anggota ${i + 1}${u.email === S.me.email ? ' (Anda)' : ''}</div><div class="form-grid">
        <div class="field full"><label for="ue${i}">Email Google${i === 0 ? ' <span class="req">*</span>' : ''}</label><input class="input" id="ue${i}" type="email" value="${esc(u.email)}" placeholder="nama@gmail.com" ${u.email === S.me.email ? 'readonly' : ''}></div>
        <div class="field"><label for="un${i}">Nama Panggilan</label><input class="input" id="un${i}" maxlength="40" value="${esc(u.nama)}"></div>
        <div class="field"><label for="ul${i}">Peran Keluarga</label><input class="input" id="ul${i}" maxlength="20" value="${esc(u.label || '')}" placeholder="Suami / Istri"></div></div></div>`).join('')}`,
    foot: '<div class="grow"></div><button class="btn btn-ghost" data-act="modal-close">Batal</button><button class="btn btn-primary" data-act="users-save">' + ic('check') + 'Simpan Pengguna</button>'
  });
};
ACT['users-save'] = async () => {
  const list = [0, 1].map((i) => ({ email: $('#ue' + i).value.trim().toLowerCase(), nama: $('#un' + i).value.trim(), label: $('#ul' + i).value.trim() })).filter((u) => u.email);
  for (const u of list) if (!RE_EMAIL.test(u.email)) { toast('Email tidak valid: ' + u.email, 'warn'); return; }
  if (!list.some((u) => u.email === S.me.email)) { toast('Email Anda sendiri harus tetap terdaftar.', 'warn'); return; }
  if (list.length === 2 && list[0].email === list[1].email) { toast('Dua email tidak boleh sama.', 'warn'); return; }
  list.forEach((u) => { if (!u.nama) u.nama = u.email.split('@')[0]; });
  const baru = list.filter((u) => !S.users.some((x) => x.email === u.email));
  if (baru.length && !(await confirmBox({ title: 'Beri Akses?', text: `Akun <b>${baru.map((u) => esc(u.email)).join(', ')}</b> akan bisa melihat &amp; mengubah seluruh data keuangan keluarga. Lanjutkan?`, ok: 'Ya, Beri Akses', icon: 'shield_person' }))) return;
  DB.simpanPengguna(list);
  closeModal();
  SETF.draft = null;
  renderPengaturan();
  toast('Pengguna terdaftar disimpan.', 'success');
};



// ════════════════════════════════════════════════════════════
// 21. LOGIN GOOGLE, LAYAR GERBANG & INISIALISASI
// ════════════════════════════════════════════════════════════
const GSI_URL = 'https://accounts.google.com/gsi/client';
let APP_STARTED = false;
let PUBLIK = LS.get('kd_public', null);

const bootKey = (email) => 'kd_boot_' + String(email || '').toLowerCase();
/** Simpan salinan data terakhir agar aplikasi tampil instan saat dibuka lagi. */
function simpanCacheBoot(b) {
  if (!b || !b.me) return;
  try { localStorage.setItem(bootKey(b.me.email), JSON.stringify(b)); }
  catch (e) { try { localStorage.removeItem(bootKey(b.me.email)); } catch (e2) { /* abaikan */ } }
}

/** Branding layar gerbang dari info publik (nama, warna). */
function terapkanPublik(info) {
  if (!info) return;
  PUBLIK = info;
  LS.set('kd_public', info);
  $('#gateNama').textContent = info.nama || 'Keuangan Digital';
  $('#gateSub').textContent = info.subjudul || 'Aplikasi keuangan keluarga';
  if (!APP_STARTED) applyTheme(info);
  document.title = info.nama || 'Keuangan Digital';
}

function tampilkanGerbang() { $('#app').hidden = true; $('#gate').hidden = false; closeModal(true); closePopover(); document.body.classList.remove('mode-print'); }

/** Render isi kartu gerbang sesuai keadaan. */
function renderGate(mode, d = {}) {
  tampilkanGerbang();
  const nama = (PUBLIK && PUBLIK.nama) || 'Keuangan Digital';
  const anggota = (PUBLIK && PUBLIK.anggota) || [];
  const card = $('#gateCard');
  const spin = (t) => `<div style="padding:30px 0"><div class="spinner" style="width:34px;height:34px;border-width:3px;color:var(--primary)"></div><p class="muted mt-3">${esc(t)}</p></div>`;
  if (mode === 'muat') card.innerHTML = spin('Memuat aplikasi…');
  else if (mode === 'verifikasi') card.innerHTML = spin('Memverifikasi akun Google…');
  else if (mode === 'config') {
    card.innerHTML = `<div class="gate-lock" style="background:var(--gold-50);color:var(--gold-700)">${ic('build')}</div><h1>Konfigurasi Belum Lengkap</h1>
      <p>Buka file <b>js/config.js</b> di repository GitHub Anda, lalu isi:</p>
      <div class="stack" style="text-align:left;gap:8px;margin-bottom:8px">
        <div class="note ${CFG_OK ? 'ok' : 'danger'}">${ic(CFG_OK ? 'check_circle' : 'error')}<div class="small"><b>GAS_URL</b> — URL Web App Apps Script yang berakhiran <span class="mono">/exec</span></div></div>
        <div class="note ${CLIENT_OK ? 'ok' : 'danger'}">${ic(CLIENT_OK ? 'check_circle' : 'error')}<div class="small"><b>GOOGLE_CLIENT_ID</b> — OAuth Client ID dari Google Cloud (…apps.googleusercontent.com)</div></div>
      </div><div class="gate-foot">Setelah disimpan &amp; di-push, tunggu 1–2 menit lalu tekan Ctrl+Shift+R.</div>`;
  } else if (mode === 'ditolak') {
    const ang = d.anggota || [];
    card.innerHTML = `<div class="gate-lock">${ic('lock')}</div><h1>Akses Tidak Diizinkan</h1>
      <p>Aplikasi ${esc(nama)} ini bersifat privat dan dikhususkan untuk ${ang.length} anggota keluarga terdaftar${ang.length ? ` (${esc(ang.map((a) => a.nama).join(' & '))})` : ''}. Akun Google yang sedang aktif tidak memiliki izin akses.</p>
      <div class="gate-acc"><span class="g">G</span><div class="grow"><div class="label-xs">Akun Terhubung:</div><div class="semi">${esc(d.email || '')}</div></div><span class="pill pill-danger"><span class="dot"></span>Ditolak / Tidak Terdaftar</span></div>
      <div class="gate-members">${ic('verified_user')}<div class="grow"><div class="semi small">Anggota Sah Terdaftar</div><div class="xs muted">${esc(ang.map((a) => a.email).join(' • '))}</div></div><span class="pill pill-ok">${ang.length}/2 Kuota</span></div>
      <div class="stack"><button class="btn btn-primary btn-lg btn-block" data-act="gate-other">${ic('switch_account')}Masuk dengan Akun Google Lain</button>
        <button class="btn btn-outline btn-lg btn-block" data-act="gate-reload">${ic('refresh')}Muat Ulang Halaman</button></div>
      <div class="gate-foot">Jika Anda merasa ini adalah kekeliruan, silakan hubungi ${ang.map((a) => `<b>${esc(a.nama)}</b>`).join(' atau ') || 'pemilik aplikasi'} untuk menambahkan email Anda ke daftar whitelist.</div>`;
  } else if (mode === 'galat') {
    card.innerHTML = `<div class="gate-lock">${ic('cloud_off')}</div><h1>Gagal Terhubung</h1><p>${esc(d.pesan || 'Server tidak dapat dihubungi.')}</p>
      <button class="btn btn-primary btn-lg btn-block" data-act="gate-retry">${ic('refresh')}Coba Lagi</button>`;
  } else {
    card.innerHTML = `<div class="gate-lock" style="background:var(--primary-50);color:var(--primary-text)">${ic('account_balance_wallet')}</div>
      <h1>Masuk ke ${esc(nama)}</h1>
      <p>Aplikasi privat keuangan keluarga${anggota.length ? ' ' + esc(anggota.join(' & ')) : ''}. Masuk dengan akun Google yang terdaftar.</p>
      ${d.pesan ? `<div class="note ${d.tipe === 'info' ? 'info' : 'danger'} mb-3" style="text-align:left">${ic(d.tipe === 'info' ? 'info' : 'error')}<div class="small">${esc(d.pesan)}</div></div>` : ''}
      <div id="gsiBtn" style="display:flex;justify-content:center;min-height:44px">${spin('Menyiapkan tombol Google…').replace('padding:30px 0', 'padding:4px 0')}</div>
      <div class="gate-foot">${ic('lock', 'sm')} Hanya akun Google yang terdaftar di whitelist yang dapat masuk. Login diverifikasi langsung oleh server.</div>`;
    siapkanTombolGoogle();
  }
}

/** Muat Google Identity Services lalu tampilkan tombol "Masuk dengan Google". */
function siapkanTombolGoogle() {
  loadScript(GSI_URL).then(() => {
    const el = $('#gsiBtn');
    if (!el || !window.google || !google.accounts || !google.accounts.id) throw new Error('GIS');
    google.accounts.id.initialize({ client_id: CFG.GOOGLE_CLIENT_ID, callback: onGoogleCredential, auto_select: true, cancel_on_tap_outside: true, itp_support: true, use_fedcm_for_prompt: true });
    el.innerHTML = '';
    google.accounts.id.renderButton(el, { type: 'standard', theme: 'outline', size: 'large', shape: 'pill', text: 'signin_with', logo_alignment: 'left', locale: 'id', width: Math.min(320, el.clientWidth || 320) });
    if (!LS.get('kd_no_auto', false)) google.accounts.id.prompt();
  }).catch(() => {
    const el = $('#gsiBtn');
    if (el) el.innerHTML = `<div class="note danger" style="text-align:left">${ic('error')}<div class="small">Layanan login Google gagal dimuat. Periksa koneksi internet atau pemblokir iklan, lalu muat ulang.</div></div>`;
  });
}

/** Callback GIS: tukar ID token Google dengan sesi aplikasi. */
async function onGoogleCredential(resp) {
  renderGate('verifikasi');
  try {
    const res = await apiRaw({ action: 'login', idToken: resp.credential });
    if (res.success) {
      LS.del('kd_no_auto');
      AUTH.set(res.data.token, res.data.user);
      simpanCacheBoot(res.data.boot);
      startApp(res.data.boot);
      toast(res.message || 'Berhasil masuk.', 'success');
    } else if (res.code === 'DITOLAK') {
      LS.set('kd_no_auto', true);
      renderGate('ditolak', res.data || {});
    } else renderGate('login', { pesan: res.message });
  } catch (err) {
    renderGate('login', { pesan: err.message });
  }
}

/** Dipanggil lapisan API saat server menolak sesi (habis/dicabut). */
function onSesiBerakhir() {
  if (!AUTH.token && !APP_STARTED) return;
  AUTH.clear();
  renderGate('login', { pesan: 'Sesi Anda berakhir. Silakan masuk lagi — perubahan yang belum terkirim tetap aman di perangkat ini.', tipe: 'info' });
}

/** Tampilkan aplikasi dari data boot. */
function startApp(boot) {
  OB_KEY = 'kd_outbox_' + String(boot.me.email).toLowerCase();
  loadBoot(boot);
  applyTheme(S.settings);
  SYNC.outbox = LS.get(OB_KEY, []);
  $('#gate').hidden = true;
  $('#app').hidden = false;
  renderShell();
  goPage(CUR || 'dashboard');
  if (SYNC.outbox.length) {
    SYNC.refreshAfter = true;
    toast(`Mengirim ${SYNC.outbox.length} perubahan yang tertunda…`, 'info', 4000);
    scheduleFlush(400);
  }
  if (APP_STARTED) return;
  APP_STARTED = true;
  setInterval(pollVersi, 45000);
  let hari = today();
  setInterval(() => { if (today() !== hari) { hari = today(); bump(); rerender(); } }, 60000);
  document.addEventListener('visibilitychange', () => { if (!document.hidden) { pollVersi(); if (SYNC.outbox.length) scheduleFlush(100); } });
  window.addEventListener('online', () => { SYNC.retry = 0; scheduleFlush(100); });
  window.addEventListener('beforeunload', (e) => { if (SYNC.outbox.length) { flushOutbox(); e.preventDefault(); e.returnValue = ''; } });
  if (window.matchMedia) matchMedia('(prefers-color-scheme: dark)').addEventListener?.('change', (ev) => { if (!localStorage.getItem('kd_theme')) { document.documentElement.dataset.theme = ev.matches ? 'dark' : 'light'; applyTheme(S.settings); rerender(true); } });
}

/** Keluar: akhiri sesi server + hapus data lokal. */
ACT['do-logout'] = (el) => {
  const token = AUTH.token, email = S.me && S.me.email;
  closeModal(true);
  if (token) apiRaw({ action: 'logout', token }).catch(() => {});
  AUTH.clear();
  if (email) LS.del(bootKey(email));
  LS.set('kd_no_auto', true);
  try { if (window.google && google.accounts && google.accounts.id) google.accounts.id.disableAutoSelect(); } catch (e) { /* abaikan */ }
  $$('.page').forEach((p) => (p.innerHTML = ''));
  Object.keys(CHARTS).forEach((id) => { CHARTS[id].destroy(); delete CHARTS[id]; });
  renderGate('login', { pesan: el.dataset.switch ? 'Pilih akun Google lain untuk masuk.' : 'Anda sudah keluar.', tipe: 'info' });
};
ACT['gate-other'] = () => { try { google.accounts.id.disableAutoSelect(); } catch (e) { /* abaikan */ } renderGate('login'); };
ACT['gate-reload'] = () => location.reload();
ACT['gate-retry'] = () => initApp();
ACT['sync-now'] = () => { SYNC.retry = 0; toast('Mencoba mengirim ulang…', 'info', 1500); scheduleFlush(50); };

/** Titik awal aplikasi. */
async function initApp() {
  if (PUBLIK) terapkanPublik(PUBLIK);
  apiInfo().then(terapkanPublik);
  if (!CFG_OK || !CLIENT_OK) { renderGate('config'); return; }
  const user = AUTH.user;
  if (!AUTH.token || !user) { renderGate('login'); return; }
  // 1) tampil instan dari cache perangkat, 2) perbarui dari server di latar belakang
  const cached = LS.get(bootKey(user.email), null);
  if (cached && cached.me) {
    startApp(cached);
    if (SYNC.outbox.length) SYNC.refreshAfter = true; else refreshData({ silent: true });
    return;
  }
  renderGate('muat');
  try {
    const res = await api('ambilDataAwal');
    simpanCacheBoot(res.data);
    startApp(res.data);
  } catch (err) {
    if (err.code === 'AUTH') return; // onSesiBerakhir sudah menampilkan login
    renderGate('galat', { pesan: err.message });
  }
}
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initApp); else initApp();
