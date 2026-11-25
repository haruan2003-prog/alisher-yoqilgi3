/* ============================================================
   ALISHER XUSUSIY KORXONA — Yoqilg‘i hisob-kitoblari tizimi
   Barcha raqamlar /api/sheets orqali Google Sheets'dan olinadi.
   ============================================================ */
'use strict';

const CFG = {
  api: '/api/sheets',
  company: 'ALISHER XUSUSIY KORXONA',
  EPS: 0.5,                      // 0,5 litrdan kichik qoldiq jadvalda ham "0" ko‘rinadi
  cacheKey: 'axk-cache-v1',
  passKey: 'axk-pass',
  libs: {
    xlsx: 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
    pdf: 'https://cdn.jsdelivr.net/npm/html2pdf.js@0.10.2/dist/html2pdf.bundle.min.js'
  }
};

const UZ_MONTHS = ['Yanvar', 'Fevral', 'Mart', 'Aprel', 'May', 'Iyun', 'Iyul', 'Avgust', 'Sentabr', 'Oktabr', 'Noyabr', 'Dekabr'];
const BRANCH_COLORS = ['#6aa9ff', '#4fd1a5', '#ffc46b', '#c792ff', '#67d4e8', '#ff8fa3'];

/* ---------------- Ikonkalar (Lucide uslubidagi chiziqli SVG) ---------------- */
const ICONS = {
  home: '<path d="M3 10.5 12 3l9 7.5V20a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/>',
  alert: '<path d="M12 3 2 20h20L12 3z"/><path d="M12 10v4M12 17.5v.01"/>',
  fuel: '<path d="M4 21V5a2 2 0 0 1 2-2h7a2 2 0 0 1 2 2v16M3 21h13M15 10h2a2 2 0 0 1 2 2v4a1.5 1.5 0 0 0 3 0V8l-3-3"/><path d="M7 7h5v4H7z"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  scale: '<path d="M12 3v18M7 21h10M5 7h14M5 7l-3 7a3 3 0 0 0 6 0L5 7zM19 7l-3 7a3 3 0 0 0 6 0l-3-7z"/>',
  building: '<path d="M4 21V5l8-3v19M12 9h8v12M3 21h18M8 8v.01M8 12v.01M8 16v.01M16 13v.01M16 17v.01"/>',
  pin: '<path d="M12 21s-7-6.2-7-11.5a7 7 0 0 1 14 0C19 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/>',
  sheet: '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M4 9h16M4 15h16M10 3v18"/>',
  menu: '<path d="M4 6h16M4 12h16M4 18h16"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
  refresh: '<path d="M20 11a8 8 0 0 0-14.9-3.9L4 8.5M4 4v4.5h4.5M4 13a8 8 0 0 0 14.9 3.9l1.1-1.4M20 20v-4.5h-4.5"/>',
  wifiOff: '<path d="M2 2l20 20M8.5 16.5a5 5 0 0 1 7 0M5 13a10 10 0 0 1 5.2-2.8M19 13a10 10 0 0 0-2.3-1.6M2 8.8a15 15 0 0 1 4.2-2.7M22 8.8A15 15 0 0 0 11 5M12 20h.01"/>',
  users: '<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0M16 4.5a3.5 3.5 0 0 1 0 7M18 14a6.5 6.5 0 0 1 3.5 6"/>',
  debt: '<path d="M3 7h18v12H3z"/><path d="M3 11h18M7 15h4"/><path d="M16 3l-4 4-4-4"/>',
  drop: '<path d="M12 3c3.6 4.4 6.5 7.8 6.5 11.2a6.5 6.5 0 0 1-13 0C5.5 10.8 8.4 7.4 12 3z"/>',
  percent: '<path d="M19 5 5 19"/><circle cx="7" cy="7" r="2.5"/><circle cx="17" cy="17" r="2.5"/>',
  download: '<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>',
  print: '<path d="M7 9V3h10v6M7 18H5a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><path d="M7 14h10v7H7z"/>',
  excel: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="m8 8 8 8M16 8l-8 8"/>',
  file: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M8 13h8M8 17h5"/>',
  check: '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
  x: '<circle cx="12" cy="12" r="9"/><path d="m9 9 6 6M15 9l-6 6"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8v.01"/>',
  chevron: '<path d="m9 6 6 6-6 6"/>',
  inbox: '<path d="M3 13h5l2 3h4l2-3h5"/><path d="M5 5h14l2 8v6H3v-6z"/>',
  arrowDown: '<path d="M12 5v14M6 13l6 6 6-6"/>',
  arrowUp: '<path d="M12 19V5M6 11l6-6 6 6"/>'
};
const icon = (name) => `<svg class="ic" viewBox="0 0 24 24">${ICONS[name] || ''}</svg>`;
const ic = (name) => `<i data-icon="${name}">${icon(name)}</i>`;
function hydrateIcons(root = document) {
  root.querySelectorAll('i[data-icon]:empty').forEach(el => { el.innerHTML = icon(el.dataset.icon); });
}

/* ---------------- Yordamchi funksiyalar ---------------- */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const pad = (n) => String(n).padStart(2, '0');
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const normKey = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
const nz = (v) => Math.abs(v) > 1e-6;

const NF0 = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 });
const NF2 = new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
function fmt(n, dec = 0) {
  if (n === null || n === undefined || Number.isNaN(n)) return '—';
  const v = dec ? Math.round(n * 100) / 100 : Math.round(n);
  if (v === 0) return '0';
  return (dec ? NF2 : NF0).format(v).replace(/ /g, ' ');
}
const fmtSigned = (n, dec = 0) => (Math.round(n * (dec ? 100 : 1)) > 0 ? '+' : '') + fmt(n, dec);
const dateStr = (d) => `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()}`;
const timeStr = (d) => `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
const monthEnd = (y, m) => new Date(y, m, 0);          // m: 1..12
const monthStart = (y, m) => new Date(y, m - 1, 1);
const DAY = 86400000;

function store(key, val) { try { localStorage.setItem(key, JSON.stringify(val)); } catch (_) { /* xotira to‘la yoki ruxsat yo‘q */ } }
function load(key) { try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : null; } catch (_) { return null; } }
function sGet(k) { try { return sessionStorage.getItem(k) || ''; } catch (_) { return ''; } }
function sSet(k, v) { try { sessionStorage.setItem(k, v); } catch (_) { /* */ } }

const loadedLibs = {};
function loadScript(url) {
  if (!loadedLibs[url]) {
    loadedLibs[url] = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = url; s.async = true;
      s.onload = resolve;
      s.onerror = () => { delete loadedLibs[url]; reject(new Error('Kutubxona yuklanmadi. Internet aloqasini tekshiring.')); };
      document.head.appendChild(s);
    });
  }
  return loadedLibs[url];
}

/* ============================================================
   1. MA'LUMOTLAR MODELI (jadval formulalari asosida)
   ------------------------------------------------------------
   Har oy, har fermer uchun (litrda):
     Yakuniy qoldiq = Kreditor(boshi) + Kirim − Debitor(boshi) − Chiqim − Xizmat haqi
     Xizmat haqi    = jadvaldagi "Ёкилги саклаш хизмати / микдор" ustuni (Chiqim × 8%)
     Qarzdorlik     = Debitor (oy oxiri, кол-во)
     Qoldiq         = Kreditor (oy oxiri, кол-во)
   ============================================================ */
const isGroupHeader = (name, inn) => !inn && /массив|туман/i.test(name);

function buildModel(raw) {
  const months = (raw.months || []).map((m, idx) => ({
    idx,
    key: `${m.year}-${pad(m.month)}`,
    year: m.year,
    month: m.month,
    label: `${UZ_MONTHS[m.month - 1]} ${m.year}`,
    short: `${UZ_MONTHS[m.month - 1].slice(0, 3)} ${String(m.year).slice(2)}`,
    sheet: m.sheet,
    sheetId: m.sheetId,
    kirimBranches: m.kirimBranches || [],
    chiqimBranches: m.chiqimBranches || [],
    rows: m.rows || []
  }));

  // Filiallar — jadval sarlavhalaridan (2-qator) avtomatik
  const branchMap = new Map();
  const addBranch = (name) => {
    const k = normKey(name);
    if (!branchMap.has(k)) branchMap.set(k, { key: k, name: String(name).trim(), i: branchMap.size });
    return branchMap.get(k).i;
  };
  months.forEach(m => {
    m.kIdx = m.kirimBranches.map(addBranch);
    m.cIdx = m.chiqimBranches.map(addBranch);
  });
  const branches = [...branchMap.values()].map(b => ({ ...b, color: BRANCH_COLORS[b.i % BRANCH_COLORS.length] }));
  const B = branches.length;

  // Fermerlar — INN bo‘yicha (INN bo‘lmasa, nomi bo‘yicha) oylar kesimida birlashtiriladi
  const farmers = new Map();
  const groupHeaders = [];
  months.forEach(m => {
    let massiv = '';
    m.rows.forEach(r => {
      const inn = /^\d{5,}$/.test(String(r.inn).trim()) ? String(r.inn).trim() : '';
      if (isGroupHeader(r.name, inn)) { massiv = r.name; if (m.idx === months.length - 1) groupHeaders.push(r.name); return; }
      const key = inn || 'n:' + normKey(r.name);
      let f = farmers.get(key);
      if (!f) {
        f = { key, inn, name: r.name, massiv, rows: new Array(months.length).fill(null) };
        farmers.set(key, f);
      }
      f.name = r.name;
      if (massiv) f.massiv = massiv;

      const kb = new Array(B).fill(0), cb = new Array(B).fill(0);
      (r.kirim || []).forEach((v, i) => { kb[m.kIdx[i]] += v || 0; });
      (r.chiqim || []).forEach((v, i) => { cb[m.cIdx[i]] += v || 0; });
      const row = { ...r, kb, cb };

      const prev = f.rows[m.idx];
      if (prev) { // bir oyda bir xil INN ikki marta uchrasa — qo‘shiladi
        ['debQ', 'debS', 'kreQ', 'kreS', 'kirimQ', 'kirimS', 'cash', 'chiqimQ', 'chiqimS', 'servQ', 'servS', 'debEndQ', 'debEndS', 'kreEndQ', 'kreEndS']
          .forEach(k => { prev[k] = (prev[k] || 0) + (row[k] || 0); });
        prev.kb = prev.kb.map((v, i) => v + kb[i]);
        prev.cb = prev.cb.map((v, i) => v + cb[i]);
      } else {
        f.rows[m.idx] = row;
      }
    });
  });

  const list = [...farmers.values()];
  list.forEach(f => {
    const lat = latinKey(f.name);
    f.search = normKey(f.name + ' ' + f.inn) + ' ' + lat + ' ' + lat.replace(/x/g, 'h');
    // Filialga tegishlilik: har oy uchun "shu oygacha operatsiya qilgan filiallar"
    const seen = new Array(B).fill(false);
    f.branchesUpto = months.map((_, i) => {
      const r = f.rows[i];
      if (r) for (let b = 0; b < B; b++) if (nz(r.kb[b]) || nz(r.cb[b])) seen[b] = true;
      return seen.map((v, b) => v ? b : -1).filter(b => b >= 0);
    });
  });

  return {
    months, branches, farmers: list, groupHeaders,
    spreadsheet: raw.spreadsheet || 'Google Sheets',
    generatedAt: raw.generatedAt ? new Date(raw.generatedAt) : new Date(),
    warnings: raw.warnings || []
  };
}

/** Kirill → lotin (qidiruv lotin harflarida ham ishlashi uchun) */
function translit(s) {
  const map = { 'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'ғ': "g'", 'д': 'd', 'е': 'e', 'ё': 'yo', 'ж': 'j', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'қ': 'q', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ў': "o'", 'ф': 'f', 'х': 'x', 'ҳ': 'h', 'ц': 's', 'ч': 'ch', 'ш': 'sh', 'щ': 'sh', 'ъ': '', 'ы': 'i', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya' };
  return String(s).toLowerCase().split('').map(c => map[c] ?? c).join('');
}

/** Qidiruv kaliti: lotinga o‘giriladi, tutuq belgilari olib tashlanadi */
const latinKey = (s) => translit(normKey(s)).replace(/['ʻʼ’‘`]/g, '');
/** Qidiruv so‘zi fermerga mos keladimi (kirill yoki lotin, tutuq belgisiz ham) */
const matchQ = (f, q) => {
  if (!q) return true;
  const n = normKey(q);
  return f.search.includes(n) || f.search.includes(latinKey(n));
};

/** Fermer bo‘yicha davr statistikasi. bi = filial indeksi (−1 = barcha filiallar) */
function stats(f, from, to, bi = -1) {
  let first = null, last = null;
  const s = { kirim: 0, chiqim: 0, serv: 0, kirimS: 0, cash: 0, chiqimS: 0, servS: 0, active: false };
  for (let i = from; i <= to; i++) {
    const r = f.rows[i];
    if (!r) continue;
    if (!first) first = r;
    last = r;
    if (bi < 0) {
      s.kirim += r.kirimQ; s.chiqim += r.chiqimQ; s.serv += r.servQ;
    } else {
      s.kirim += r.kb[bi]; s.chiqim += r.cb[bi];
      // filial ulushi: fermerning shu oydagi xizmat haqi chiqim hajmiga mutanosib taqsimlanadi
      s.serv += r.chiqimQ > 0 ? r.servQ * r.cb[bi] / r.chiqimQ : 0;
    }
    s.kirimS += r.kirimS; s.cash += r.cash; s.chiqimS += r.chiqimS; s.servS += r.servS;
    if (nz(r.kirimQ) || nz(r.chiqimQ) || nz(r.kirimS) || nz(r.cash) || nz(r.chiqimS)) s.active = true;
  }
  if (!last) return null;
  // Yakuniy qoldiq faqat davrning oxirgi oyi varag‘idan olinadi (jadvaldagi "Жами" bilan mos bo‘lishi uchun).
  // Fermer oxirgi oy varag‘ida bo‘lmasa, uning qoldig‘i noma'lum — 0 deb olinadi va belgilanadi.
  const end = f.rows[to];
  s.missingAtEnd = !end;
  s.openDebt = first.debQ; s.openStock = first.kreQ; s.open = first.kreQ - first.debQ;
  s.openS = first.kreS - first.debS;
  s.debt = end ? end.debEndQ : 0; s.stock = end ? end.kreEndQ : 0; s.close = s.stock - s.debt;
  s.debtS = end ? end.debEndS : 0; s.stockS = end ? end.kreEndS : 0; s.closeS = s.stockS - s.debtS;
  return s;
}

/** Oxirgi harakat bo‘lgan oy indeksi (upto gacha), topilmasa −1 */
function lastActivity(f, upto) {
  for (let i = upto; i >= 0; i--) {
    const r = f.rows[i];
    if (r && (nz(r.kirimQ) || nz(r.chiqimQ) || nz(r.kirimS) || nz(r.cash) || nz(r.chiqimS))) return i;
  }
  return -1;
}

/** Davr bo‘yicha barcha fermerlar qatori (sahifalar uchun umumiy) */
function farmerRows(bi = -1) {
  const { from, to } = S.range;
  const out = [];
  for (const f of M.farmers) {
    const st = stats(f, from, to, bi);
    if (!st) continue;
    const memberOf = f.branchesUpto[to];
    if (bi >= 0 && !memberOf.includes(bi)) continue;
    out.push({ f, st, branches: memberOf });
  }
  return out;
}

/* ============================================================
   2. HOLAT, YUKLASH, ULANISH
   ============================================================ */
let M = null;                      // model
const S = {                        // UI holati
  range: { from: 0, to: 0 },
  sort: {},
  filters: {},
  charts: [],
  loading: false,
  fromCache: false
};

async function fetchData({ silent = false } = {}) {
  if (S.loading) return;
  S.loading = true;
  const btn = $('#refresh-btn');
  btn.classList.add('spinning'); btn.disabled = true;
  if (!silent) showLoader(true);
  setStatus('wait', 'Ulanmoqda…');

  try {
    if (!navigator.onLine) throw Object.assign(new Error('Internet aloqasi yo‘q.'), { offline: true });
    const headers = {};
    const pass = sGet(CFG.passKey);
    if (pass) headers['x-app-key'] = pass;

    let res = await fetch(`${CFG.api}?t=${Date.now()}`, { headers, cache: 'no-store' });
    if (res.status === 404) {
      // netlify.toml'dagi yo‘naltirish ishlamasa — funksiyaning to‘g‘ridan-to‘g‘ri manzili
      res = await fetch(`/.netlify/functions/sheets?t=${Date.now()}`, { headers, cache: 'no-store' });
      if (res.status === 404) throw new Error('Server funksiyasi topilmadi (404). Netlify\'da "netlify/functions/sheets.js" joylashmagan: sayt drag & drop orqali yuklangan yoki netlify.toml loyiha ildizida emas. README, 2-qadamga qarang.');
    }
    let data = null;
    try { data = await res.json(); } catch (_) { /* JSON emas */ }

    if (res.status === 401 || (data && data.auth)) { showLogin(pass ? 'Parol noto‘g‘ri.' : ''); setStatus('err', 'Kirish talab qilinadi'); return; }
    if (!res.ok || !data || !data.ok) throw new Error((data && data.error) || `Server xatosi (${res.status}).`);
    if (!data.months || !data.months.length) throw new Error('Jadvalda oylik varaqlar topilmadi.');

    applyData(data, false);
    store(CFG.cacheKey, data);
    setStatus('ok', 'Google Sheets ulangan');
    if (silent) toast('ok', 'Ma\'lumotlar yangilandi');
    (data.warnings || []).forEach(w => toast('warn', w));
  } catch (err) {
    const cached = load(CFG.cacheKey);
    if (!M && cached) {
      applyData(cached, true);
      toast('warn', 'Google Sheets bilan bog‘lanib bo‘lmadi. Oxirgi saqlangan ma\'lumotlar ko‘rsatilmoqda.');
    }
    if (err.offline || !navigator.onLine) setStatus('off', 'Oflayn');
    else setStatus('err', 'Ulanishda xato');
    if (!M) renderFatal(err.message);
    else toast('err', err.message);
  } finally {
    S.loading = false;
    btn.classList.remove('spinning'); btn.disabled = false;
    showLoader(false);
  }
}

function applyData(raw, fromCache) {
  const prevKeys = M ? { from: M.months[S.range.from]?.key, to: M.months[S.range.to]?.key } : null;
  M = buildModel(raw);
  S.fromCache = fromCache;
  $('#offline-banner').hidden = !fromCache;
  $('#updated-at').textContent = `${dateStr(M.generatedAt)} ${timeStr(M.generatedAt)}`;
  $('#src-name').textContent = M.spreadsheet;

  const last = M.months.length - 1;
  if (prevKeys) {
    const fi = M.months.findIndex(m => m.key === prevKeys.from);
    const ti = M.months.findIndex(m => m.key === prevKeys.to);
    // yangi oy qo‘shilgan bo‘lsa va oldin oxirgi oy tanlangan bo‘lsa — avtomatik yangi oyga o‘tadi
    S.range.to = ti < 0 || ti === last - 1 ? last : ti;
    S.range.from = fi < 0 ? 0 : Math.min(fi, S.range.to);
  } else {
    setQuickRange('ytd', false);
  }
  buildPeriodBar();
  buildBranchNav();
  route();
}

function setStatus(kind, text) {
  const pill = $('#status-pill');
  pill.className = 'status-pill ' + (kind === 'ok' ? 'ok' : kind === 'err' ? 'err' : kind === 'off' ? 'off' : '');
  $('#status-text').textContent = text;
  pill.title = text;
}
function showLoader(on) { $('#loader').classList.toggle('hide', !on); }

function showLogin(msg) {
  showLoader(false);
  $('#login-modal').hidden = false;
  $('#login-error').textContent = msg || '';
  setTimeout(() => $('#login-pass').focus(), 50);
}

function toast(kind, text) {
  const ico = kind === 'ok' ? 'check' : kind === 'err' ? 'x' : 'info';
  const el = document.createElement('div');
  el.className = `toast ${kind}`;
  el.innerHTML = `${ic(ico)}<div>${esc(text)}</div>`;
  $('#toasts').appendChild(el);
  setTimeout(() => { el.style.opacity = '0'; el.style.transition = 'opacity .3s'; setTimeout(() => el.remove(), 320); }, kind === 'err' ? 7000 : 4200);
}

/* ============================================================
   3. DAVR PANELI (yil, oy, oraliq)
   ============================================================ */
function buildPeriodBar() {
  const years = [...new Set(M.months.map(m => m.year))];
  const ys = $('#year-select');
  ys.innerHTML = `<option value="">Barcha yillar</option>` + years.map(y => `<option value="${y}">${y}</option>`).join('');
  const opts = M.months.map(m => `<option value="${m.idx}">${m.label}</option>`).join('');
  $('#from-select').innerHTML = opts;
  $('#to-select').innerHTML = opts;
  syncPeriodBar();
}

function syncPeriodBar() {
  const { from, to } = S.range;
  $('#from-select').value = from;
  $('#to-select').value = to;
  const fy = M.months[from].year, ty = M.months[to].year;
  $('#year-select').value = fy === ty && isFullYear(fy, from, to) ? fy : '';
  const last = M.months.length - 1;
  const ytdFrom = M.months.findIndex(m => m.year === M.months[last].year);
  const active = to !== last ? '' : from === last ? 'last' : from === ytdFrom ? 'ytd' : from === 0 ? 'all' : '';
  $$('#quick-periods button').forEach(b => b.classList.toggle('on', b.dataset.q === active));
  const a = M.months[from], b = M.months[to];
  $('#period-note').innerHTML = `Hisobot davri: <b>${dateStr(monthStart(a.year, a.month))} – ${dateStr(monthEnd(b.year, b.month))}</b>`;
}
function isFullYear(y, from, to) {
  const idx = M.months.filter(m => m.year === y).map(m => m.idx);
  return idx[0] === from && idx[idx.length - 1] === to;
}
function setQuickRange(q, rerender = true) {
  const last = M.months.length - 1;
  if (q === 'last') S.range = { from: last, to: last };
  if (q === 'all') S.range = { from: 0, to: last };
  if (q === 'ytd') S.range = { from: M.months.findIndex(m => m.year === M.months[last].year), to: last };
  if (rerender) { syncPeriodBar(); route(); }
}

/* ============================================================
   4. NAVIGATSIYA
   ============================================================ */
function buildBranchNav() {
  $('#nav-branches').innerHTML = M.branches.map(b =>
    `<a href="#/filial/${b.i}" data-route="filial/${b.i}"><span class="branch-dot" style="background:${b.color}"></span><span>${esc(b.name)}</span></a>`
  ).join('');
}

const ROUTES = {
  bosh: renderDashboard,
  qarzdorlar: renderDebtors,
  yoqilgi: renderHolders,
  harakatsiz: renderInactive,
  aksverka: renderAksverka,
  filiallar: renderBranchesOverview,
  filial: renderBranch
};

function route() {
  if (!M) return;
  const hash = location.hash.replace(/^#\/?/, '') || 'bosh';
  const [name, arg] = hash.split('/');
  const fn = ROUTES[name] || renderDashboard;
  S.charts.forEach(c => c.destroy()); S.charts = [];
  $$('#nav a').forEach(a => a.classList.toggle('active', a.dataset.route === (arg !== undefined && name === 'filial' ? `filial/${arg}` : name)));
  document.body.classList.remove('nav-open');
  const view = $('#view');
  view.innerHTML = '';
  fn(view, arg);
  hydrateIcons(view);
}

/* ============================================================
   5. UMUMIY UI BLOKLARI
   ============================================================ */
function pageHead(title, desc, actions = '') {
  return `<div class="page-head"><div><h2>${title}</h2>${desc ? `<p>${desc}</p>` : ''}</div>${actions ? `<div class="actions">${actions}</div>` : ''}</div>`;
}
function kpi(label, value, unit, sub, iconName, tone = '') {
  return `<div class="glass kpi ${tone}">
    <div class="kpi-top"><span class="kpi-label">${label}</span><span class="kpi-ico">${ic(iconName)}</span></div>
    <div class="kpi-value">${value}${unit ? `<small>${unit}</small>` : ''}</div>
    ${sub ? `<div class="kpi-sub">${sub}</div>` : ''}
  </div>`;
}
const branchTags = (idxs) => idxs.length
  ? idxs.map(i => `<span class="tag"><span class="sw" style="background:${M.branches[i].color}"></span>${esc(M.branches[i].name)}</span>`).join('')
  : '<span class="muted">—</span>';
const nameCell = (f) => `<div class="fname">${esc(f.name)}<small>${f.inn ? 'INN ' + esc(f.inn) : 'INN ko‘rsatilmagan'}${f.massiv ? ' · ' + esc(f.massiv) : ''}</small></div>`;
const rangeLabel = () => {
  const a = M.months[S.range.from], b = M.months[S.range.to];
  return a.idx === b.idx ? a.label : `${a.label} – ${b.label}`;
};
const emptyState = (title, text) => `<div class="empty-state">${ic('inbox')}<h4>${title}</h4><div>${text}</div></div>`;

function branchOptions(selected, allLabel = 'Barcha filiallar') {
  return `<option value="-1">${allLabel}</option>` + M.branches.map(b => `<option value="${b.i}" ${+selected === b.i ? 'selected' : ''}>${esc(b.name)}</option>`).join('');
}

/** Saralanadigan jadval */
function dataTable(id, columns, rows, { defaultSort, footer, onRowClick } = {}) {
  const st = S.sort[id] || defaultSort || { key: columns[1].key, dir: 'desc' };
  S.sort[id] = st;
  const col = columns.find(c => c.key === st.key);
  if (col && col.sortVal) {
    rows = [...rows].sort((a, b) => {
      const va = col.sortVal(a), vb = col.sortVal(b);
      const r = typeof va === 'string' ? va.localeCompare(vb, 'ru') : va - vb;
      return st.dir === 'asc' ? r : -r;
    });
  }
  const head = columns.map(c => {
    const sorted = c.key === st.key;
    const arrow = sorted ? (st.dir === 'asc' ? '▲' : '▼') : '↕';
    return `<th class="${c.num ? 'num' : ''} ${sorted ? 'sorted' : ''}" data-sort="${c.sortVal ? c.key : ''}">${c.label}${c.sortVal ? `<span class="sort">${arrow}</span>` : ''}</th>`;
  }).join('');
  const body = rows.map((r, i) => `<tr class="${onRowClick ? 'clickable' : ''}" data-i="${i}">${columns.map(c => `<td class="${c.num ? 'num' : ''} ${c.cls || ''}">${c.render(r, i)}</td>`).join('')}</tr>`).join('');
  const foot = footer ? `<tfoot><tr>${columns.map(c => `<td class="${c.num ? 'num' : ''}">${footer[c.key] ?? ''}</td>`).join('')}</tr></tfoot>` : '';
  const html = `<div class="glass table-card"><div class="table-scroll"><table class="data" id="${id}"><thead><tr>${head}</tr></thead><tbody>${body}</tbody>${foot}</table></div></div>`;
  return {
    html,
    bind(root, rerender) {
      const t = $(`#${id}`, root);
      if (!t) return;
      $$('th[data-sort]', t).forEach(th => th.addEventListener('click', () => {
        const k = th.dataset.sort; if (!k) return;
        S.sort[id] = { key: k, dir: S.sort[id].key === k && S.sort[id].dir === 'desc' ? 'asc' : 'desc' };
        rerender();
      }));
      if (onRowClick) $$('tbody tr', t).forEach(tr => tr.addEventListener('click', () => onRowClick(rows[+tr.dataset.i])));
    }
  };
}

const openAkt = (f) => { S.filters.akt = { ...(S.filters.akt || {}), farmer: f.key }; location.hash = '#/aksverka'; };

/** Excel eksport (umumiy) */
async function exportExcel(filename, sheetName, aoa, colWidths) {
  try {
    await loadScript(CFG.libs.xlsx);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (colWidths) ws['!cols'] = colWidths.map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
    XLSX.writeFile(wb, filename);
    toast('ok', 'Excel fayl yuklab olindi');
  } catch (e) { toast('err', e.message); }
}

/* ---------------- Grafiklar ---------------- */
function chart(canvas, config) {
  if (!window.Chart) { canvas.parentElement.innerHTML = `<div class="empty-state">Grafik kutubxonasi yuklanmadi.</div>`; return; }
  Chart.defaults.color = '#a9b6cf';
  Chart.defaults.font.family = "'Manrope', system-ui, sans-serif";
  Chart.defaults.font.size = 12;
  Chart.defaults.borderColor = 'rgba(148,178,232,0.10)';
  Chart.defaults.plugins.legend.labels.boxWidth = 10;
  Chart.defaults.plugins.legend.labels.boxHeight = 10;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.tooltip.backgroundColor = '#0f1a30';
  Chart.defaults.plugins.tooltip.borderColor = 'rgba(148,178,232,0.25)';
  Chart.defaults.plugins.tooltip.borderWidth = 1;
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.callbacks.label = (ctx) => `${ctx.dataset.label || ctx.label}: ${fmt(ctx.parsed.y ?? ctx.parsed.x ?? ctx.parsed)} l`;
  Chart.defaults.maintainAspectRatio = false;
  Chart.defaults.animation.duration = 700;
  const c = new Chart(canvas, config);
  S.charts.push(c);
  return c;
}
const axisL = { ticks: { callback: (v) => fmt(v) }, grid: { color: 'rgba(148,178,232,0.07)' } };
const grad = (ctx, a, b) => { const g = ctx.createLinearGradient(0, 0, 0, 280); g.addColorStop(0, a); g.addColorStop(1, b); return g; };

/* ============================================================
   6. SAHIFALAR
   ============================================================ */

/* ---------- 6.1 Bosh sahifa ---------- */
function renderDashboard(view) {
  const { from, to } = S.range;
  const rows = farmerRows();
  const presentNow = M.farmers.filter(f => f.rows[to]);
  const debtors = rows.filter(r => r.st.debt > CFG.EPS);
  const holders = rows.filter(r => r.st.stock > CFG.EPS);
  const totalDebt = debtors.reduce((a, r) => a + r.st.debt, 0);
  const totalStock = holders.reduce((a, r) => a + r.st.stock, 0);
  const totalServ = rows.reduce((a, r) => a + r.st.serv, 0);
  const totalServS = rows.reduce((a, r) => a + r.st.servS, 0);
  const activeCnt = rows.filter(r => r.st.active).length;
  const endLabel = dateStr(monthEnd(M.months[to].year, M.months[to].month));

  // filiallar
  const bStats = M.branches.map(b => {
    const rs = rows.map(r => ({ r, s: stats(r.f, from, to, b.i) })).filter(x => x.s && (nz(x.s.kirim) || nz(x.s.chiqim)));
    return {
      b,
      kirim: rs.reduce((a, x) => a + x.s.kirim, 0),
      chiqim: rs.reduce((a, x) => a + x.s.chiqim, 0),
      serv: rs.reduce((a, x) => a + x.s.serv, 0),
      farmers: rs.length
    };
  });
  const tot = {
    kirim: rows.reduce((a, r) => a + r.st.kirim, 0),
    chiqim: rows.reduce((a, r) => a + r.st.chiqim, 0),
    serv: totalServ,
    farmers: activeCnt
  };

  view.innerHTML = `
    ${pageHead('Bosh sahifa', `Umumiy moliyaviy holat · ${rangeLabel()} · qoldiqlar ${endLabel} holatiga`)}
    <div class="kpi-grid">
      ${kpi('Jami fermerlar soni', fmt(presentNow.length), '', `Davrda faol: <b>${fmt(activeCnt)}</b> ta`, 'users')}
      ${kpi('Jami qarzdorlik', fmt(totalDebt), 'l', `${fmt(debtors.length)} ta fermer qarzdor`, 'debt', 'red')}
      ${kpi('Jami yoqilg‘i qoldig‘i', fmt(totalStock), 'l', `${fmt(holders.length)} ta fermerda qoldiq bor`, 'drop', 'green')}
      ${kpi('Hisoblangan xizmat haqi', fmt(totalServ), 'l', nz(totalServS) ? `So‘mda: <b>${fmt(totalServS)}</b>` : 'Chiqim × 8% (jadval bo‘yicha)', 'percent', 'amber')}
      ${kpi('Filiallar soni', fmt(M.branches.length), '', M.branches.map(b => esc(b.name)).join(' · '), 'building', 'cyan')}
    </div>

    <div class="section-title"><h3>Filiallar kesimida</h3><span>${rangeLabel()} · litrda</span></div>
    <div class="branch-grid">
      ${bStats.map(x => branchCard(x, tot)).join('')}
      <div class="glass branch-card total">
        <div class="bc-head"><div class="bc-name">${ic('building')} Barcha filiallar</div></div>
        <div class="bc-rows">
          <div><span>Kirim</span><b>${fmt(tot.kirim)}</b></div>
          <div><span>Chiqim</span><b>${fmt(tot.chiqim)}</b></div>
          <div><span>Xizmat haqi</span><b>${fmt(tot.serv)}</b></div>
          <div><span>Faol fermerlar</span><b>${fmt(tot.farmers)}</b></div>
        </div>
      </div>
    </div>

    <div class="section-title"><h3>Grafiklar</h3><span>Google Sheets ma'lumotlari asosida</span></div>
    <div class="chart-grid">
      <div class="glass chart-card"><h4>Oylar bo‘yicha kirim va chiqim</h4><p>Litrda, tanlangan davr</p><div class="chart-box"><canvas id="ch-flow"></canvas></div></div>
      <div class="glass chart-card"><h4>Chiqim filiallar bo‘yicha</h4><p>Ulush, litrda</p><div class="chart-box"><canvas id="ch-branch"></canvas></div></div>
    </div>
    <div class="chart-grid even" style="margin-top:16px">
      <div class="glass chart-card"><h4>Qarzdorlik va qoldiq dinamikasi</h4><p>Har oy oxiriga, litrda</p><div class="chart-box"><canvas id="ch-balance"></canvas></div></div>
      <div class="glass chart-card"><h4>Eng katta 10 qarzdor</h4><p>${endLabel} holatiga, litrda</p><div class="chart-box"><canvas id="ch-top"></canvas></div></div>
    </div>`;

  // Grafiklar
  const idx = []; for (let i = from; i <= to; i++) idx.push(i);
  const sumMonth = (i, fn) => M.farmers.reduce((a, f) => a + (f.rows[i] ? fn(f.rows[i]) : 0), 0);
  const ctxFlow = $('#ch-flow').getContext('2d');
  chart(ctxFlow, {
    type: 'bar',
    data: {
      labels: idx.map(i => M.months[i].short),
      datasets: [
        { label: 'Kirim', data: idx.map(i => sumMonth(i, r => r.kirimQ)), backgroundColor: grad(ctxFlow, 'rgba(106,169,255,.95)', 'rgba(61,127,224,.35)'), borderRadius: 6, maxBarThickness: 26 },
        { label: 'Chiqim', data: idx.map(i => sumMonth(i, r => r.chiqimQ)), backgroundColor: grad(ctxFlow, 'rgba(79,209,165,.95)', 'rgba(79,209,165,.3)'), borderRadius: 6, maxBarThickness: 26 },
        { label: 'Xizmat haqi', type: 'line', data: idx.map(i => sumMonth(i, r => r.servQ)), borderColor: '#ffc46b', backgroundColor: '#ffc46b', tension: .35, pointRadius: 3, borderWidth: 2 }
      ]
    },
    options: { scales: { y: axisL, x: { grid: { display: false } } }, plugins: { legend: { position: 'top', align: 'end' } } }
  });

  chart($('#ch-branch'), {
    type: 'doughnut',
    data: { labels: bStats.map(x => x.b.name), datasets: [{ data: bStats.map(x => x.chiqim), backgroundColor: bStats.map(x => x.b.color), borderColor: '#111a2d', borderWidth: 3, hoverOffset: 8 }] },
    options: {
      cutout: '66%',
      plugins: {
        legend: { position: 'bottom' },
        tooltip: { callbacks: { label: (c) => { const t = c.dataset.data.reduce((a, v) => a + v, 0) || 1; return `${c.label}: ${fmt(c.parsed)} l (${(c.parsed / t * 100).toFixed(1)}%)`; } } }
      }
    }
  });

  const ctxBal = $('#ch-balance').getContext('2d');
  chart(ctxBal, {
    type: 'line',
    data: {
      labels: idx.map(i => M.months[i].short),
      datasets: [
        { label: 'Qarzdorlik', data: idx.map(i => sumMonth(i, r => r.debEndQ)), borderColor: '#ff7a8a', backgroundColor: grad(ctxBal, 'rgba(255,122,138,.28)', 'rgba(255,122,138,0)'), fill: true, tension: .35, pointRadius: 3, borderWidth: 2 },
        { label: 'Qoldiq', data: idx.map(i => sumMonth(i, r => r.kreEndQ)), borderColor: '#4fd1a5', backgroundColor: grad(ctxBal, 'rgba(79,209,165,.22)', 'rgba(79,209,165,0)'), fill: true, tension: .35, pointRadius: 3, borderWidth: 2 }
      ]
    },
    options: { scales: { y: axisL, x: { grid: { display: false } } }, plugins: { legend: { position: 'top', align: 'end' } }, interaction: { mode: 'index', intersect: false } }
  });

  const top = [...debtors].sort((a, b) => b.st.debt - a.st.debt).slice(0, 10);
  const ctxTop = $('#ch-top').getContext('2d');
  chart(ctxTop, {
    type: 'bar',
    data: { labels: top.map(r => r.f.name.length > 26 ? r.f.name.slice(0, 25) + '…' : r.f.name), datasets: [{ label: 'Qarzdorlik', data: top.map(r => r.st.debt), backgroundColor: 'rgba(255,122,138,.75)', borderRadius: 5, maxBarThickness: 18 }] },
    options: {
      indexAxis: 'y',
      scales: { x: axisL, y: { grid: { display: false }, ticks: { font: { size: 11 } } } },
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => `${fmt(c.parsed.x)} l` } } },
      onClick: (_, els) => { if (els[0]) openAkt(top[els[0].index].f); }
    }
  });
}

function branchCard(x, tot) {
  const share = tot.chiqim > 0 ? x.chiqim / tot.chiqim * 100 : 0;
  return `<a class="glass branch-card" href="#/filial/${x.b.i}">
    <div class="bc-head"><div class="bc-name"><span class="sw" style="background:${x.b.color}"></span>${esc(x.b.name)}</div>${ic('chevron')}</div>
    <div class="bc-rows">
      <div><span>Kirim</span><b>${fmt(x.kirim)}</b></div>
      <div><span>Chiqim</span><b>${fmt(x.chiqim)}</b></div>
      <div><span>Xizmat haqi</span><b>${fmt(x.serv)}</b></div>
      <div><span>Fermerlar</span><b>${fmt(x.farmers)}</b></div>
    </div>
    <div class="share-bar" title="Chiqimdagi ulushi: ${share.toFixed(1)}%"><span style="width:${share}%;background:${x.b.color}"></span></div>
  </a>`;
}

/* ---------- Umumiy filtr paneli ---------- */
function filterBar(key, { minLabel, showMin = true, extra = '' } = {}) {
  const f = S.filters[key] = S.filters[key] || { q: '', b: -1, min: '' };
  return `<div class="glass filters">
    <div class="field grow"><label>Qidiruv</label><input type="search" data-f="q" placeholder="Fermer xo‘jaligi nomi yoki INN…" value="${esc(f.q)}"></div>
    <div class="field"><label>Filial</label><select data-f="b">${branchOptions(f.b)}</select></div>
    ${showMin ? `<div class="field"><label>${minLabel}</label><input type="number" min="0" step="1" data-f="min" placeholder="0" value="${esc(f.min)}"></div>` : ''}
    ${extra}
  </div>`;
}
function bindFilters(view, key, rerender) {
  $$('.filters [data-f]', view).forEach(el => {
    const handler = () => {
      const f = S.filters[key];
      f[el.dataset.f] = el.dataset.f === 'b' ? +el.value : el.value;
      if (el.dataset.f === 'q') {
        clearTimeout(el._t);
        el._t = setTimeout(() => { rerender(); const ni = $(`.filters [data-f="q"]`, view); if (ni) { ni.focus(); ni.setSelectionRange(ni.value.length, ni.value.length); } }, 220);
      } else rerender();
    };
    el.addEventListener(el.tagName === 'SELECT' ? 'change' : 'input', handler);
  });
}

/* ---------- 6.2 Eng katta qarzdorlar ---------- */
function renderDebtors(view) {
  const key = 'debt';
  const F = S.filters[key] || { q: '', b: -1, min: '' };
  const bi = +F.b;
  const min = parseFloat(F.min) || 0;
  const rows = farmerRows(bi).filter(r => r.st.debt > CFG.EPS && r.st.debt >= min && matchQ(r.f, F.q));
  const maxDebt = Math.max(1, ...rows.map(r => r.st.debt));
  const bl = bi >= 0 ? ` (${esc(M.branches[bi].name)})` : '';
  const endLabel = dateStr(monthEnd(M.months[S.range.to].year, M.months[S.range.to].month));

  const sums = rows.reduce((a, r) => ({ debt: a.debt + r.st.debt, k: a.k + r.st.kirim, c: a.c + r.st.chiqim, s: a.s + r.st.serv, q: a.q + r.st.close }), { debt: 0, k: 0, c: 0, s: 0, q: 0 });
  const cols = [
    { key: 'n', label: '№', cls: 'idx', render: (_, i) => i + 1 },
    { key: 'name', label: 'Fermer xo‘jaligi nomi', render: r => nameCell(r.f), sortVal: r => r.f.name },
    { key: 'br', label: 'Filial', render: r => branchTags(r.branches) },
    { key: 'k', label: 'Kirim' + bl, num: 1, render: r => fmt(r.st.kirim), sortVal: r => r.st.kirim },
    { key: 'c', label: 'Chiqim' + bl, num: 1, render: r => fmt(r.st.chiqim), sortVal: r => r.st.chiqim },
    { key: 'q', label: 'Qoldiq', num: 1, render: r => `<span class="${r.st.close < 0 ? 'neg' : 'pos'}">${fmtSigned(r.st.close)}</span>`, sortVal: r => r.st.close },
    { key: 'debt', label: 'Qarzdorlik', num: 1, render: r => `<div class="bar-cell"><b class="neg">${fmt(r.st.debt)}</b><span class="mini"><span style="width:${r.st.debt / maxDebt * 100}%"></span></span></div>`, sortVal: r => r.st.debt },
    { key: 's', label: 'Xizmat haqi' + bl, num: 1, render: r => fmt(r.st.serv), sortVal: r => r.st.serv }
  ];
  const t = dataTable('t-debt', cols, rows, {
    defaultSort: { key: 'debt', dir: 'desc' },
    footer: { name: `Jami: ${fmt(rows.length)} ta`, k: fmt(sums.k), c: fmt(sums.c), q: fmtSigned(sums.q), debt: fmt(sums.debt), s: fmt(sums.s) },
    onRowClick: r => openAkt(r.f)
  });

  view.innerHTML = `
    ${pageHead('Eng katta yoqilg‘i qarzdorliklar', `Qarzdorlik — ${endLabel} holatiga Debitor (litr). Kirim, chiqim va xizmat haqi — ${rangeLabel()} davri uchun. Qatorni bossangiz, aksverka ochiladi.`,
      `<button class="btn" id="xl">${ic('excel')}Excel</button>`)}
    ${filterBar(key, { minLabel: 'Qarz kamida (l)' })}
    ${bi >= 0 ? `<div class="note">${ic('info')}<div>Filial tanlanganda shu filialda operatsiya qilgan fermerlar ko‘rinadi. Kirim, chiqim va xizmat haqi shu filial bo‘yicha, qoldiq va qarzdorlik esa fermerning barcha filiallar bo‘yicha umumiy ko‘rsatkichi.</div></div>` : ''}
    <div class="summary-strip">
      <div class="pill">Qarzdorlar:<b>${fmt(rows.length)}</b></div>
      <div class="pill">Jami qarzdorlik:<b class="neg">${fmt(sums.debt)} l</b></div>
      <div class="pill">Xizmat haqi:<b>${fmt(sums.s)} l</b></div>
    </div>
    ${rows.length ? t.html : `<div class="glass">${emptyState('Qarzdor topilmadi', 'Filtr shartlarini o‘zgartirib ko‘ring.')}</div>`}`;

  const rerender = () => { renderDebtors(view); hydrateIcons(view); };
  bindFilters(view, key, rerender);
  t.bind(view, rerender);
  $('#xl', view).onclick = () => exportExcel(`Qarzdorlar_${M.months[S.range.to].key}.xlsx`, 'Qarzdorlar', [
    [CFG.company], [`Eng katta yoqilg‘i qarzdorliklar · ${rangeLabel()}${bi >= 0 ? ' · ' + M.branches[bi].name : ''}`], [],
    ['№', 'Fermer xo‘jaligi', 'INN', 'Filial', 'Kirim (l)', 'Chiqim (l)', 'Qoldiq (l)', 'Qarzdorlik (l)', 'Xizmat haqi (l)'],
    ...sortedFor('t-debt', cols, rows).map((r, i) => [i + 1, r.f.name, r.f.inn, r.branches.map(b => M.branches[b].name).join(', '), rnd(r.st.kirim), rnd(r.st.chiqim), rnd(r.st.close), rnd(r.st.debt), rnd(r.st.serv)]),
    ['', 'Jami', '', '', rnd(sums.k), rnd(sums.c), rnd(sums.q), rnd(sums.debt), rnd(sums.s)]
  ], [5, 38, 12, 24, 12, 12, 12, 14, 14]);
}
const rnd = (v) => Math.round(v * 100) / 100;
function sortedFor(id, cols, rows) {
  const st = S.sort[id]; const col = cols.find(c => c.key === st.key);
  if (!col || !col.sortVal) return rows;
  return [...rows].sort((a, b) => { const va = col.sortVal(a), vb = col.sortVal(b); const r = typeof va === 'string' ? va.localeCompare(vb, 'ru') : va - vb; return st.dir === 'asc' ? r : -r; });
}

/* ---------- 6.3 Yoqilg‘isi borlar ---------- */
function renderHolders(view) {
  const key = 'hold';
  const F = S.filters[key] || { q: '', b: -1, min: '' };
  const bi = +F.b;
  const min = parseFloat(F.min) || 0;
  const rows = farmerRows(bi).filter(r => r.st.stock > CFG.EPS && r.st.stock >= min && matchQ(r.f, F.q));
  const maxS = Math.max(1, ...rows.map(r => r.st.stock));
  const bl = bi >= 0 ? ` (${esc(M.branches[bi].name)})` : '';
  const endLabel = dateStr(monthEnd(M.months[S.range.to].year, M.months[S.range.to].month));
  const sums = rows.reduce((a, r) => ({ st: a.st + r.st.stock, k: a.k + r.st.kirim, c: a.c + r.st.chiqim, s: a.s + r.st.serv }), { st: 0, k: 0, c: 0, s: 0 });

  const cols = [
    { key: 'n', label: '№', cls: 'idx', render: (_, i) => i + 1 },
    { key: 'name', label: 'Nomi', render: r => nameCell(r.f), sortVal: r => r.f.name },
    { key: 'br', label: 'Filiali', render: r => branchTags(r.branches) },
    { key: 'k', label: 'Jami kirim' + bl, num: 1, render: r => fmt(r.st.kirim), sortVal: r => r.st.kirim },
    { key: 'c', label: 'Jami chiqim' + bl, num: 1, render: r => fmt(r.st.chiqim), sortVal: r => r.st.chiqim },
    { key: 'st', label: 'Qolgan yoqilg‘i', num: 1, render: r => `<div class="bar-cell"><b class="pos">${fmt(r.st.stock)}</b><span class="mini g"><span style="width:${r.st.stock / maxS * 100}%"></span></span></div>`, sortVal: r => r.st.stock },
    { key: 's', label: 'Xizmat haqi' + bl, num: 1, render: r => fmt(r.st.serv), sortVal: r => r.st.serv }
  ];
  const t = dataTable('t-hold', cols, rows, {
    defaultSort: { key: 'st', dir: 'desc' },
    footer: { name: `Jami: ${fmt(rows.length)} ta`, k: fmt(sums.k), c: fmt(sums.c), st: fmt(sums.st), s: fmt(sums.s) },
    onRowClick: r => openAkt(r.f)
  });

  view.innerHTML = `
    ${pageHead('Yoqilg‘isi borlar', `Hisobida ijobiy qoldiq (Kreditor) mavjud fermer xo‘jaliklari · ${endLabel} holatiga. Kirim/chiqim — ${rangeLabel()}.`,
      `<button class="btn" id="xl">${ic('excel')}Excel</button>`)}
    ${filterBar(key, { minLabel: 'Qoldiq kamida (l)' })}
    <div class="summary-strip">
      <div class="pill">Fermerlar:<b>${fmt(rows.length)}</b></div>
      <div class="pill">Jami qoldiq:<b class="pos">${fmt(sums.st)} l</b></div>
      <div class="pill">Xizmat haqi:<b>${fmt(sums.s)} l</b></div>
    </div>
    ${rows.length ? t.html : `<div class="glass">${emptyState('Ma\'lumot topilmadi', 'Filtr shartlarini o‘zgartirib ko‘ring.')}</div>`}`;

  const rerender = () => { renderHolders(view); hydrateIcons(view); };
  bindFilters(view, key, rerender);
  t.bind(view, rerender);
  $('#xl', view).onclick = () => exportExcel(`Yoqilgisi_borlar_${M.months[S.range.to].key}.xlsx`, 'Yoqilgisi borlar', [
    [CFG.company], [`Yoqilg‘isi borlar · ${rangeLabel()}`], [],
    ['№', 'Fermer xo‘jaligi', 'INN', 'Filial', 'Jami kirim (l)', 'Jami chiqim (l)', 'Qolgan yoqilg‘i (l)', 'Xizmat haqi (l)'],
    ...sortedFor('t-hold', cols, rows).map((r, i) => [i + 1, r.f.name, r.f.inn, r.branches.map(b => M.branches[b].name).join(', '), rnd(r.st.kirim), rnd(r.st.chiqim), rnd(r.st.stock), rnd(r.st.serv)]),
    ['', 'Jami', '', '', rnd(sums.k), rnd(sums.c), rnd(sums.st), rnd(sums.s)]
  ], [5, 38, 12, 24, 14, 14, 16, 14]);
}

/* ---------- 6.4 Harakatsiz qarzdorlar ---------- */
function renderInactive(view) {
  const key = 'inact';
  const F = S.filters[key] = S.filters[key] || { q: '', b: -1, min: '', days: 30 };
  const bi = +F.b;
  const { to } = S.range;
  const lastIdx = M.months.length - 1;
  const ref = to === lastIdx ? new Date() : monthEnd(M.months[to].year, M.months[to].month);
  ref.setHours(0, 0, 0, 0);
  const dataStart = monthStart(M.months[0].year, M.months[0].month);

  const all = farmerRows(bi).filter(r => r.st.debt > CFG.EPS && matchQ(r.f, F.q)).map(r => {
    const li = lastActivity(r.f, to);
    const lastDate = li >= 0 ? monthEnd(M.months[li].year, M.months[li].month) : null;
    const days = Math.max(0, Math.floor((ref - (lastDate || dataStart)) / DAY));
    return { ...r, li, lastDate, days };
  });
  const rows = all.filter(r => r.days >= F.days);
  const sumDebt = rows.reduce((a, r) => a + r.st.debt, 0);
  const cnt = (d) => all.filter(r => r.days >= d).length;

  const cols = [
    { key: 'n', label: '№', cls: 'idx', render: (_, i) => i + 1 },
    { key: 'name', label: 'Fermer xo‘jaligi', render: r => nameCell(r.f), sortVal: r => r.f.name },
    { key: 'br', label: 'Filiali', render: r => branchTags(r.branches) },
    { key: 'debt', label: 'Qarzdorlik (l)', num: 1, render: r => `<b class="neg">${fmt(r.st.debt)}</b>`, sortVal: r => r.st.debt },
    { key: 'debtS', label: 'Qarz (so‘m)', num: 1, render: r => nz(r.st.debtS) ? fmt(r.st.debtS) : '<span class="muted">—</span>', sortVal: r => r.st.debtS },
    { key: 'last', label: 'Oxirgi operatsiya', render: r => r.lastDate ? `${dateStr(r.lastDate)} <span class="muted">(${M.months[r.li].label})</span>` : `<span class="muted">Davrda harakat yo‘q</span>`, sortVal: r => r.lastDate ? +r.lastDate : 0 },
    { key: 'days', label: 'Harakatsiz', num: 1, render: r => `<span class="badge ${r.days >= 90 ? 'red' : r.days >= 60 ? 'amber' : 'blue'}">${r.lastDate ? '' : '≥ '}${fmt(r.days)} kun</span>`, sortVal: r => r.days }
  ];
  const t = dataTable('t-inact', cols, rows, {
    defaultSort: { key: 'debt', dir: 'desc' },
    footer: { name: `Jami: ${fmt(rows.length)} ta`, debt: fmt(sumDebt) },
    onRowClick: r => openAkt(r.f)
  });

  const chips = `<div class="chips" id="days-chips">${[30, 60, 90].map(d => `<button data-d="${d}" class="${F.days === d ? 'on' : ''}">${d}+ kun · ${cnt(d)}</button>`).join('')}</div>`;
  view.innerHTML = `
    ${pageHead('Harakatsiz qarzdorlar', `Qarzi bor, lekin belgilangan muddatda kirim, chiqim yoki to‘lov qilmagan fermerlar. Hisob sanasi: ${dateStr(ref)}.`,
      `<button class="btn" id="xl">${ic('excel')}Excel</button>`)}
    ${filterBar(key, { showMin: false, extra: `<div class="field" style="min-width:auto"><label>Harakatsizlik davri</label>${chips}</div>` })}
    <div class="note">${ic('info')}<div>Jadvalda kunlik sanalar yo‘q, shuning uchun oxirgi operatsiya sanasi deb harakat bo‘lgan oyning oxirgi kuni olinadi. "≥" belgisi fermer yuklangan ma'lumotlar davrida (${M.months[0].label} oyidan beri) umuman harakat qilmaganini bildiradi.</div></div>
    <div class="summary-strip">
      <div class="pill">Harakatsiz qarzdorlar:<b>${fmt(rows.length)}</b></div>
      <div class="pill">Ularning qarzi:<b class="neg">${fmt(sumDebt)} l</b></div>
    </div>
    ${rows.length ? t.html : `<div class="glass">${emptyState('Harakatsiz qarzdor yo‘q', `${F.days} kundan ortiq harakatsiz qarzdor topilmadi.`)}</div>`}`;

  const rerender = () => { renderInactive(view); hydrateIcons(view); };
  bindFilters(view, key, rerender);
  $$('#days-chips button', view).forEach(b => b.addEventListener('click', () => { F.days = +b.dataset.d; rerender(); }));
  t.bind(view, rerender);
  $('#xl', view).onclick = () => exportExcel(`Harakatsiz_qarzdorlar_${F.days}kun.xlsx`, 'Harakatsiz', [
    [CFG.company], [`Harakatsiz qarzdorlar (${F.days}+ kun) · hisob sanasi ${dateStr(ref)}`], [],
    ['№', 'Fermer xo‘jaligi', 'INN', 'Filial', 'Qarzdorlik (l)', 'Qarz (so‘m)', 'Oxirgi operatsiya', 'Harakatsiz (kun)'],
    ...sortedFor('t-inact', cols, rows).map((r, i) => [i + 1, r.f.name, r.f.inn, r.branches.map(b => M.branches[b].name).join(', '), rnd(r.st.debt), rnd(r.st.debtS), r.lastDate ? dateStr(r.lastDate) : 'Harakat yo‘q', r.days])
  ], [5, 38, 12, 24, 14, 14, 18, 16]);
}

/* ---------- 6.5 Aksverka ---------- */
function renderAksverka(view) {
  const A = S.filters.akt = S.filters.akt || {};
  if (A.from === undefined) { A.from = S.range.from; A.to = S.range.to; }
  if (A.b === undefined) A.b = -1;
  A.from = Math.min(A.from, M.months.length - 1); A.to = Math.min(A.to, M.months.length - 1);
  const f = A.farmer ? M.farmers.find(x => x.key === A.farmer) : null;
  const monthOpts = (sel) => M.months.map(m => `<option value="${m.idx}" ${m.idx === +sel ? 'selected' : ''}>${m.label}</option>`).join('');
  const sorted = [...M.farmers].sort((a, b) => a.name.localeCompare(b.name, 'ru'));

  view.innerHTML = `
    ${pageHead('Aksverka — o‘zaro hisob-kitoblarni solishtirish', 'Fermer xo‘jaligi va korxona o‘rtasidagi yoqilg‘i hamda pul hisob-kitoblari dalolatnomasi.')}
    <div class="akt-layout">
      <div class="glass akt-form">
        <div class="field"><label>Fermer xo‘jaligi</label>
          <input list="farmer-list" id="akt-farmer" placeholder="Nomi yoki INN bo‘yicha qidiring…" value="${f ? esc(f.name) : ''}">
          <datalist id="farmer-list">${sorted.map(x => `<option value="${esc(x.name)}">${x.inn ? 'INN ' + esc(x.inn) : ''}</option>`).join('')}</datalist>
        </div>
        <div class="field"><label>Filial</label><select id="akt-branch">${branchOptions(A.b, 'Barcha filiallar')}</select></div>
        <div class="row2">
          <div class="field"><label>Boshlanish</label><select id="akt-from">${monthOpts(A.from)}</select></div>
          <div class="field"><label>Tugash</label><select id="akt-to">${monthOpts(A.to)}</select></div>
        </div>
        <button class="btn btn-primary wide" id="akt-go">${ic('file')}Shakllantirish</button>
        <div class="muted" style="font-size:12px">Sana oraliqlari oy aniqligida: boshlanish oyining 1-kunidan tugash oyining oxirgi kunigacha.</div>
      </div>
      <div id="akt-out">${f ? '' : `<div class="glass">${emptyState('Fermer xo‘jaligini tanlang', 'Chapdagi maydonda nomi yoki INN bo‘yicha qidiring, so‘ng "Shakllantirish" tugmasini bosing.')}</div>`}</div>
    </div>`;

  const pick = () => {
    const v = normKey($('#akt-farmer', view).value);
    if (!v) return null;
    return M.farmers.find(x => normKey(x.name) === v) || M.farmers.find(x => x.inn === v) || M.farmers.find(x => matchQ(x, v));
  };
  const go = () => {
    const ff = pick();
    if (!ff) { toast('err', 'Fermer xo‘jaligi topilmadi. Ro‘yxatdan tanlang.'); return; }
    A.farmer = ff.key; A.b = +$('#akt-branch', view).value;
    A.from = +$('#akt-from', view).value; A.to = +$('#akt-to', view).value;
    if (A.from > A.to) [A.from, A.to] = [A.to, A.from];
    $('#akt-farmer', view).value = ff.name;
    $('#akt-from', view).value = A.from; $('#akt-to', view).value = A.to;
    renderAkt($('#akt-out', view), ff, A);
  };
  $('#akt-go', view).onclick = go;
  $('#akt-farmer', view).addEventListener('keydown', e => { if (e.key === 'Enter') go(); });
  $('#akt-farmer', view).addEventListener('change', () => { if (pick()) go(); });
  ['#akt-branch', '#akt-from', '#akt-to'].forEach(s => $(s, view).addEventListener('change', () => { if (A.farmer) go(); }));
  if (f) renderAkt($('#akt-out', view), f, A);
}

function aktData(f, A) {
  const bi = +A.b;
  const rows = [];
  for (let i = A.from; i <= A.to; i++) {
    const r = f.rows[i];
    const m = M.months[i];
    if (!r) { rows.push({ m, missing: true }); continue; }
    const kirim = bi < 0 ? r.kirimQ : r.kb[bi];
    const chiqim = bi < 0 ? r.chiqimQ : r.cb[bi];
    const serv = bi < 0 ? r.servQ : (r.chiqimQ > 0 ? r.servQ * r.cb[bi] / r.chiqimQ : 0);
    rows.push({ m, r, kirim, chiqim, serv });
  }
  const st = stats(f, A.from, A.to, bi);
  return { rows, st, bi };
}

function renderAkt(out, f, A) {
  const { rows, st, bi } = aktData(f, A);
  if (!st) { out.innerHTML = `<div class="glass">${emptyState('Ma\'lumot yo‘q', 'Tanlangan davrda bu fermer xo‘jaligi jadvalda uchramaydi.')}</div>`; hydrateIcons(out); return; }
  const a = M.months[A.from], b = M.months[A.to];
  const start = monthStart(a.year, a.month), end = monthEnd(b.year, b.month);
  const today = new Date();
  const bName = bi >= 0 ? M.branches[bi].name : 'Barcha filiallar';
  const d2 = (v) => fmt(v, 2);
  const cell = (v) => nz(v) ? d2(v) : '—';
  const real = rows.filter(x => !x.missing);
  const tot = real.reduce((t, x) => ({ k: t.k + x.kirim, c: t.c + x.chiqim, s: t.s + x.serv }), { k: 0, c: 0, s: 0 });

  // filiallar kesimi (faqat "barcha filiallar"da)
  const brRows = bi < 0 ? M.branches.map(br => {
    const k = real.reduce((t, x) => t + x.r.kb[br.i], 0), c = real.reduce((t, x) => t + x.r.cb[br.i], 0);
    return { br, k, c };
  }).filter(x => nz(x.k) || nz(x.c)) : [];

  const moneyAny = real.some(x => nz(x.r.debS) || nz(x.r.kreS) || nz(x.r.kirimS) || nz(x.r.cash) || nz(x.r.chiqimS) || nz(x.r.servS) || nz(x.r.debEndS) || nz(x.r.kreEndS));
  const mTot = real.reduce((t, x) => ({ n: t.n + x.r.kirimS, o: t.o + x.r.cash, t: t.t + x.r.chiqimS, v: t.v + x.r.servS }), { n: 0, o: 0, t: 0, v: 0 });

  const resultLine = (debt, stock, unit, dec) => debt > (unit === 'l' ? CFG.EPS : 0.5)
    ? `fermer xo‘jaligining korxona oldidagi qarzi: <b>${fmt(debt, dec)} ${unit}</b>`
    : stock > (unit === 'l' ? CFG.EPS : 0.5)
      ? `korxonaning fermer xo‘jaligi oldidagi qarzi (fermer qoldig‘i): <b>${fmt(stock, dec)} ${unit}</b>`
      : `o‘zaro qarzdorlik mavjud emas`;

  out.innerHTML = `
    <div class="akt-actions">
      <button class="btn btn-primary" id="akt-pdf">${ic('download')}PDF yuklab olish</button>
      <button class="btn" id="akt-print">${ic('print')}Chop etish</button>
      <button class="btn" id="akt-xl">${ic('excel')}Excel</button>
    </div>
    <div class="paper-wrap"><div class="paper" id="akt-doc">
      <div class="p-head">
        <div class="p-company">${CFG.company}<small>Yoqilg‘i saqlash va hisob-kitob xizmati</small></div>
        <div class="p-date">Tuzilgan sana: <b>${dateStr(today)}</b><br>Manba: ${esc(M.spreadsheet)}</div>
      </div>
      <h3>O‘ZARO HISOB-KITOBLARNI SOLISHTIRISH DALOLATNOMASI</h3>
      <div class="p-subtitle">(aksverka) ${dateStr(start)} – ${dateStr(end)}</div>
      <div class="p-info">
        <span>Korxona:</span><b>${CFG.company}</b>
        <span>Filial:</span><b>${esc(bName)}</b>
        <span>Fermer xo‘jaligi:</span><b>${esc(f.name)}</b>
        <span>INN:</span><b>${f.inn ? esc(f.inn) : '—'}</b>
        ${f.massiv ? `<span>Massiv:</span><b>${esc(f.massiv)}</b>` : ''}
        <span>Hisobot davri:</span><b>${dateStr(start)} – ${dateStr(end)}</b>
      </div>

      <h5>1. Yoqilg‘i hisob-kitobi (litr)</h5>
      <table>
        <thead>
          <tr><th rowspan="2">Davr</th><th colspan="2">Boshlang‘ich qoldiq</th><th rowspan="2">Kirim</th><th rowspan="2">Chiqim</th><th rowspan="2">Xizmat haqi</th><th colspan="2">Yakuniy qoldiq</th></tr>
          <tr><th>Qarz (debet)</th><th>Qoldiq (kredit)</th><th>Qarz (debet)</th><th>Qoldiq (kredit)</th></tr>
        </thead>
        <tbody>
          ${rows.map(x => x.missing
            ? `<tr><td>${x.m.label}</td><td colspan="7" style="text-align:center;color:#6b7280">Jadvalda bu oy uchun qator yo‘q</td></tr>`
            : `<tr><td>${x.m.label}</td><td class="n">${cell(x.r.debQ)}</td><td class="n">${cell(x.r.kreQ)}</td><td class="n">${cell(x.kirim)}</td><td class="n">${cell(x.chiqim)}</td><td class="n">${cell(x.serv)}</td><td class="n">${cell(x.r.debEndQ)}</td><td class="n">${cell(x.r.kreEndQ)}</td></tr>`).join('')}
        </tbody>
        <tfoot><tr><td>Jami</td><td class="n">${cell(st.openDebt)}</td><td class="n">${cell(st.openStock)}</td><td class="n">${cell(tot.k)}</td><td class="n">${cell(tot.c)}</td><td class="n">${cell(tot.s)}</td><td class="n">${cell(st.debt)}</td><td class="n">${cell(st.stock)}</td></tr></tfoot>
      </table>

      ${brRows.length ? `
      <h5>Filiallar kesimida (litr)</h5>
      <table>
        <thead><tr><th>Filial</th><th>Kirim</th><th>Chiqim</th></tr></thead>
        <tbody>${brRows.map(x => `<tr><td>${esc(x.br.name)}</td><td class="n">${cell(x.k)}</td><td class="n">${cell(x.c)}</td></tr>`).join('')}</tbody>
      </table>` : ''}

      ${moneyAny ? `
      <h5>2. Pul hisob-kitobi (so‘m)</h5>
      <table>
        <thead>
          <tr><th rowspan="2">Davr</th><th colspan="2">Boshlang‘ich saldo</th><th rowspan="2">Kirim summasi</th><th rowspan="2">Naqd pul</th><th rowspan="2">Realizatsiya</th><th rowspan="2">Xizmat summasi</th><th colspan="2">Yakuniy saldo</th></tr>
          <tr><th>Debet</th><th>Kredit</th><th>Debet</th><th>Kredit</th></tr>
        </thead>
        <tbody>
          ${real.map(x => `<tr><td>${x.m.label}</td><td class="n">${cell(x.r.debS)}</td><td class="n">${cell(x.r.kreS)}</td><td class="n">${cell(x.r.kirimS)}</td><td class="n">${cell(x.r.cash)}</td><td class="n">${cell(x.r.chiqimS)}</td><td class="n">${cell(x.r.servS)}</td><td class="n">${cell(x.r.debEndS)}</td><td class="n">${cell(x.r.kreEndS)}</td></tr>`).join('')}
        </tbody>
        <tfoot><tr><td>Jami</td><td class="n">${cell(real[0].r.debS)}</td><td class="n">${cell(real[0].r.kreS)}</td><td class="n">${cell(mTot.n)}</td><td class="n">${cell(mTot.o)}</td><td class="n">${cell(mTot.t)}</td><td class="n">${cell(mTot.v)}</td><td class="n">${cell(st.debtS)}</td><td class="n">${cell(st.stockS)}</td></tr></tfoot>
      </table>` : ''}

      <div class="p-result">
        <p><b>Natija:</b> ${dateStr(end)} holatiga ko‘ra</p>
        <p>• Yoqilg‘i bo‘yicha ${resultLine(st.debt, st.stock, 'l', 2)};</p>
        ${moneyAny ? `<p>• Pul bo‘yicha ${resultLine(st.debtS, st.stockS, 'so‘m', 0)}.</p>` : ''}
        <p>• Davr mobaynida hisoblangan saqlash xizmati haqi: <b>${d2(tot.s)} l</b>${nz(mTot.v) ? ` / <b>${fmt(mTot.v)} so‘m</b>` : ''}.</p>
      </div>
      ${bi >= 0 ? `<div class="p-foot-note">Izoh: kirim, chiqim va xizmat haqi "${esc(bName)}" filiali bo‘yicha ko‘rsatilgan. Boshlang‘ich va yakuniy qoldiqlar fermer xo‘jaligining barcha filiallar bo‘yicha umumiy qoldig‘i.</div>` : ''}
      <div class="p-foot-note">Xizmat haqi — Google Sheets'dagi "Ёкилги саклаш хизмати" ustuni (chiqim × 8%). Qoldiqlar jadval formulasi bo‘yicha: Kreditor + Kirim − Debitor − Chiqim − Xizmat haqi.</div>

      <div class="p-sign">
        <div><b>${CFG.company}</b><div class="line"></div><small>Rahbar (F.I.Sh., imzo) · M.O‘.</small></div>
        <div><b>${esc(f.name)}</b><div class="line"></div><small>Rahbar (F.I.Sh., imzo) · M.O‘.</small></div>
      </div>
    </div></div>`;
  hydrateIcons(out);

  const fileBase = `Aksverka_${f.name.replace(/[^\p{L}\p{N}]+/gu, '_').replace(/^_|_$/g, '')}_${a.key}_${b.key}`;
  $('#akt-print', out).onclick = () => window.print();
  $('#akt-pdf', out).onclick = async (e) => {
    const btn = e.currentTarget; btn.disabled = true;
    try {
      await loadScript(CFG.libs.pdf);
      await html2pdf().set({
        margin: [8, 8, 10, 8], filename: fileBase + '.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, backgroundColor: '#ffffff', useCORS: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
        pagebreak: { mode: ['css', 'legacy'] }
      }).from($('#akt-doc', out)).save();
      toast('ok', 'PDF yuklab olindi');
    } catch (err) { toast('err', err.message); } finally { btn.disabled = false; }
  };
  $('#akt-xl', out).onclick = () => {
    const aoa = [
      [CFG.company], ['O‘ZARO HISOB-KITOBLARNI SOLISHTIRISH DALOLATNOMASI (aksverka)'], [],
      ['Filial:', bName], ['Fermer xo‘jaligi:', f.name], ['INN:', f.inn || '—'],
      ['Hisobot davri:', `${dateStr(start)} – ${dateStr(end)}`], ['Tuzilgan sana:', dateStr(today)], [],
      ['1. Yoqilg‘i hisob-kitobi (litr)'],
      ['Davr', 'Boshl. qarz (debet)', 'Boshl. qoldiq (kredit)', 'Kirim', 'Chiqim', 'Xizmat haqi', 'Yak. qarz (debet)', 'Yak. qoldiq (kredit)'],
      ...rows.map(x => x.missing ? [x.m.label, 'Jadvalda qator yo‘q'] : [x.m.label, rnd(x.r.debQ), rnd(x.r.kreQ), rnd(x.kirim), rnd(x.chiqim), rnd(x.serv), rnd(x.r.debEndQ), rnd(x.r.kreEndQ)]),
      ['Jami', rnd(st.openDebt), rnd(st.openStock), rnd(tot.k), rnd(tot.c), rnd(tot.s), rnd(st.debt), rnd(st.stock)]
    ];
    if (brRows.length) aoa.push([], ['Filiallar kesimida (litr)'], ['Filial', 'Kirim', 'Chiqim'], ...brRows.map(x => [x.br.name, rnd(x.k), rnd(x.c)]));
    if (moneyAny) aoa.push([], ['2. Pul hisob-kitobi (so‘m)'],
      ['Davr', 'Boshl. debet', 'Boshl. kredit', 'Kirim summasi', 'Naqd pul', 'Realizatsiya', 'Xizmat summasi', 'Yak. debet', 'Yak. kredit'],
      ...real.map(x => [x.m.label, rnd(x.r.debS), rnd(x.r.kreS), rnd(x.r.kirimS), rnd(x.r.cash), rnd(x.r.chiqimS), rnd(x.r.servS), rnd(x.r.debEndS), rnd(x.r.kreEndS)]));
    aoa.push([], [`Natija (${dateStr(end)}):`, st.debt > CFG.EPS ? `Fermer qarzi ${rnd(st.debt)} l` : st.stock > CFG.EPS ? `Fermer qoldig‘i ${rnd(st.stock)} l` : 'O‘zaro qarzdorlik yo‘q']);
    exportExcel(fileBase + '.xlsx', 'Aksverka', aoa, [22, 18, 18, 14, 14, 14, 18, 18, 14]);
  };
}

/* ---------- 6.6 Filiallar — umumiy hisobot ---------- */
function renderBranchesOverview(view) {
  const { from, to } = S.range;
  const rows = farmerRows();
  const bStats = M.branches.map(b => {
    const rs = rows.map(r => stats(r.f, from, to, b.i)).filter(s => s && (nz(s.kirim) || nz(s.chiqim)));
    return { b, kirim: rs.reduce((a, s) => a + s.kirim, 0), chiqim: rs.reduce((a, s) => a + s.chiqim, 0), serv: rs.reduce((a, s) => a + s.serv, 0), farmers: rs.length };
  });
  const tot = bStats.reduce((t, x) => ({ kirim: t.kirim + x.kirim, chiqim: t.chiqim + x.chiqim, serv: t.serv + x.serv, farmers: rows.filter(r => r.st.active).length }), { kirim: 0, chiqim: 0, serv: 0 });
  const idx = []; for (let i = from; i <= to; i++) idx.push(i);
  const monthB = (i, bi, k) => M.farmers.reduce((a, f) => a + (f.rows[i] ? f.rows[i][k][bi] : 0), 0);

  const thead = `<tr><th rowspan="2">Oy</th>${M.branches.map(b => `<th colspan="2" class="num" style="text-align:center"><span class="tag"><span class="sw" style="background:${b.color}"></span>${esc(b.name)}</span></th>`).join('')}<th colspan="2" class="num" style="text-align:center">Jami</th></tr>
    <tr>${M.branches.map(() => '<th class="num">Kirim</th><th class="num">Chiqim</th>').join('')}<th class="num">Kirim</th><th class="num">Chiqim</th></tr>`;
  const body = idx.map(i => {
    const ks = M.branches.map(b => monthB(i, b.i, 'kb')), cs = M.branches.map(b => monthB(i, b.i, 'cb'));
    return `<tr><td>${M.months[i].label}</td>${M.branches.map((b, j) => `<td class="num">${fmt(ks[j])}</td><td class="num">${fmt(cs[j])}</td>`).join('')}<td class="num"><b>${fmt(ks.reduce((a, v) => a + v, 0))}</b></td><td class="num"><b>${fmt(cs.reduce((a, v) => a + v, 0))}</b></td></tr>`;
  }).join('');
  const foot = `<tr><td>Jami</td>${bStats.map(x => `<td class="num">${fmt(x.kirim)}</td><td class="num">${fmt(x.chiqim)}</td>`).join('')}<td class="num">${fmt(tot.kirim)}</td><td class="num">${fmt(tot.chiqim)}</td></tr>`;

  view.innerHTML = `
    ${pageHead('Filiallar — umumiy hisobot', `Barcha filiallar bo‘yicha kirim, chiqim va xizmat haqi · ${rangeLabel()} · litrda`, `<button class="btn" id="xl">${ic('excel')}Excel</button>`)}
    <div class="branch-grid">
      ${bStats.map(x => branchCard(x, tot)).join('')}
      <div class="glass branch-card total">
        <div class="bc-head"><div class="bc-name">${ic('building')} Barcha filiallar</div></div>
        <div class="bc-rows">
          <div><span>Kirim</span><b>${fmt(tot.kirim)}</b></div><div><span>Chiqim</span><b>${fmt(tot.chiqim)}</b></div>
          <div><span>Xizmat haqi</span><b>${fmt(tot.serv)}</b></div><div><span>Faol fermerlar</span><b>${fmt(tot.farmers)}</b></div>
        </div>
      </div>
    </div>
    <div class="section-title"><h3>Oylar bo‘yicha chiqim (filiallar kesimida)</h3><span>litrda</span></div>
    <div class="glass chart-card"><div class="chart-box"><canvas id="ch-bm"></canvas></div></div>
    <div class="section-title"><h3>Oylik jadval</h3><span>litrda</span></div>
    <div class="glass table-card"><div class="table-scroll"><table class="data"><thead>${thead}</thead><tbody>${body}</tbody><tfoot>${foot}</tfoot></table></div></div>
    <div class="note" style="margin-top:14px">${ic('info')}<div>Filiallar bo‘yicha xizmat haqi har bir fermerning oylik xizmat haqini uning filiallardagi chiqim hajmiga mutanosib taqsimlash orqali hisoblanadi (jadvalda 8% bo‘lsa, filial chiqimi × 8%).</div></div>`;

  chart($('#ch-bm'), {
    type: 'bar',
    data: { labels: idx.map(i => M.months[i].short), datasets: M.branches.map(b => ({ label: b.name, data: idx.map(i => monthB(i, b.i, 'cb')), backgroundColor: b.color + 'cc', borderRadius: 4, maxBarThickness: 34 })) },
    options: { scales: { x: { stacked: true, grid: { display: false } }, y: { ...axisL, stacked: true } }, plugins: { legend: { position: 'top', align: 'end' } } }
  });
  $('#xl', view).onclick = () => exportExcel(`Filiallar_${M.months[from].key}_${M.months[to].key}.xlsx`, 'Filiallar', [
    [CFG.company], [`Filiallar bo‘yicha umumiy hisobot · ${rangeLabel()} (litr)`], [],
    ['Oy', ...M.branches.flatMap(b => [`${b.name} kirim`, `${b.name} chiqim`]), 'Jami kirim', 'Jami chiqim'],
    ...idx.map(i => { const ks = M.branches.map(b => rnd(monthB(i, b.i, 'kb'))), cs = M.branches.map(b => rnd(monthB(i, b.i, 'cb'))); return [M.months[i].label, ...ks.flatMap((k, j) => [k, cs[j]]), rnd(ks.reduce((a, v) => a + v, 0)), rnd(cs.reduce((a, v) => a + v, 0))]; }),
    ['Jami', ...bStats.flatMap(x => [rnd(x.kirim), rnd(x.chiqim)]), rnd(tot.kirim), rnd(tot.chiqim)],
    ['Xizmat haqi', ...bStats.flatMap(x => [rnd(x.serv), '']), rnd(tot.serv), '']
  ], [16, ...M.branches.flatMap(() => [16, 16]), 14, 14]);
}

/* ---------- 6.7 Bitta filial sahifasi ---------- */
function renderBranch(view, arg) {
  const bi = +arg;
  const br = M.branches[bi];
  if (!br) { location.hash = '#/filiallar'; return; }
  const key = 'br' + bi;
  const F = S.filters[key] = S.filters[key] || { q: '', b: bi, min: '', tab: 'all' };
  const { from, to } = S.range;
  const all = farmerRows(bi);
  const inRange = all.filter(r => nz(r.st.kirim) || nz(r.st.chiqim));
  const debtors = all.filter(r => r.st.debt > CFG.EPS);
  const holders = all.filter(r => r.st.stock > CFG.EPS);
  const sum = (arr, fn) => arr.reduce((a, r) => a + fn(r), 0);
  const kir = sum(all, r => r.st.kirim), chq = sum(all, r => r.st.chiqim), srv = sum(all, r => r.st.serv);

  const tabRows = (F.tab === 'debt' ? debtors : F.tab === 'stock' ? holders : all).filter(r => matchQ(r.f, F.q));
  const cols = [
    { key: 'n', label: '№', cls: 'idx', render: (_, i) => i + 1 },
    { key: 'name', label: 'Fermer xo‘jaligi', render: r => nameCell(r.f), sortVal: r => r.f.name },
    { key: 'k', label: 'Kirim', num: 1, render: r => fmt(r.st.kirim), sortVal: r => r.st.kirim },
    { key: 'c', label: 'Chiqim', num: 1, render: r => fmt(r.st.chiqim), sortVal: r => r.st.chiqim },
    { key: 's', label: 'Xizmat haqi', num: 1, render: r => fmt(r.st.serv), sortVal: r => r.st.serv },
    { key: 'st', label: 'Qoldiq', num: 1, render: r => r.st.stock > CFG.EPS ? `<span class="pos">${fmt(r.st.stock)}</span>` : '<span class="muted">—</span>', sortVal: r => r.st.stock },
    { key: 'debt', label: 'Qarzdorlik', num: 1, render: r => r.st.debt > CFG.EPS ? `<b class="neg">${fmt(r.st.debt)}</b>` : '<span class="muted">—</span>', sortVal: r => r.st.debt }
  ];
  const t = dataTable('t-br' + bi + F.tab, cols, tabRows, {
    defaultSort: { key: F.tab === 'debt' ? 'debt' : F.tab === 'stock' ? 'st' : 'c', dir: 'desc' },
    footer: { name: `Jami: ${fmt(tabRows.length)} ta`, k: fmt(sum(tabRows, r => r.st.kirim)), c: fmt(sum(tabRows, r => r.st.chiqim)), s: fmt(sum(tabRows, r => r.st.serv)), st: fmt(sum(tabRows, r => Math.max(0, r.st.stock))), debt: fmt(sum(tabRows, r => Math.max(0, r.st.debt))) },
    onRowClick: r => { S.filters.akt = { farmer: r.f.key, b: bi, from: S.range.from, to: S.range.to }; location.hash = '#/aksverka'; }
  });
  const idx = []; for (let i = from; i <= to; i++) idx.push(i);
  const monthB = (i, k) => M.farmers.reduce((a, f) => a + (f.rows[i] ? f.rows[i][k][bi] : 0), 0);

  view.innerHTML = `
    ${pageHead(`<span class="tag" style="font-size:14px;padding:4px 10px;margin-right:8px"><span class="sw" style="background:${br.color};width:10px;height:10px"></span>Filial</span>${esc(br.name)}`,
      `Faqat shu filialga tegishli ma'lumotlar · ${rangeLabel()} · litrda. Qoldiq va qarzdorlik — fermerning umumiy ko‘rsatkichi.`,
      `<button class="btn" id="xl">${ic('excel')}Excel</button>`)}
    <div class="kpi-grid">
      ${kpi('Kirim', fmt(kir), 'l', rangeLabel(), 'arrowDown')}
      ${kpi('Chiqim', fmt(chq), 'l', rangeLabel(), 'arrowUp', 'cyan')}
      ${kpi('Xizmat haqi (8%)', fmt(srv), 'l', 'Filial chiqimiga mutanosib', 'percent', 'amber')}
      ${kpi('Qarzdorlar', fmt(sum(debtors, r => r.st.debt)), 'l', `${fmt(debtors.length)} ta fermer`, 'debt', 'red')}
      ${kpi('Yoqilg‘i qoldiqlari', fmt(sum(holders, r => r.st.stock)), 'l', `${fmt(holders.length)} ta fermer · faol: ${fmt(inRange.length)} / ${fmt(all.length)}`, 'drop', 'green')}
    </div>
    <div class="glass chart-card" style="margin-bottom:18px"><h4>Oylar bo‘yicha kirim va chiqim</h4><p>${esc(br.name)} filiali, litrda</p><div class="chart-box" style="height:240px"><canvas id="ch-b"></canvas></div></div>
    <div class="tabs" id="br-tabs">
      <button data-t="all" class="${F.tab === 'all' ? 'on' : ''}">Fermerlar ro‘yxati · ${all.length}</button>
      <button data-t="debt" class="${F.tab === 'debt' ? 'on' : ''}">Qarzdorlar · ${debtors.length}</button>
      <button data-t="stock" class="${F.tab === 'stock' ? 'on' : ''}">Yoqilg‘i qoldiqlari · ${holders.length}</button>
    </div>
    <div class="glass filters"><div class="field grow"><label>Qidiruv</label><input type="search" data-f="q" placeholder="Fermer xo‘jaligi nomi yoki INN…" value="${esc(F.q)}"></div></div>
    ${tabRows.length ? t.html : `<div class="glass">${emptyState('Ma\'lumot yo‘q', 'Bu bo‘limda tanlangan davr uchun fermer topilmadi.')}</div>`}`;

  const ctx = $('#ch-b').getContext('2d');
  chart(ctx, {
    type: 'bar',
    data: { labels: idx.map(i => M.months[i].short), datasets: [
      { label: 'Kirim', data: idx.map(i => monthB(i, 'kb')), backgroundColor: grad(ctx, br.color, br.color + '33'), borderRadius: 6, maxBarThickness: 26 },
      { label: 'Chiqim', data: idx.map(i => monthB(i, 'cb')), backgroundColor: 'rgba(238,243,252,.55)', borderRadius: 6, maxBarThickness: 26 }
    ] },
    options: { scales: { y: axisL, x: { grid: { display: false } } }, plugins: { legend: { position: 'top', align: 'end' } } }
  });

  const rerender = () => { renderBranch(view, arg); hydrateIcons(view); };
  bindFilters(view, key, rerender);
  $$('#br-tabs button', view).forEach(b => b.addEventListener('click', () => { F.tab = b.dataset.t; rerender(); }));
  t.bind(view, rerender);
  $('#xl', view).onclick = () => exportExcel(`Filial_${br.name}_${M.months[to].key}.xlsx`, br.name, [
    [CFG.company], [`Filial: ${br.name} · ${rangeLabel()} (litr)`], [],
    ['№', 'Fermer xo‘jaligi', 'INN', 'Kirim', 'Chiqim', 'Xizmat haqi', 'Qoldiq', 'Qarzdorlik'],
    ...sortedFor('t-br' + bi + F.tab, cols, tabRows).map((r, i) => [i + 1, r.f.name, r.f.inn, rnd(r.st.kirim), rnd(r.st.chiqim), rnd(r.st.serv), rnd(r.st.stock), rnd(r.st.debt)])
  ], [5, 38, 12, 12, 12, 12, 12, 12]);
}

/* ---------- Xato sahifasi ---------- */
function renderFatal(msg) {
  const view = $('#view');
  view.innerHTML = `<div class="glass error-card">
    <div class="eico">${ic('wifiOff')}</div>
    <h3>Ma'lumotlarni yuklab bo‘lmadi</h3>
    <p>${navigator.onLine ? 'Google Sheets bilan bog‘lanishda xatolik yuz berdi.' : 'Internet aloqasi yo‘q. Ulanishni tekshirib, qayta urinib ko‘ring.'}</p>
    <code>${esc(msg)}</code>
    <button class="btn btn-primary" id="retry">${ic('refresh')}Qayta urinish</button>
  </div>`;
  $('#retry').onclick = () => fetchData();
}

/* ============================================================
   7. GLOBAL QIDIRUV
   ============================================================ */
function initSearch() {
  const input = $('#global-search'), box = $('#search-results');
  let hl = -1, items = [];
  const close = () => { box.classList.remove('open'); hl = -1; };
  const draw = () => {
    const q = normKey(input.value);
    if (!M || q.length < 2) { close(); return; }
    items = M.farmers.filter(f => matchQ(f, q)).slice(0, 12);
    const to = S.range.to;
    box.innerHTML = items.length ? items.map((f, i) => {
      const r = f.rows[to];
      const bal = r ? r.kreEndQ - r.debEndQ : null;
      return `<button data-i="${i}" class="${i === hl ? 'hl' : ''}"><span>${esc(f.name)}<br><small>${f.inn ? 'INN ' + esc(f.inn) : ''}</small></span><small class="${bal < -CFG.EPS ? 'neg' : bal > CFG.EPS ? 'pos' : ''}">${bal === null ? '' : fmtSigned(bal) + ' l'}</small></button>`;
    }).join('') : `<div class="empty">Hech narsa topilmadi</div>`;
    box.classList.add('open');
    $$('button', box).forEach(b => b.addEventListener('mousedown', (e) => { e.preventDefault(); choose(items[+b.dataset.i]); }));
  };
  const choose = (f) => { if (!f) return; input.value = ''; close(); openAkt(f); };
  input.addEventListener('input', () => { hl = -1; draw(); });
  input.addEventListener('focus', draw);
  input.addEventListener('blur', () => setTimeout(close, 120));
  input.addEventListener('keydown', (e) => {
    if (!box.classList.contains('open')) return;
    if (e.key === 'ArrowDown') { hl = Math.min(items.length - 1, hl + 1); draw(); e.preventDefault(); }
    if (e.key === 'ArrowUp') { hl = Math.max(0, hl - 1); draw(); e.preventDefault(); }
    if (e.key === 'Enter') choose(items[Math.max(0, hl)]);
    if (e.key === 'Escape') close();
  });
}

/* ============================================================
   8. ISHGA TUSHIRISH
   ============================================================ */
function tickClock() {
  const d = new Date();
  const days = ['Yakshanba', 'Dushanba', 'Seshanba', 'Chorshanba', 'Payshanba', 'Juma', 'Shanba'];
  $('#today').textContent = `${days[d.getDay()]}, ${d.getDate()}-${UZ_MONTHS[d.getMonth()].toLowerCase()} ${d.getFullYear()}`;
}

function init() {
  hydrateIcons();
  tickClock(); setInterval(tickClock, 60000);
  initSearch();

  $('#refresh-btn').addEventListener('click', () => fetchData({ silent: !!M }));
  $('#menu-btn').addEventListener('click', () => document.body.classList.toggle('nav-open'));
  $('#scrim').addEventListener('click', () => document.body.classList.remove('nav-open'));
  window.addEventListener('hashchange', route);

  $('#year-select').addEventListener('change', (e) => {
    const y = +e.target.value;
    if (!y) { setQuickRange('all'); return; }
    const idx = M.months.filter(m => m.year === y).map(m => m.idx);
    S.range = { from: idx[0], to: idx[idx.length - 1] };
    syncPeriodBar(); route();
  });
  $('#from-select').addEventListener('change', (e) => { S.range.from = +e.target.value; if (S.range.from > S.range.to) S.range.to = S.range.from; syncPeriodBar(); route(); });
  $('#to-select').addEventListener('change', (e) => { S.range.to = +e.target.value; if (S.range.to < S.range.from) S.range.from = S.range.to; syncPeriodBar(); route(); });
  $$('#quick-periods button').forEach(b => b.addEventListener('click', () => setQuickRange(b.dataset.q)));

  $('#login-form').addEventListener('submit', (e) => {
    e.preventDefault();
    sSet(CFG.passKey, $('#login-pass').value);
    $('#login-modal').hidden = true;
    fetchData();
  });

  window.addEventListener('offline', () => { setStatus('off', 'Oflayn'); toast('warn', 'Internet aloqasi uzildi. Ko‘rsatilayotgan ma\'lumotlar yangilanmaydi.'); });
  window.addEventListener('online', () => { toast('ok', 'Internet aloqasi tiklandi'); fetchData({ silent: !!M }); });

  // Sahifa qayta ochilganda yoki fonga o‘tib qaytganda (5 daqiqadan keyin) avtomatik yangilanadi
  let hiddenAt = 0;
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) hiddenAt = Date.now();
    else if (M && hiddenAt && Date.now() - hiddenAt > 5 * 60000) fetchData({ silent: true });
  });

  fetchData();
}

document.addEventListener('DOMContentLoaded', init);
