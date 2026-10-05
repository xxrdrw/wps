
/* ===== 换电站调试检查表 App 核心逻辑 ===== */
'use strict';

/* ---------- 核心库按需加载（解析/导出时才从 CDN 加载，首屏零依赖秒开） ---------- */
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error('组件加载失败：' + src + '，请检查网络后重试'));
    document.head.appendChild(s);
  });
}
const _libLoading = {};
async function ensureLib(kind) {
  if (kind === 'xlsx' && typeof XLSX !== 'undefined') return;
  if (kind === 'exceljs' && typeof ExcelJS !== 'undefined') return;
  if (_libLoading[kind]) return _libLoading[kind];
  const url = kind === 'xlsx'
    ? 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
    : 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
  _libLoading[kind] = loadScript(url).then(() => {
    if ((kind === 'xlsx' && typeof XLSX === 'undefined') || (kind === 'exceljs' && typeof ExcelJS === 'undefined')) {
      throw new Error(kind + ' 加载后仍不可用');
    }
  }).catch((e) => { delete _libLoading[kind]; throw e; });
  return _libLoading[kind];
}

/* ---------- 常量 ---------- */
const BUSY_SHEETS = ['调试检查清单', '调试功能检查单（厂外）', '电气特殊特性检查清单', '抬车臂'];
const MAX_IMAGES = 20;
const IMG_MAX_EDGE = 1600;      // 拍照最长边
const IMG_MAX_BYTES = 500 * 1024; // 单张≤500KB
const SALT = 'dbs_station_check_2026';

/* ---------- 工具函数 ---------- */
function pad2(n) { return String(n).padStart(2, '0'); }
function fmtDateTime(ts) {
  if (!ts) return '';
  const d = new Date(ts);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
}
function fmtDateFile(ts) {
  const d = new Date(ts);
  return `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}_${pad2(d.getHours())}${pad2(d.getMinutes())}${pad2(d.getSeconds())}`;
}
function colLetter(idx) { // 0-based -> A, B...
  let s = '';
  idx += 1;
  while (idx > 0) { const m = (idx - 1) % 26; s = String.fromCharCode(65 + m) + s; idx = Math.floor((idx - 1) / 26); }
  return s;
}
function randStr(n) {
  const c = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let s = '';
  for (let i = 0; i < n; i++) s += c[Math.floor(Math.random() * c.length)];
  return s;
}
async function genAntiFake(stationId, itemId, ts, inspector, random) {
  const raw = `${stationId}|${itemId}|${ts}|${inspector}|${random}|${SALT}`;
  if (crypto && crypto.subtle) {
    const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(raw));
    return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, '0')).join('').toUpperCase().slice(0, 12);
  }
  return sha256(raw).toUpperCase().slice(0, 12); // 极旧环境降级（HTTPS 下不会走到）
}

/* ---------- 防伪签名层（设备端密码学签名，离线可用） ---------- */
let sigKey = null;
async function hashBlob(blob) {
  if (crypto && crypto.subtle) {
    const buf = await blob.arrayBuffer();
    const digest = await crypto.subtle.digest('SHA-256', buf);
    return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, '0')).join('');
  }
  // 降级：js-sha256
  const buf = await blob.arrayBuffer();
  return sha256(new Uint8Array(buf));
}
async function ensureSigKey() {
  if (sigKey) return sigKey;
  try {
    const stored = await Store.idbGet('keys', 'sig-keypair');
    if (stored && stored.privateKey) {
      sigKey = await crypto.subtle.importKey('jwk', stored.privateKey,
        { name: 'ECDSA', namedCurve: 'P-256' }, false, ['sign']);
      sigPubJwk = stored.publicJwk;
      return sigKey;
    }
  } catch (e) { /* 继续生成新密钥 */ }
  try {
    const kp = await crypto.subtle.generateKey({ name: 'ECDSA', namedCurve: 'P-256' }, true, ['sign', 'verify']);
    sigKey = kp.privateKey;
    sigPubJwk = await crypto.subtle.exportKey('jwk', kp.publicKey);
    const privJwk = await crypto.subtle.exportKey('jwk', kp.privateKey);
    await Store.idbPut('keys', 'sig-keypair', { privateKey: privJwk, publicJwk: sigPubJwk });
  } catch (e) { sigKey = null; }
  return sigKey;
}
let sigPubJwk = null;
async function signPhotoPayload(payload) {
  const key = await ensureSigKey();
  if (!key) return '';
  const sig = await crypto.subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, key, new TextEncoder().encode(payload));
  const bytes = new Uint8Array(sig);
  let bin = '';
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}
/** 验真：重算哈希 + 验签，返回验证报告 */
async function verifyPhoto(imgId) {
  const meta = Store.imgMeta[imgId];
  if (!meta || !meta.hash) return { ok: false, reason: '该照片没有防伪签名（旧版本拍摄，请重拍）' };
  const blob = await Store.idbGet('images', imgId);
  if (!blob) return { ok: false, reason: '照片数据缺失，无法验证' };
  const hash = await hashBlob(blob);
  if (hash !== meta.hash) return { ok: false, reason: '内容被篡改：照片像素与拍摄时不一致（哈希不匹配）' };
  if (!meta.sig || !meta.pub) return { ok: false, reason: '缺少签名数据，无法完成密码学校验' };
  try {
    const pub = await crypto.subtle.importKey('jwk', meta.pub,
      { name: 'ECDSA', namedCurve: 'P-256' }, false, ['verify']);
    const payload = `${hash}|${meta.taskId}|${meta.itemId}|${meta.createdAt}|${meta.code}`;
    const sigBytes = Uint8Array.from(atob(meta.sig), c => c.charCodeAt(0));
    const good = await crypto.subtle.verify(
      { name: 'ECDSA', hash: 'SHA-256' }, pub, sigBytes, new TextEncoder().encode(payload));
    if (!good) return { ok: false, reason: '防伪签名校验失败（元数据可能被篡改）' };
  } catch (e) {
    return { ok: false, reason: '签名校验异常：' + (e.message || e) };
  }
  return { ok: true, meta };
}
function calcGrid(n) {
  if (n <= 1) return { cols: 1, rows: 1 };
  if (n === 2) return { cols: 2, rows: 1 };
  if (n === 3) return { cols: 3, rows: 1 };
  if (n === 4) return { cols: 2, rows: 2 };
  if (n <= 6) return { cols: 3, rows: 2 };
  if (n <= 8) return { cols: 4, rows: 2 };
  if (n === 9) return { cols: 3, rows: 3 };
  if (n === 10) return { cols: 5, rows: 2 };
  const cols = Math.ceil(Math.sqrt(n));
  return { cols, rows: Math.ceil(n / cols) };
}
function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function $(id) { return document.getElementById(id); }

/* ---------- 存储层 ---------- */
const Store = {
  task: null, items: [], imgMeta: {}, amapKey: '', photoEdge: 1600,
  sheetSettings: {}, sheetCols: {},
  async init() {
    try { this.task = JSON.parse(localStorage.getItem('sts.task')); } catch (e) { this.task = null; }
    try { this.items = JSON.parse(localStorage.getItem('sts.items')) || []; } catch (e) { this.items = []; }
    try { this.imgMeta = JSON.parse(localStorage.getItem('sts.imgMeta')) || {}; } catch (e) { this.imgMeta = {}; }
    try { this.amapKey = localStorage.getItem('sts.amapKey') || ''; } catch (e) { this.amapKey = ''; }
    try { this.photoEdge = Number(localStorage.getItem('sts.photoEdge')) || 1600; } catch (e) { this.photoEdge = 1600; }
    try { this.sheetSettings = JSON.parse(localStorage.getItem('sts.sheetSettings')) || {}; } catch (e) { this.sheetSettings = {}; }
    this.db = await this._openDB();
  },
  _openDB() {
    return new Promise((res) => {
      try {
        if (!window.indexedDB) { res(null); return; }
        const rq = indexedDB.open('sts-db', 1);
        rq.onupgradeneeded = () => {
          const db = rq.result;
          if (!db.objectStoreNames.contains('images')) db.createObjectStore('images');
          if (!db.objectStoreNames.contains('source')) db.createObjectStore('source');
        };
        rq.onsuccess = () => res(rq.result);
        rq.onerror = () => res(null);   // 数据库不可用时降级为内存存储，不阻断使用
      } catch (e) { res(null); }
    });
  },
  _memStore: {},
  _tx(store, mode) { return this.db.transaction(store, mode).objectStore(store); },
  async idbPut(store, key, val) {
    if (!this.db) { (this._memStore[store] = this._memStore[store] || {})[key] = val; return; }
    return new Promise((res, rej) => {
      const rq = this._tx(store, 'readwrite').put(val, key);
      rq.onsuccess = () => res();
      rq.onerror = () => rej(rq.error);
    });
  },
  async idbGet(store, key) {
    if (!this.db) { const m = this._memStore[store]; return m ? m[key] : undefined; }
    return new Promise((res, rej) => {
      const rq = this._tx(store, 'readonly').get(key);
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => rej(rq.error);
    });
  },
  async idbDel(store, key) {
    if (!this.db) { const m = this._memStore[store]; if (m) delete m[key]; return; }
    return new Promise((res, rej) => {
      const rq = this._tx(store, 'readwrite').delete(key);
      rq.onsuccess = () => res();
      rq.onerror = () => rej(rq.error);
    });
  },
  save() {
    localStorage.setItem('sts.task', JSON.stringify(this.task));
    localStorage.setItem('sts.items', JSON.stringify(this.items));
    localStorage.setItem('sts.imgMeta', JSON.stringify(this.imgMeta));
    try { localStorage.setItem('sts.amapKey', this.amapKey || ''); } catch (e) {}
    try { localStorage.setItem('sts.photoEdge', String(this.photoEdge || 1600)); } catch (e) {}
    try { localStorage.setItem('sts.sheetSettings', JSON.stringify(this.sheetSettings || {})); } catch (e) {}
  },
  clearTask() {
    this.task = null; this.items = []; this.imgMeta = {};
    localStorage.removeItem('sts.task');
    localStorage.removeItem('sts.items');
    localStorage.removeItem('sts.imgMeta');
    try { this._tx('images', 'readwrite').clear(); } catch (e) {}
    try { this._tx('source', 'readwrite').clear(); } catch (e) {}
  },
  progress() {
    const done = this.items.filter(it => it.status === 'done').length;
    return { done, total: this.items.length, pct: this.items.length ? Math.round(done / this.items.length * 100) : 0 };
  }
};

/* ---------- 页面导航 ---------- */
const VIEWS = ['view-login', 'view-home', 'view-list', 'view-detail', 'view-camera', 'view-export', 'view-settings'];
const TITLES = {
  'view-login': '登录',
  'view-home': 'wps表格插图手机版',
  'view-list': '检查项列表',
  'view-detail': '检查项详情',
  'view-camera': '水印相机',
  'view-export': '导出 Excel',
  'view-settings': '设置'
};
let currentView = 'view-home';
let currentItemId = null;
function showView(name) {
  VIEWS.forEach(v => { $(v).style.display = 'none'; });
  $(name).style.display = 'block';
  currentView = name;
  $('topTitle').textContent = TITLES[name] || '换电站调试检查表';
  $('btnBack').style.display = (name === 'view-home' || name === 'view-camera' || name === 'view-login') ? 'none' : 'block';
}

/* ---------- 设置页：高德 Key / 水印模板 / 照片尺寸 ---------- */
function syncSettingsUI() {
  try {
    document.querySelectorAll('#photoEdgeChips .chip').forEach(c => {
      c.classList.toggle('active', Number(c.dataset.edge) === (Store.photoEdge || 1600));
    });
  } catch (e) {}
}
/* ---------- 检查表设置（每表：序号起始 / 照片插入列） ---------- */
function sheetSeqStart(base) {
  const st = Store.sheetSettings && Store.sheetSettings[base];
  const n = st && st.seqStart ? Number(st.seqStart) : 0;
  return n >= 1 ? n : 5;
}
function sheetPhotoCol(base) {
  const st = Store.sheetSettings && Store.sheetSettings[base];
  const n = st && st.photoCol ? Number(st.photoCol) : 0;
  return n >= 1 ? n : 0;
}
function colFromInput(s) {
  s = String(s == null ? '' : s).trim().toUpperCase();
  if (!s) return 0;
  if (/^\d+$/.test(s)) { const n = parseInt(s, 10); return (n >= 1 && n <= 200) ? n : 0; }
  if (/^[A-Z]{1,2}$/.test(s)) { let n = 0; for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  return 0;
}
function renderSheetSettings() {
  const card = $('sheetSettingsCard');
  if (!card) return;
  if (!Store.items.length) { card.style.display = 'none'; return; }
  card.style.display = 'block';
  const bases = [];
  Store.items.forEach(x => { if (bases.indexOf(x.sheetName) < 0) bases.push(x.sheetName); });
  const box = $('sheetSettingsBox');
  box.innerHTML = bases.map(base => {
    const st = Store.sheetSettings && Store.sheetSettings[base];
    const seq = (st && st.seqStart) ? st.seqStart : 5;
    const colN = (st && st.photoCol) ? st.photoCol : ((Store.sheetCols[base] && Store.sheetCols[base].image) || 0);
    const colTxt = colN ? colN : '';
    return `<div class="ss-row">
      <div class="ss-name">${esc(base)}</div>
      <label class="ss-field">序号从<input type="number" min="1" max="200" class="ss-seq" data-base="${esc(base)}" value="${seq}">开始</label>
      <label class="ss-field">照片插入<input type="text" class="ss-col" data-base="${esc(base)}" value="${colTxt}" placeholder="列字母/数字 如 G 或 7">列</label>
    </div>`;
  }).join('');
}
function saveSheetSettings() {
  if (!Store.items.length) { Bridge.toast('请先导入检查表'); return; }
  const bases = [];
  Store.items.forEach(x => { if (bases.indexOf(x.sheetName) < 0) bases.push(x.sheetName); });
  let bad = null;
  bases.forEach(base => {
    const seqEl = document.querySelector(`.ss-seq[data-base="${CSS_ESC(base)}"]`);
    const colEl = document.querySelector(`.ss-col[data-base="${CSS_ESC(base)}"]`);
    const seq = seqEl ? parseInt(seqEl.value, 10) : 5;
    const col = colEl ? colFromInput(colEl.value) : 0;
    if (!(seq >= 1 && seq <= 200)) { bad = base + '：序号需为 1-200 的整数'; return; }
    if (colEl && colEl.value.trim() && !col) { bad = base + '：照片列格式不对（如 G 或 7）'; return; }
    Store.sheetSettings[base] = { seqStart: seq, photoCol: col };
  });
  if (bad) { alert(bad); return; }
  Store.save();
  renderSheetSettings();
  Bridge.toast('已保存。序号起始下次导入表格后生效，照片列对导出立即生效');
}
function CSS_ESC(s) { return String(s).replace(/"/g, '\\"').replace(/\s/g, '_'); }
function openSettings() {
  $('inpAmapKey').value = Store.amapKey || '';
  const r = $('keyTestResult');
  r.style.display = 'none';
  r.textContent = '';
  syncSettingsUI();
  renderSheetSettings();
  showView('view-settings');
}
function saveKey() {
  const k = $('inpAmapKey').value.trim();
  if (!k) { alert('请输入高德 Key（留空无法保存）'); return; }
  Store.amapKey = k;
  Store.save();
  Bridge.toast('已保存，拍照将使用新 Key');
  showView('view-home');
}
$('btnSaveSheetSettings').onclick = saveSheetSettings;
async function testKey() {
  const k = $('inpAmapKey').value.trim();
  if (!k) { alert('请先输入 Key 再校验'); return; }
  const r = $('keyTestResult');
  r.style.display = 'block';
  r.textContent = '正在校验 Key…（联网请求高德，约需几秒）';
  r.className = 'hint key-test-result';
  let obj;
  try {
    obj = await new Promise((res) => {
      const timer = setTimeout(() => { window.__amapKeyTestResult = null; res({ ok: false, msg: '校验超时（请检查网络）' }); }, 15000);
      window.__amapKeyTestResult = (o) => { clearTimeout(timer); res(o); };
      if (window.stsBridge) { window.stsBridge.testAmapKey(k); }
      else { res({ ok: false, msg: '当前环境非安卓 App，跳过校验' }); }
    });
  } catch (e) { obj = { ok: false, msg: '校验异常：' + (e.message || e) }; }
  r.textContent = obj && obj.ok
    ? 'Key 有效，可正常解析地址：' + (obj.addr || '（返回了地址）')
    : 'Key 校验未通过：' + ((obj && obj.msg) || '未知错误');
  r.className = obj && obj.ok ? 'hint key-test-result ok' : 'hint key-test-result bad';
}

function base64ToArrayBuffer(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}
function recordError(msg, stack) {
  try {
    localStorage.setItem('sts.lastError', JSON.stringify({ time: Date.now(), msg: msg || '', stack: stack || '' }));
  } catch (e) { /* ignore */ }
}
function readLastError() {
  try { return JSON.parse(localStorage.getItem('sts.lastError') || 'null'); } catch (e) { return null; }
}

/* ---------- 桥接层（安卓 / 浏览器回退） ---------- */
const Bridge = {
  hasNative: !!(window.stsBridge),
  chooseExcel() {
    return new Promise((resolve) => {
      if (this.hasNative) {
        window.__fileChosen = (path) => resolve(path);
        window.stsBridge.chooseExcel();
      } else {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xls';
        input.onchange = () => {
          const f = input.files && input.files[0];
          resolve(f ? f : '');
        };
        input.click();
      }
    });
  },
  pickGallery() {
    return new Promise((resolve) => {
      if (this.hasNative) {
        let done = false;
        const timer = setTimeout(() => { done = true; window.__galleryResult = null; resolve(''); }, 60000);
        window.__galleryResult = (p) => { if (!done) { done = true; clearTimeout(timer); resolve(p || ''); } };
        window.stsBridge.pickFromGallery();
      } else {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.onchange = () => {
          const f = input.files && input.files[0];
          resolve(f ? f : '');
        };
        input.click();
      }
    });
  },
  pickGalleryMulti() {
    return new Promise((resolve) => {
      if (this.hasNative) {
        let done = false;
        const timer = setTimeout(() => { done = true; window.__galleryMultiResult = null; resolve([]); }, 120000);
        window.__galleryMultiResult = (json) => { if (!done) { done = true; clearTimeout(timer); resolve(json || '[]'); } };
        window.stsBridge.pickGalleryMulti();
      } else {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.multiple = true;
        input.onchange = () => {
          const fs = input.files ? Array.from(input.files) : [];
          resolve(fs.map(f => f));
        };
        input.click();
      }
    });
  },
  async readLocalFile(handle) {
    if (this.hasNative) {
      try {
        return await new Promise((resolve, reject) => {
          const xhr = new XMLHttpRequest();
          xhr.open('GET', encodeURI(handle), true);
          xhr.responseType = 'arraybuffer';
          xhr.onload = () => {
            if (xhr.status === 0 || xhr.status === 200) resolve(xhr.response);
            else reject(new Error('状态码 ' + xhr.status));
          };
          xhr.onerror = () => reject(new Error('XHR_FAILED'));
          xhr.send();
        });
      } catch (e) {
        const b64 = window.stsBridge.readFileAsBase64(handle);
        if (!b64) throw new Error('读取相册照片失败（' + (e && e.message || 'XHR') + '）');
        return base64ToArrayBuffer(b64);
      }
    }
    return await handle.arrayBuffer();
  },
  deleteTemp(path) {
    if (this.hasNative) { try { window.stsBridge.deleteTempFile(path); } catch (e) {} }
  },
  async readExcel(handle) {
    if (this.hasNative) {
      // 首选：XMLHttpRequest 读取本地文件
      try {
        return await new Promise((resolve, reject) => {
          const xhr = new XMLHttpRequest();
          xhr.open('GET', encodeURI(handle), true);
          xhr.responseType = 'arraybuffer';
          xhr.onload = () => {
            // file:// 请求成功时 status 为 0
            if (xhr.status === 0 || xhr.status === 200) {
              resolve(xhr.response);
            } else {
              reject(new Error('读取文件失败（状态码 ' + xhr.status + '）'));
            }
          };
          xhr.onerror = () => reject(new Error('XHR_FAILED'));
          xhr.send();
        });
      } catch (e) {
        // 兜底：部分机型 WebView 会拦截 file:// 网页请求，改由原生层直接读取
        const b64 = window.stsBridge.readFileAsBase64(handle);
        if (!b64) throw new Error('无法读取所选文件，请重新选择（错误代码 ' + (e && e.message || 'XHR') + '）');
        return base64ToArrayBuffer(b64);
      }
    }
    return await handle.arrayBuffer();
  },
  saveExcel(arrayBuffer, filename) {
    return new Promise((resolve) => {
      if (this.hasNative) {
        // 分块传输给原生层，避免大文件一次传递
        const bytes = new Uint8Array(arrayBuffer);
        const CHUNK = 512 * 1024;
        const total = Math.ceil(bytes.length / CHUNK);
        window.__fileSaved = (path) => resolve(path);
        window.stsBridge.saveExcelInit(filename);
        for (let i = 0; i < total; i++) {
          const sub = bytes.subarray(i * CHUNK, Math.min((i + 1) * CHUNK, bytes.length));
          let bin = '';
          const STEP = 0x8000;
          for (let j = 0; j < sub.length; j += STEP) {
            bin += String.fromCharCode.apply(null, sub.subarray(j, j + STEP));
          }
          window.stsBridge.saveExcelChunk(btoa(bin));
        }
        window.stsBridge.saveExcelFinish();
      } else {
        const blob = new Blob([arrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 3000);
        resolve('已保存到浏览器下载目录：' + filename);
      }
    });
  },
  shareFile(path) {
    if (this.hasNative) window.stsBridge.shareFile(path);
  },
  toast(msg) {
    if (this.hasNative) window.stsBridge.toast(msg);
  }
};

/* ---------- Excel 解析 ---------- */
function detectHeaderRow(ws, maxRow) {
  for (let r = 1; r <= Math.min(maxRow, 10); r++) {
    const rowCells = [];
    for (let c = 1; c <= 30; c++) {
      const cell = ws[colLetter(c - 1) + r];
      if (cell && cell.v != null && String(cell.v).trim() !== '') rowCells.push(String(cell.v).trim());
    }
    if (rowCells.some(t => t.includes('需求描述') || t.includes('编号') || t.includes('检查标准'))) return r;
  }
  return 0;
}
function colMapFromRow(ws, headerRow) {
  const map = {};
  const R = ws['!ref'];
  const maxCol = R ? XLSX.utils.decode_range(R).e.c + 1 : 30;
  for (let c = 1; c <= maxCol; c++) {
    const cell = ws[colLetter(c - 1) + headerRow];
    if (!cell || cell.v == null) continue;
    const t = String(cell.v).trim();
    if (!t) continue;
    if (!map.description && t.includes('需求描述')) map.description = c;
    if (!map.no && (t.includes('编号') || t.includes('需求编号'))) map.no = c;
    if (!map.spec && (t.includes('规范') || t.includes('检查标准') || t.includes('公差'))) map.spec = c;
    if (!map.tool && (t.includes('工具') || t.includes('型号') || t.includes('品牌'))) map.tool = c;
    if (!map.measured && t.includes('实测')) map.measured = c;
    if (t.includes('实拍配图')) map.image = c;                 // 优先：实拍配图列
    else if (!map.image && t.includes('附图')) map.image = c;  // 其次：附图列
  }
  return map;
}
function cellColNum(addr) { // 'B2' -> 2
  const m = String(addr).match(/^([A-Z]+)\d+$/);
  if (!m) return 0;
  let n = 0;
  for (let i = 0; i < m[1].length; i++) n = n * 26 + (m[1].charCodeAt(i) - 64);
  return n;
}
function getTaskInfo(ws) {
  const info = { stationName: '', supplier: '', inspector: '', startTime: '' };
  const labels = { '换电站名称': 'stationName', '落站供应商': 'supplier', '落站检查人': 'inspector', '落站起始时间': 'startTime' };
  for (const addr in ws) {
    if (!addr || addr.indexOf(':') >= 0) continue;
    const cell = ws[addr];
    if (!cell || cell.v == null) continue;
    const t = String(cell.v).trim();
    if (labels[t] && !info[labels[t]]) {
      const c = cellColNum(addr) + 1; // 值在标签右侧合并区
      const r = addr.replace(/^[A-Z]+/, '');
      const vcell = ws[colLetter(c - 1) + r];
      if (vcell && vcell.v != null) {
        const v = vcell.v;
        if (v instanceof Date) info[labels[t]] = fmtDateTime(v.getTime());
        else info[labels[t]] = String(v).trim();
      }
    }
  }
  return info;
}
function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const task = { id: 'T' + Date.now(), excelName: '', stationName: '', supplier: '', inspector: '', startTime: '', createdAt: Date.now() };
  const items = [];
  const sheetCols = {}; // sheetName -> colMap

  const infoSheet = wb.Sheets['调试检查清单'];
  if (infoSheet) Object.assign(task, getTaskInfo(infoSheet));
  if (task.startTime) {
    const st = new Date(task.startTime);
    if (!isNaN(st.getTime()) && typeof task.startTime === 'object') task.startTime = fmtDateTime(st.getTime());
  }
  // 兼容字符串时间
  if (typeof task.startTime === 'string' && task.startTime.includes('T')) {
    const st = new Date(task.startTime);
    if (!isNaN(st.getTime())) task.startTime = fmtDateTime(st.getTime());
  }

  wb.SheetNames.forEach(sn => {
    const base = sn.trim();
    if (!BUSY_SHEETS.includes(base)) return;
    const ws = wb.Sheets[sn];
    const headerRow = detectHeaderRow(ws, 10);
    if (!headerRow) return;
    const cmap = colMapFromRow(ws, headerRow);
    sheetCols[base] = cmap;
    if (!cmap.description) return;
    const range = XLSX.utils.decode_range(ws['!ref']);
    let idx = (sheetSeqStart(base)) - 1;
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const descCell = ws[colLetter(cmap.description - 1) + r];
      const desc = descCell && descCell.v != null ? String(descCell.v).trim() : '';
      if (!desc) continue;
      idx++;
      const get = (c) => {
        if (!c) return '';
        const cell = ws[colLetter(c - 1) + r];
        const v = cell && cell.v != null ? cell.v : '';
        if (v instanceof Date) return fmtDateTime(v.getTime());
        return String(v).trim();
      };
      items.push({
        id: `${base}_${r}`,
        sheetName: base,
        row: r,
        idx,
        no: get(cmap.no),
        desc,
        spec: get(cmap.spec),
        tool: get(cmap.tool),
        measured: get(cmap.measured),
        status: get(cmap.measured) ? 'done' : 'todo',
        images: []
      });
    }
  });
  return { task, items, sheetCols };
}

/* ---------- 渲染：首页 ---------- */
function renderHome() {
  const rt = $('recentTask');
  if (!Store.task || !Store.items.length) { rt.style.display = 'none'; return; }
  const p = Store.progress();
  rt.style.display = 'block';
  rt.innerHTML = `
    <div class="recent-title">最近任务</div>
    <div class="recent-row">换电站：<b>${esc(Store.task.stationName) || '-'}</b></div>
    <div class="recent-row">检查人：<b>${esc(Store.task.inspector) || '-'}</b>　供应商：<b>${esc(Store.task.supplier) || '-'}</b></div>
    <div class="recent-row">进度：<b>${p.done}/${p.total}</b> 项已完成（${p.pct}%）</div>
    <button id="btnContinue" class="btn btn-secondary">继续填写</button>`;
  $('btnContinue').onclick = () => { renderList(); showView('view-list'); };
}

/* ---------- 渲染：检查项列表 ---------- */
let listFilter = 'all';          // all | todo | done
let listSearch = '';
const expandedSheets = new Set(); // 已展开的表格分组（默认折叠）
function cardHtml(it) {
  const imgs = it.images;
  const thumbHtml = imgs.length ? '<div class="thumb-row">' + imgs.slice(0, 3).map(id =>
    `<img data-imgid="${id}" data-itemid="${it.id}" src="img://${id}" onerror="this.style.display='none'">`).join('') +
    (imgs.length > 3 ? `<div class="thumb-more">+${imgs.length - 3}</div>` : '') + '</div>' : '';
  return `
    <div class="item-card ${it.status === 'done' ? 'done' : ''}" data-item="${it.id}">
      <div class="item-top">
        <span class="item-no">${it.idx}</span>
        <span class="badge ${it.status === 'done' ? 'ok' : 'todo'}">${it.status === 'done' ? '已完成' : '未完成'}</span>
      </div>
      <div class="item-desc">${esc(it.desc)}</div>
      ${it.spec ? `<div class="item-meta">规范：${esc(it.spec)}</div>` : ''}
      ${it.measured ? `<div class="item-measured">${esc(it.measured)}</div>` : ''}
      ${thumbHtml}
    </div>`;
}
function groupBySheet(list) {
  const g = {};
  list.forEach(it => { (g[it.sheetName] = g[it.sheetName] || []).push(it); });
  return g;
}
function renderList() {
  if (!Store.task || !Store.items.length) { showView('view-home'); return; }
  $('listTaskName').textContent = Store.task.stationName || '调试检查表';
  const p = Store.progress();
  $('listProgressText').textContent = `${p.done} / ${p.total}`;
  $('listProgressBar').style.width = p.pct + '%';
  $('listProgressSub').textContent = `完成率 ${p.pct}%　·　检查人：${esc(Store.task.inspector) || '-'}`;

  // 筛选：点「未完成」只看未完成，点「已完成」只看已完成
  let list = Store.items;
  if (listFilter === 'todo') list = list.filter(it => it.status !== 'done');
  else if (listFilter === 'done') list = list.filter(it => it.status === 'done');

  // 搜索（≥2 个字生效，自动展开匹配分组）
  const kw = listSearch.trim();
  const searching = kw.length >= 2;
  if (searching) {
    list = list.filter(it =>
      (it.desc + ' ' + it.no + ' ' + it.spec + ' ' + it.tool + ' ' + (it.measured || '')).indexOf(kw) >= 0);
  }
  const isExpanded = (sn) => searching || expandedSheets.has(sn);

  // 单列分组渲染（按表格固定顺序）
  if (!list.length) {
    $('sheetGroups').innerHTML = '<div class="col-empty">' +
      (searching ? '没有找到匹配的检查项' : (listFilter === 'all' ? '暂无检查项' : '当前状态没有检查项')) + '</div>';
  } else {
    const g = groupBySheet(list);
    $('sheetGroups').innerHTML = BUSY_SHEETS.filter(sn => g[sn]).map(sn => `
      <div class="sheet-group">
        <div class="sheet-head" data-sheet="${sn}">
          <span class="sheet-arrow">${isExpanded(sn) ? '▾' : '▸'}</span>
          <span class="sheet-name">${esc(sn)}</span>
          <span class="sheet-count">${g[sn].length} 项</span>
        </div>
        <div class="sheet-body" style="${isExpanded(sn) ? '' : 'display:none'}">
          ${g[sn].map(cardHtml).join('')}
        </div>
      </div>`).join('');
  }
  $('listEmpty').style.display = Store.items.length ? 'none' : 'block';

  // 分组折叠切换
  document.querySelectorAll('.sheet-head').forEach(el => {
    el.addEventListener('click', (e) => {
      e.stopPropagation();
      const sn = el.dataset.sheet;
      if (expandedSheets.has(sn)) expandedSheets.delete(sn); else expandedSheets.add(sn);
      renderList();
    });
  });
  // 卡片点击
  document.querySelectorAll('.item-card').forEach(el => {
    el.addEventListener('click', (e) => {
      if (e.target.tagName === 'IMG' && e.target.dataset.imgid) {
        openLightbox(e.target.dataset.imgid, e.target.dataset.itemid);
        return;
      }
      currentItemId = el.dataset.item;
      renderDetail(currentItemId);
      showView('view-detail');
    });
  });
}

/* ---------- 渲染：详情 ---------- */
function renderDetail(itemId) {
  const it = Store.items.find(x => x.id === itemId);
  if (!it) return;
  $('dSheet').textContent = it.sheetName;
  $('dNo').textContent = it.idx + '　' + it.desc;
  $('dDesc').textContent = it.desc;
  $('dSpec').textContent = it.spec ? '规范/公差：' + it.spec : '';
  $('dTool').textContent = it.tool ? '工具：' + it.tool : '';
  const st = $('dStatus');
  st.textContent = it.status === 'done' ? '已完成' : '未完成';
  st.className = 'detail-status ' + (it.status === 'done' ? 'ok' : 'todo');
  $('dMeasured').value = it.measured || '';
  $('dImgCount').textContent = it.images.length;
  $('dImages').innerHTML = it.images.map(id =>
    `<img data-imgid="${id}" data-itemid="${it.id}" src="img://${id}" onerror="this.style.display='none'">`).join('') +
    `<div class="img-add" id="btnAddImages" role="button"><span class="ia-plus">＋</span><span class="ia-text">添加图片</span></div>`;
  document.querySelectorAll('#dImages img').forEach(el => {
    el.addEventListener('click', () => openLightbox(el.dataset.imgid, el.dataset.itemid));
  });
  const iaBtn = document.getElementById('btnAddImages');
  if (iaBtn) iaBtn.onclick = addGalleryPhotos;
  $('btnTakePhoto').style.display = it.images.length >= MAX_IMAGES ? 'none' : 'block';
}

/* 详情页批量添加相册图片（今日水印相机照片，不加水印，原图直接归档 + 防伪签名） */
async function addGalleryPhotos() {
  const it = Store.items.find(x => x.id === currentItemId);
  if (!it) return;
  if (it.images.length >= MAX_IMAGES) { Bridge.toast('照片已达上限（' + MAX_IMAGES + ' 张）'); return; }
  const picked = await Bridge.pickGalleryMulti();
  let paths = [];
  if (typeof picked === 'string') {
    try { paths = JSON.parse(picked || '[]'); } catch (e) { paths = []; }
  } else {
    paths = picked.map(f => f);
  }
  if (!paths.length) return;
  const task = Store.task;
  const lat = (camLocation && camLocation.lat) || 0, lng = (camLocation && camLocation.lng) || 0, addr = (camLocation && camLocation.addr) || '';
  let added = 0;
  for (let i = 0; i < paths.length; i++) {
    if (it.images.length + added >= MAX_IMAGES) { Bridge.toast('已达上限，部分未添加'); break; }
    const p = paths[i];
    if (!p) continue;
    let buf;
    try {
      if (typeof p === 'string') buf = await Bridge.readLocalFile(p);
      else buf = await p.arrayBuffer();
    } catch (e) { continue; }
    if (Bridge.deleteTemp && typeof p === 'string') Bridge.deleteTemp(p);
    const blob = new Blob([buf], { type: 'image/jpeg' });
    const imgId = `img_${Date.now()}_${randStr(6)}`;
    await Store.idbPut('images', imgId, blob);
    const code = await genAntiFake(task.id, it.id, Date.now(), task.inspector || '', randStr(6));
    const createdAt = Date.now();
    let hash = '', sig = '';
    try {
      hash = await hashBlob(blob);
      const payload = `${hash}|${task.id}|${it.id}|${createdAt}|${code}`;
      sig = await signPhotoPayload(payload);
    } catch (e) { console.warn('签名失败', e); }
    Store.imgMeta[imgId] = { code, createdAt, lat, lng, addr, hash, sig, pub: sigPubJwk, taskId: task.id, itemId: it.id, from: 'gallery-raw' };
    it.images.push(imgId);
    it.status = 'done';
    added++;
    await yieldUI();
  }
  Store.save();
  renderDetail(currentItemId);
  renderList();
  Bridge.toast(added ? '已添加 ' + added + ' 张图片' : '未添加任何图片');
}
/* 缩略图渲染（img:// 协议 -> IndexedDB Blob） */
const imgCache = {};
async function loadThumb(srcAttr) {
  if (!srcAttr || !srcAttr.startsWith('img://')) return;
  const id = srcAttr.slice(6);
  const imgs = document.querySelectorAll(`img[src="${srcAttr}"]`);
  if (!imgs.length) return;
  const draw = (url) => imgs.forEach(im => { im.src = url; });
  if (imgCache[id]) { draw(imgCache[id]); return; }
  try {
    const blob = await Store.idbGet('images', id);
    if (!blob) { imgs.forEach(im => im.style.display = 'none'); return; }
    const url = URL.createObjectURL(blob);
    imgCache[id] = url;
    draw(url);
  } catch (e) { imgs.forEach(im => im.style.display = 'none'); }
}
new MutationObserver(() => {
  document.querySelectorAll('img[src^="img://"]').forEach(im => {
    if (!im.dataset.loading) { im.dataset.loading = '1'; loadThumb(im.src); }
  });
}).observe(document.body, { childList: true, subtree: true });

/* ---------- 大图预览 ---------- */
let lbItemId = null, lbImgId = null;
function openLightbox(imgId, itemId) {
  lbImgId = imgId; lbItemId = itemId;
  const meta = Store.imgMeta[imgId];
  $('lbCode').textContent = meta ? '防伪码：' + meta.code : '';
  Store.idbGet('images', imgId).then(blob => {
    if (!blob) return;
    $('lbImg').src = URL.createObjectURL(blob);
    $('lightbox').style.display = 'flex';
  });
}
$('btnLbClose').onclick = () => { $('lightbox').style.display = 'none'; };
$('btnLbVerify').onclick = async () => {
  if (!lbImgId) return;
  const meta = Store.imgMeta[lbImgId];
  $('btnLbVerify').textContent = '正在验真…';
  try {
    const r = await verifyPhoto(lbImgId);
    if (r.ok) {
      const m = r.meta;
      alert('防伪验证通过\n\n' +
        '内容完整性：照片与拍摄时一致（SHA-256 校验通过）\n' +
        '密码学签名：设备密钥验证通过\n\n' +
        '拍摄时间：' + fmtDateTime(m.createdAt) + '\n' +
        '拍摄地点：' + (m.addr || ('经纬度 ' + m.lat + ', ' + m.lng)) + '\n' +
        '防伪码：' + (m.code || '-'));
    } else {
      alert('验真失败\n\n' + r.reason);
    }
  } catch (e) {
    alert('验真异常：' + (e.message || e));
  }
  $('btnLbVerify').textContent = '防伪验真';
};
$('btnLbDelete').onclick = async () => {
  if (!lbImgId || !lbItemId) return;
  if (!confirm('确定删除这张照片？')) return;
  const it = Store.items.find(x => x.id === lbItemId);
  if (it) {
    it.images = it.images.filter(id => id !== lbImgId);
    it.status = (it.measured && it.measured.trim()) || it.images.length ? 'done' : 'todo';
  }
  delete Store.imgMeta[lbImgId];
  try { await Store.idbDel('images', lbImgId); } catch (e) {}
  delete imgCache[lbImgId];
  Store.save();
  $('lightbox').style.display = 'none';
  renderDetail(lbItemId);
  renderList();
};

/* ---------- 水印相机 ---------- */
let camStream = null, camFacing = 'environment', camTorch = false, camItem = null, camLocation = { addr: '', lat: '', lng: '', state: 'pending' };
function setShutterState() {
  const disabled = camLocation.state !== 'ok';
  $('btnShutter').classList.toggle('disabled', disabled);
}
async function openCamera(itemId) {
  camItem = Store.items.find(x => x.id === itemId);
  if (!camItem) return;
  currentItemId = itemId;
  showView('view-camera');
  $('camStatus').style.display = 'none';
  const task = Store.task;
  $('wmStation').textContent = '换电站：' + (task.stationName || '-');
  $('wmItem').textContent = '检查项：' + camItem.idx + ' ' + camItem.desc;
  $('wmInspector').textContent = '检查人：' + (task.inspector || '-');
  camLocation = { addr: '', lat: '', lng: '', state: 'pending' };
  setShutterState();
  updateClock();
  setInterval(updateClock, 1000);
  // 手动填写地点始终可用（定位正常会自动填充后隐藏；定位慢/失败时可立即手动填写拍照）
  $('inpManualLoc').value = '';
  $('manualLocBox').style.display = 'block';
  getLocation();
  await startStream();
}
function confirmManualLoc() {
  const s = $('inpManualLoc').value.trim();
  if (s.length < 4) { Bridge.toast('请输入完整地点（至少4个字）'); return; }
  camLocation.addr = s;
  camLocation.state = 'ok';
  setShutterState();
  $('wmLocation').textContent = '地点：' + s;
  $('manualLocBox').style.display = 'none';
  Bridge.toast('已使用手动填写的地点');
}
function updateClock() {
  if (currentView === 'view-camera') {
    $('wmTime').textContent = '时间：' + fmtDateTime(Date.now());
  }
}
async function resolveAddress(lat, lng) {
  // 通道1：高德地图逆地理编码（原生转发，国内最准确）
  if (window.stsBridge) {
    try {
      const amap = await new Promise((res) => {
        let done = false;
        const timer = setTimeout(() => { done = true; window.__amapRegeoResult = null; res(''); }, 9000);
        window.__amapRegeoResult = (addr) => { if (!done) { done = true; clearTimeout(timer); res(addr || ''); } };
        window.stsBridge.amapRegeo(lat, lng, Store.amapKey || '');
      });
      if (amap) return amap;
    } catch (e) { /* 通道1异常，继续走通道2 */ }
  }
  // 通道2：OpenStreetMap 在线逆地理编码
  try {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 8000);
    const r = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lng}&zoom=16&accept-language=zh`, { signal: ctrl.signal });
    clearTimeout(t);
    const j = await r.json();
    if (j && j.display_name) return j.display_name;
  } catch (e) { /* 走通道3 */ }
  // 通道3：安卓系统逆地理编码（Geocoder）
  if (window.stsBridge) {
    return await new Promise((res) => {
      const timer = setTimeout(() => { window.__geoResult = null; res(''); }, 6000);
      window.__geoResult = (addr) => { clearTimeout(timer); res(addr || ''); };
      window.stsBridge.reverseGeocode(lat, lng);
    });
  }
  return '';
}
function getLocation() {
  if (!navigator.geolocation) {
    camLocation.state = 'fail';
    setShutterState();
    $('wmLocation').textContent = '地点：定位不可用（无法拍照）';
    return;
  }
  const tryLocate = (highAccuracy, attempt) => {
    navigator.geolocation.getCurrentPosition(async (pos) => {
      const lat = pos.coords.latitude.toFixed(6), lng = pos.coords.longitude.toFixed(6);
      camLocation.lat = lat; camLocation.lng = lng;
      $('wmGeo').textContent = '经纬度：' + lat + ', ' + lng;
      $('wmLocation').textContent = '地点：正在获取正式位置…';
      // 地址解析自动重试（最多3次，间隔1.2秒）
      let addr = '';
      for (let i = 0; i < 3 && !addr; i++) {
        addr = await resolveAddress(lat, lng);
        if (!addr && i < 2) await new Promise(r => setTimeout(r, 1200));
      }
      if (addr) {
        camLocation.addr = addr;
        camLocation.state = 'ok';
        setShutterState();
        $('wmLocation').textContent = '地点：' + addr;
        $('manualLocBox').style.display = 'none';
      } else {
        camLocation.state = 'fail';
        setShutterState();
        $('wmLocation').textContent = '地点：已获取坐标但地址解析失败（高德/在线/系统均无结果，可手动填写地点）';
        $('manualLocBox').style.display = 'block';
      }
    }, (err) => {
      // 第一次失败自动用普通精度重试一次（GPS 慢/室内场景）
      if (attempt === 0) { tryLocate(false, 1); return; }
      camLocation.state = 'fail';
      setShutterState();
      const msg = err.code === 1 ? '定位权限被拒绝（请到系统设置开启定位权限）'
        : err.code === 2 ? '定位不可用（请开启手机定位服务）'
        : '定位超时（信号弱，请到开阔处重试）';
      $('wmLocation').textContent = '地点：' + msg + '（无法拍照）';
      $('wmGeo').textContent = '经纬度：-';
      $('manualLocBox').style.display = 'block';
    }, { enableHighAccuracy: highAccuracy, timeout: highAccuracy ? 15000 : 10000, maximumAge: 30000 });
  };
  tryLocate(true, 0);
}
async function startStream() {
  stopStream();
  try {
    camStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: camFacing, width: { ideal: 1920 }, height: { ideal: 1080 } },
      audio: false
    });
    const v = $('camVideo');
    v.srcObject = camStream;
    await v.play().catch(() => {});
    $('camStatus').style.display = 'none';
  } catch (e) {
    $('camStatus').style.display = 'block';
    $('camStatus').textContent = '相机不可用：' + (e && e.message ? e.message : '请检查相机权限');
  }
}
function stopStream() {
  if (camStream) { camStream.getTracks().forEach(t => t.stop()); camStream = null; }
}
/* ---------- 水印模板（仿今日水印相机） ---------- */
function wmInfo(code) {
  const task = Store.task;
  return {
    station: task.stationName || '-',
    item: camItem.idx + ' ' + camItem.desc,
    time: fmtDateTime(Date.now()),
    addr: camLocation.addr || ('经纬度 ' + camLocation.lat + ',' + camLocation.lng),
    geo: camLocation.lat ? camLocation.lat + ', ' + camLocation.lng : '-',
    inspector: task.inspector || '-',
    code
  };
}
/* 无背景水印：白字黑描边直接叠在画面（全透明，不遮挡） */
function drawWatermark(ctx, w, h, code) {
  const task = Store.task;
  const fs = Math.max(14, Math.round(w / 30));
  const lh = fs * 1.5;
  const pad = Math.max(10, Math.round(w * 0.025));
  const lines = [
    '换电站：' + (task.stationName || '-'),
    '检查项：' + camItem.idx + ' ' + camItem.desc,
    '时间：' + fmtDateTime(Date.now()),
    '地点：' + (camLocation.addr || ('经纬度 ' + camLocation.lat + ',' + camLocation.lng)),
    '经纬度：' + (camLocation.lat ? camLocation.lat + ', ' + camLocation.lng : '-'),
    '检查人：' + (task.inspector || '-'),
    '防伪码：' + code
  ];
  const totalH = lines.length * lh;
  let y = h - pad - totalH;
  ctx.font = `600 ${fs}px "Microsoft YaHei", sans-serif`;
  ctx.textBaseline = 'top';
  ctx.lineJoin = 'round';
  ctx.strokeStyle = 'rgba(0,0,0,0.85)';
  ctx.lineWidth = Math.max(2, Math.round(fs / 12));
  ctx.fillStyle = '#ffffff';
  const maxW = w - pad * 2;
  lines.forEach((ln) => {
    let t = ln;
    while (ctx.measureText(t).width > maxW && t.length > 1) t = t.slice(0, -1);
    if (t !== ln) t += '…';
    ctx.strokeText(t, pad, y);
    ctx.fillText(t, pad, y);
    y += lh;
  });
}
async function takePhoto() {
  const v = $('camVideo');
  if (!v.videoWidth) { Bridge.toast('相机尚未就绪'); return; }
  if (camLocation.state !== 'ok') {
    Bridge.toast(camLocation.state === 'pending' ? '正在获取定位，请稍候再拍，或手动填写地点' : '未获取到地点，请手动填写地点后拍照');
    return;
  }
  if (!camLocation.addr) {
    // 有坐标但地址为空：拍照前强制再解析一次
    Bridge.toast('正在重新获取正式地址…');
    camLocation.state = 'pending';
    setShutterState();
    const addr = await resolveAddress(camLocation.lat, camLocation.lng);
    if (addr) {
      camLocation.addr = addr;
      camLocation.state = 'ok';
      setShutterState();
      $('wmLocation').textContent = '地点：' + addr;
      $('manualLocBox').style.display = 'none';
    } else {
      camLocation.state = 'fail';
      setShutterState();
      $('manualLocBox').style.display = 'block';
      Bridge.toast('仍未获取到正式地址，请手动填写地点后拍照');
      return;
    }
  }
  const task = Store.task;
  const code = await genAntiFake(task.id, camItem.id, Date.now(), task.inspector || '', randStr(6));
  const edge = Store.photoEdge || IMG_MAX_EDGE;
  const scale = Math.min(1, edge / Math.max(v.videoWidth, v.videoHeight));
  const W = Math.round(v.videoWidth * scale), H = Math.round(v.videoHeight * scale);
  const c = $('camCanvas');
  c.width = W; c.height = H;
  const ctx = c.getContext('2d');
  ctx.drawImage(v, 0, 0, W, H);
  drawWatermark(ctx, W, H, code);
  const doSave = async (blob) => {
    const imgId = `img_${Date.now()}_${randStr(6)}`;
    await Store.idbPut('images', imgId, blob);
    // 防伪签名：内容哈希 + ECDSA 签名 + 元数据锁定
    const createdAt = Date.now();
    let hash = '', sig = '';
    try {
      hash = await hashBlob(blob);
      const payload = `${hash}|${task.id}|${camItem.id}|${createdAt}|${code}`;
      sig = await signPhotoPayload(payload);
    } catch (e) { console.warn('签名失败', e); }
    Store.imgMeta[imgId] = {
      code, createdAt, lat: camLocation.lat, lng: camLocation.lng, addr: camLocation.addr,
      hash, sig, pub: sigPubJwk, taskId: task.id, itemId: camItem.id, w: W, h: H
    };
    camItem.images.push(imgId);
    camItem.status = 'done';
    Store.save();
    const name = `${fmtDateFile(Date.now())}_${code.slice(0, 6)}.jpg`;
    Bridge.toast('已拍照并保存');
    return { imgId, code };
  };
  const makeBlob = (quality) => new Promise((res, rej) => {
    c.toBlob(b => b ? res(b) : rej(new Error('toBlob failed')), 'image/jpeg', quality);
  });
  let blob = await makeBlob(0.78);
  if (blob.size > IMG_MAX_BYTES) blob = await makeBlob(0.6);
  if (blob.size > IMG_MAX_BYTES) blob = await makeBlob(0.45);
  await doSave(blob);
  exitCamera();
}
/* 从相册选照片：解码 → 缩放 → 当前水印 → 防伪签名 → 归档（与拍照一致） */
async function galleryPhoto() {
  if (camLocation.state !== 'ok') {
    Bridge.toast(camLocation.state === 'pending' ? '正在获取定位，请稍候再选，或手动填写地点' : '未获取到地点，请手动填写地点后再选');
    return;
  }
  const path = await Bridge.pickGallery();
  if (!path) { return; }
  Bridge.toast('正在处理相册照片…');
  let buf;
  try { buf = await Bridge.readLocalFile(path); } catch (e) { Bridge.toast('读取相册照片失败'); return; }
  if (Bridge.deleteTemp) Bridge.deleteTemp(path);
  const task = Store.task;
  const code = await genAntiFake(task.id, camItem.id, Date.now(), task.inspector || '', randStr(6));
  const edge = Store.photoEdge || IMG_MAX_EDGE;
  // 解码图片
  const blob0 = new Blob([buf], { type: 'image/jpeg' });
  const url = URL.createObjectURL(blob0);
  const img = new Image();
  await new Promise((res, rej) => { img.onload = res; img.onerror = () => rej(new Error('图片解码失败')); img.src = url; });
  URL.revokeObjectURL(url);
  const scale = Math.min(1, edge / Math.max(img.naturalWidth || 1, img.naturalHeight || 1));
  const W = Math.round((img.naturalWidth || 1) * scale), H = Math.round((img.naturalHeight || 1) * scale);
  const c = $('camCanvas');
  c.width = W; c.height = H;
  const ctx = c.getContext('2d');
  ctx.drawImage(img, 0, 0, W, H);
  drawWatermark(ctx, W, H, code);
  const doSave = async (blob) => {
    const imgId = `img_${Date.now()}_${randStr(6)}`;
    await Store.idbPut('images', imgId, blob);
    const createdAt = Date.now();
    let hash = '', sig = '';
    try {
      hash = await hashBlob(blob);
      const payload = `${hash}|${task.id}|${camItem.id}|${createdAt}|${code}`;
      sig = await signPhotoPayload(payload);
    } catch (e) { console.warn('签名失败', e); }
    Store.imgMeta[imgId] = {
      code, createdAt, lat: camLocation.lat, lng: camLocation.lng, addr: camLocation.addr,
      hash, sig, pub: sigPubJwk, taskId: task.id, itemId: camItem.id, w: W, h: H, from: 'gallery'
    };
    camItem.images.push(imgId);
    camItem.status = 'done';
    Store.save();
    Bridge.toast('相册照片已添加');
    return { imgId, code };
  };
  const makeBlob = (quality) => new Promise((res, rej) => {
    c.toBlob(b => b ? res(b) : rej(new Error('toBlob failed')), 'image/jpeg', quality);
  });
  let blob = await makeBlob(0.78);
  if (blob.size > IMG_MAX_BYTES) blob = await makeBlob(0.6);
  if (blob.size > IMG_MAX_BYTES) blob = await makeBlob(0.45);
  await doSave(blob);
  exitCamera();
}
function exitCamera() {
  stopStream();
  renderDetail(currentItemId);
  renderList();
  showView('view-detail');
}
$('btnShutter').onclick = () => takePhoto();
$('btnGallery').onclick = () => galleryPhoto();
$('btnSwitchCam').onclick = async () => {
  camFacing = camFacing === 'environment' ? 'user' : 'environment';
  await startStream();
};
$('btnTorch').onclick = async () => {
  camTorch = !camTorch;
  try {
    const track = camStream && camStream.getVideoTracks()[0];
    if (track && track.applyConstraints) {
      await track.applyConstraints({ advanced: [{ torch: camTorch }] });
    }
  } catch (e) { Bridge.toast('当前设备不支持闪光灯'); }
};

/* ---------- 导出 ---------- */
function colIndexFromLetter(l) {
  let n = 0;
  for (let i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64);
  return n - 1; // 0-based
}
async function exportExcel() {
  if (!Store.task || !Store.items.length) { Bridge.toast('暂无任务'); return; }
  showView('view-export');
  $('exportDone').style.display = 'none';
  $('exportInfo').textContent = `换电站：${Store.task.stationName || '-'}｜检查项 ${Store.items.length} 项｜待导出图片 ${Store.items.reduce((s, x) => s + x.images.length, 0)} 张`;
  setExportProgress(2, '正在读取原表格…');
  try {
    const srcBlob = await Store.idbGet('source', 'source');
    if (!srcBlob) throw new Error('原表格数据缺失，请重新选择 Excel');
    const buf = await srcBlob.arrayBuffer();
    setExportProgress(3, '正在加载导出组件…');
    await ensureLib('exceljs');
    const wb = new ExcelJS.Workbook();
    setExportProgress(15, '正在解析原表格结构…');
    await wb.xlsx.load(buf);

    // 找出业务 sheet 的列映射（从 items 按 sheet 分组，动态探测）
    const sheets = {};
    Store.items.forEach(it => { (sheets[it.sheetName] = sheets[it.sheetName] || []).push(it); });

    let step = 0;
    const totalSteps = Object.keys(sheets).length;
    for (const sn of Object.keys(sheets)) {
      step++;
      const ws = wb.getWorksheet(sn.trim());
      if (!ws) continue;
      const colMap = detectColMap(ws);
      const list = sheets[sn];
      for (let k = 0; k < list.length; k++) {
        const it = list[k];
        const row = ws.getRow(it.row);
        // 写入实测结果
        if (colMap.measured) {
          const cell = row.getCell(colMap.measured);
          cell.value = it.measured || '';
        }
        // 插入图片（动态分格；插入列可由设置页按表覆盖，默认探测列）
        const imgCol = sheetPhotoCol(sn.trim()) || colMap.image;
        if (it.images.length && imgCol) {
          await insertImages(ws, it.row, imgCol, it.images);
        }
        setExportProgress(15 + Math.round(step / totalSteps * 60) + Math.round(k / list.length * 5), `正在写入：${sn} ${it.idx}/${list.length}`);
        await yieldUI();
      }
    }
    setExportProgress(82, '正在生成新文件…');
    const outBuf = await wb.xlsx.writeBuffer();
    setExportProgress(95, '正在保存文件…');
    const fname = `${(Store.task.stationName || '调试检查表')}_调试检查表_导出_${fmtDateFile(Date.now())}.xlsx`;
    const path = await Bridge.saveExcel(outBuf, fname);
    setExportProgress(100, '完成');
    $('exportDone').style.display = 'block';
    $('exportPath').textContent = path || '已保存';
    $('btnShareFile').style.display = Bridge.hasNative && path ? 'block' : 'none';
    window.__lastExportPath = path;
  } catch (e) {
    console.error(e);
    $('exportStatus').textContent = '导出失败：' + (e.message || e);
  }
}
function detectColMap(ws) {
  // 在导出副本上探测表头行与列（与解析时一致的规则）
  const cmap = {};
  for (let r = 1; r <= Math.min(ws.rowCount, 10); r++) {
    const row = ws.getRow(r);
    const vals = [];
    row.eachCell({ includeEmpty: false }, (cell) => { vals.push(String(cell.value == null ? '' : cell.value)); });
    if (!vals.some(t => t.includes('需求描述') || t.includes('编号') || t.includes('检查标准'))) continue;
    row.eachCell({ includeEmpty: false }, (cell, n) => {
      const t = String(cell.value == null ? '' : cell.value).trim();
      if (!t) return;
      if (!cmap.description && t.includes('需求描述')) cmap.description = n;
      if (!cmap.measured && t.includes('实测')) cmap.measured = n;
      if (t.includes('实拍配图')) cmap.image = n;
      else if (!cmap.image && t.includes('附图')) cmap.image = n;
    });
    break;
  }
  return cmap;
}
/* 根据照片数量与列宽自动分排（1:1行 2:各50% 3:各1/3 4:2×2 5:3+2 6:3×2 7:4+3 8:4×2 9:3×3 10-12:每行4 13-20:每行5） */
function calcGridByCols(n, colWpx) {
  let cols, rows;
  if (n <= 3) { cols = n; rows = 1; }
  else if (n === 4) { cols = 2; rows = 2; }
  else if (n <= 8) { cols = Math.ceil(n / 2); rows = 2; }
  else if (n === 9) { cols = 3; rows = 3; }
  else if (n <= 12) { cols = 4; rows = Math.ceil(n / 4); }
  else if (n <= 20) { cols = 5; rows = Math.ceil(n / 5); }
  else { cols = Math.ceil(Math.sqrt(n)); rows = Math.ceil(n / cols); }
  return { cols, rows };
}
/* 导出照片：按格子大小动态分排，照片暴力拉伸铺满各自格子（native EMU 像素级定位，不重叠） */
async function insertImages(ws, rowIndex, imageColIdx, imageIds) {
  const n = imageIds.length;
  if (!n) return;
  const EMU = 9525;
  // 列宽（px，Excel 列宽单位≈7px）
  let colWpx = 100;
  try {
    const col = ws.getColumn(imageColIdx);
    if (col && col.width) colWpx = Math.round(col.width * 7);
  } catch (e) {}
  const PAD = 8, gap = 6;
  // 列宽不足自动加宽（照片常规宽度）
  const neededColW = Math.max(colWpx / 7, Math.round((170 + PAD) / 7));
  try {
    const col = ws.getColumn(imageColIdx);
    if (!col.width || col.width < neededColW) col.width = neededColW;
    colWpx = Math.round((col.width || neededColW) * 7);
  } catch (e) {}
  // 按数量与列宽自动分排
  const { cols, rows } = calcGridByCols(n, colWpx);
  // 每格宽/高：格高=格宽×0.8（模板观感），照片暴力拉伸填满
  const cellW = (colWpx - PAD * 2 - (cols - 1) * gap) / cols;
  const cellH = Math.max(40, Math.round(cellW * 0.8));
  const imgW = Math.floor(cellW);
  const imgH = Math.floor(cellH);
  // 行高（points）：总高=行数×格高+行间隙+边距，只增不减
  const totalPx = rows * cellH + (rows - 1) * gap + PAD;
  const row = ws.getRow(rowIndex);
  const curPx = (row.height || 15) * 1.333;
  if (totalPx > curPx) row.height = Math.round(totalPx / 1.333);
  // 像素级定位插入
  for (let i = 0; i < n; i++) {
    const blob = await Store.idbGet('images', imageIds[i]);
    if (!blob) continue;
    const b64 = await blobToBase64(blob);
    const imageId = ws.workbook.addImage({ base64: b64, extension: 'jpeg' });
    const c = i % cols, r = Math.floor(i / cols);
    const xPx = Math.round(PAD + c * (cellW + gap));
    const yPx = Math.round(PAD / 2 + r * (cellH + gap));
    ws.addImage(imageId, {
      tl: {
        nativeCol: imageColIdx - 1,
        nativeColOff: Math.round(xPx * EMU),
        nativeRow: rowIndex - 1,
        nativeRowOff: Math.round(yPx * EMU)
      },
      ext: { width: imgW, height: imgH }
    });
    await yieldUI();
  }
}
function blobToBase64(blob) {
  return new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = () => { const s = fr.result; res(s.indexOf(',') >= 0 ? s.slice(s.indexOf(',') + 1) : s); };
    fr.onerror = () => rej(fr.error);
    fr.readAsDataURL(blob);
  });
}
function setExportProgress(pct, text) {
  $('exportProgressBar').style.width = pct + '%';
  $('exportStatus').textContent = text;
}
function yieldUI() { return new Promise(r => setTimeout(r, 0)); }

/* ---------- 事件绑定：首页 ---------- */
$('btnChooseExcel').onclick = async () => {
  const mask = $('loadMask');
  const loadText = $('loadText');
  mask.style.display = 'flex';
  loadText.textContent = '正在选择文件…';
  try {
    const handle = await Bridge.chooseExcel();
    if (!handle) { mask.style.display = 'none'; return; }
    loadText.textContent = '正在读取文件…';
    const buf = await Bridge.readExcel(handle);
    if (!buf || !buf.byteLength) throw new Error('未能读取到文件内容，请重新选择');
    loadText.textContent = '正在加载表格解析组件…';
    await ensureLib('xlsx');
    loadText.textContent = '正在解析检查项（大文件可能需要十几秒）…';
    const r = parseWorkbook(buf);
    if (!r.items.length) throw new Error('未在表格中找到可用的检查项，请确认选择了正确的调试检查表');
    Store.task = r.task;
    Store.items = r.items;
    Store.sheetCols = r.sheetCols || {};
    Store.task.excelName = (typeof handle === 'string' ? handle.split('/').pop() : (handle.name || ''));
    Store.save();
    try { await Store.idbPut('source', 'source', new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); }
    catch (e2) { console.warn('原表缓存失败（不影响本次使用，导出需重新选择）', e2); }
    renderHome();
    renderList();
    showView('view-list');
  } catch (e) {
    console.error(e);
    const msg = (e && e.message) ? e.message : String(e);
    recordError(msg, e && e.stack);
    alert('操作失败：' + msg + '\n\n（已记录到错误日志，可点首页底部"错误日志"查看）');
  } finally {
    mask.style.display = 'none';
  }
};
$('btnExport').onclick = () => exportExcel();
$('btnBack').onclick = () => {
  if (currentView === 'view-detail') { renderList(); showView('view-list'); }
  else if (currentView === 'view-list') { renderHome(); showView('view-home'); }
  else if (currentView === 'view-export') { renderList(); showView('view-list'); }
};
$('btnSaveMeasured').onclick = () => {
  const it = Store.items.find(x => x.id === currentItemId);
  if (!it) return;
  it.measured = $('dMeasured').value.trim();
  it.status = (it.measured || it.images.length) ? 'done' : 'todo';
  Store.save();
  Bridge.toast('实测结果已保存');
  renderDetail(currentItemId);
  renderList();
};
$('btnTakePhoto').onclick = () => openCamera(currentItemId);
$('btnShareFile').onclick = () => {
  if (window.__lastExportPath) Bridge.shareFile(window.__lastExportPath);
};
$('btnBackToList').onclick = () => { renderList(); showView('view-list'); };
$('btnSettings').onclick = openSettings;
$('btnSettingsBack').onclick = () => { showView('view-home'); };
$('btnSaveKey').onclick = saveKey;
$('btnTestKey').onclick = testKey;
try { $('btnManualLoc').onclick = confirmManualLoc; } catch (e) { recordError('btnManualLoc 绑定失败', e && e.stack ? e.stack : String(e)); }
try {
  document.querySelectorAll('#photoEdgeChips .chip').forEach(c => {
    c.onclick = () => { Store.photoEdge = Number(c.dataset.edge); Store.save(); syncSettingsUI(); Bridge.toast('照片尺寸：最长边 ' + c.textContent.trim()); };
  });
} catch (e) { recordError('设置chips绑定失败', e && e.stack ? e.stack : String(e)); }

/* ---------- 筛选与搜索 ---------- */
function syncFilterChips() {
  $('btnFilterAll').classList.toggle('active', listFilter === 'all');
  $('btnFilterTodo').classList.toggle('active', listFilter === 'todo');
  $('btnFilterDone').classList.toggle('active', listFilter === 'done');
}
$('btnFilterAll').onclick = () => { listFilter = 'all'; syncFilterChips(); renderList(); };
$('btnFilterTodo').onclick = () => { listFilter = 'todo'; syncFilterChips(); renderList(); };
$('btnFilterDone').onclick = () => { listFilter = 'done'; syncFilterChips(); renderList(); };
$('listSearch').addEventListener('input', () => {
  listSearch = $('listSearch').value;
  renderList();
});

/* ---------- 错误日志 ---------- */
function refreshErrorLogBtn() {
  const has = !!readLastError();
  $('btnErrorLog').style.display = has ? 'block' : 'none';
}
$('btnErrorLog').onclick = () => {
  const err = readLastError();
  if (!err) { Bridge.toast('暂无错误记录'); return; }
  alert('错误时间：' + fmtDateTime(err.time) + '\n错误信息：' + (err.msg || '(空)') + '\n\n详情：' + (err.stack || '无'));
};

/* ---------- 任务信息编辑 ---------- */
$('btnEditTask').onclick = () => {
  const t = Store.task || {};
  $('tmStation').value = t.stationName || '';
  $('tmSupplier').value = t.supplier || '';
  $('tmInspector').value = t.inspector || '';
  const st = t.startTime;
  if (st) {
    const d = new Date(String(st).replace(' ', 'T'));
    if (!isNaN(d.getTime())) {
      const p = (n) => pad2(n);
      $('tmStart').value = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
    } else $('tmStart').value = '';
  } else $('tmStart').value = '';
  $('taskModal').style.display = 'flex';
};
$('btnTmCancel').onclick = () => { $('taskModal').style.display = 'none'; };
$('btnTmSave').onclick = () => {
  const t = Store.task;
  if (!t) return;
  t.stationName = $('tmStation').value.trim();
  t.supplier = $('tmSupplier').value.trim();
  t.inspector = $('tmInspector').value.trim();
  const sv = $('tmStart').value;
  t.startTime = sv ? sv.replace('T', ' ') : '';
  Store.save();
  $('taskModal').style.display = 'none';
  renderList();
  Bridge.toast('任务信息已保存');
};

/* 后退键 */
document.addEventListener('backbutton', (e) => {
  if (currentView === 'view-camera') { e.preventDefault(); exitCamera(); }
  else if (currentView !== 'view-home') { e.preventDefault(); $('btnBack').click(); }
});

/* ---------- 登录权限 ---------- */
const AUTH_USER = '19121944714';
const AUTH_PASS = '123456';
const AUTH_KEY = 'sts.auth';
function isAuthed() { try { return localStorage.getItem(AUTH_KEY) === '1'; } catch (e) { return false; } }
function doLogin() {
  const u = $('inpUser').value.trim();
  const p = $('inpPass').value;
  if (u === AUTH_USER && p === AUTH_PASS) {
    try { localStorage.setItem(AUTH_KEY, '1'); } catch (e) {}
    enterApp();
  } else {
    const m = $('loginMsg');
    m.style.display = 'block';
    m.textContent = '账号或密码错误';
    $('inpPass').value = '';
  }
}
function enterApp() {
  (async function () {
    try { await Store.init(); } catch (e) { console.error(e); }
    renderHome();
    refreshErrorLogBtn();
    showView('view-home');
  })();
}
function doLogout() {
  try { localStorage.removeItem(AUTH_KEY); } catch (e) {}
  $('inpUser').value = '';
  $('inpPass').value = '';
  $('loginMsg').style.display = 'none';
  showView('view-login');
}
$('btnLogin').onclick = doLogin;
$('btnLogout').onclick = doLogout;

/* ---------- PWA 离线缓存 ---------- */
try {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch((e) => console.warn('SW 注册失败（file:// 环境忽略）', e));
  }
} catch (e) { /* ignore */ }

/* ---------- 启动 ---------- */
(async function init() {
  if (!isAuthed()) { showView('view-login'); return; }
  enterApp();
})();

