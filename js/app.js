/* ════════════════════════════════════════════════════════════
   DOC INTERIORS – app.js
   GitHub Pages frontend  ↔  Google Apps Script backend

   HOW CONFIG WORKS
   ─────────────────
   1. User pastes only the GAS Web App URL into the banner.
   2. On connect, we immediately call  ?action=getConfig
      which returns { spreadsheetId, spreadsheetUrl, … }
      directly from Code.gs (where SPREADSHEET_ID is set).
   3. Both the GAS URL and the spreadsheet ID are stored in
      localStorage so the app works offline across refreshes.
   4. A status bar shows which spreadsheet is connected.
   ════════════════════════════════════════════════════════════ */

'use strict';

const APP = (() => {

  /* ── Storage keys ──────────────────────────────────────── */
  const KEY_GAS_URL = 'doc_interiors_gas_url';
  const KEY_SS_ID   = 'doc_interiors_spreadsheet_id';
  const KEY_SS_URL  = 'doc_interiors_spreadsheet_url';

  /* ── Getters ───────────────────────────────────────────── */
  function getGasUrl() { return localStorage.getItem(KEY_GAS_URL) || ''; }
  function getSsId()   { return localStorage.getItem(KEY_SS_ID)   || ''; }
  function getSsUrl()  { return localStorage.getItem(KEY_SS_URL)  || ''; }

  /* ── saveGasUrl: called when user clicks Connect ───────── */
  async function saveGasUrl() {
    const v = document.getElementById('gasUrlInput').value.trim();
    if (!v.startsWith('https://script.google.com')) {
      toast('Please enter a valid GAS Web App URL', 'err'); return;
    }
    localStorage.setItem(KEY_GAS_URL, v);

    // Immediately fetch the config (which includes Spreadsheet ID)
    // from Code.gs so the user never has to enter the sheet ID.
    loading(true);
    try {
      const cfg = await gasGet('getConfig');           // { spreadsheetId, spreadsheetUrl, … }
      localStorage.setItem(KEY_SS_ID,  cfg.spreadsheetId);
      localStorage.setItem(KEY_SS_URL, cfg.spreadsheetUrl);
      document.getElementById('configBanner').classList.add('hidden');
      updateStatusBar(cfg.spreadsheetId, cfg.spreadsheetUrl);
      toast('Connected! Spreadsheet ID received from GAS ✓', 'ok');
      await init();
    } catch (err) {
      toast('Connection failed: ' + err.message, 'err');
    } finally {
      loading(false);
    }
  }

  /* ── openConfig: toggle the banner ─────────────────────── */
  function openConfig() {
    const banner = document.getElementById('configBanner');
    const isHidden = banner.classList.contains('hidden');
    banner.classList.toggle('hidden', !isHidden);
    if (isHidden) {
      document.getElementById('gasUrlInput').value = getGasUrl();
    }
  }

  /* ── Status bar: shows which sheet is connected ─────────── */
  function updateStatusBar(ssId, ssUrl) {
    const bar = document.getElementById('statusBar');
    if (!bar) return;
    if (ssId && ssId !== 'PASTE_YOUR_SPREADSHEET_ID_HERE') {
      bar.innerHTML = `
        📊 Connected to sheet: 
        <a href="${ssUrl}" target="_blank" style="color:var(--gold);text-decoration:none;">
          ${ssId.slice(0, 20)}…
        </a>
        <span style="margin-left:12px;opacity:.6;font-size:11px;">ID passed from GAS Code.gs</span>`;
      bar.classList.remove('hidden');
    } else {
      bar.classList.add('hidden');
    }
  }

  /* ── HTTP helpers ──────────────────────────────────────── */
  async function gasGet(action, params = {}) {
    const base = getGasUrl();
    if (!base) { toast('GAS URL not set — click 🔗 in the navbar', 'err'); return null; }
    const url = new URL(base);
    url.searchParams.set('action', action);
    Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
    const res  = await fetch(url.toString());
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || 'GAS error');
    return json.data;
  }

  async function gasPost(action, payload = {}) {
    const base = getGasUrl();
    if (!base) { toast('GAS URL not set — click 🔗 in the navbar', 'err'); return null; }
    const res  = await fetch(base, {
      method : 'POST',
      headers: { 'Content-Type': 'text/plain' },  // GAS CORS requirement
      body   : JSON.stringify({ action, payload })
    });
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || 'GAS error');
    return json;
  }

  /* ── App state ─────────────────────────────────────────── */
  let state = {
    estimates  : [],
    dropdowns  : { catA:[], catB:[], additionalWorks:[], gypsum:[], kitchenAcc:[] },
    adminData  : {},
    currentId  : null,
    delTargetId: null,
    adminTab   : 'catA',
    adminDel   : null,
    modalEdit  : null,       // { key, idx } when editing admin row
    active     : { catA:false, kitAcc:false, addWorks:false, catB:false, gypsum:false },
    rows       : { catA:[], kitAcc:[], addWorks:[], catB:[], gypsum:[] }
  };

  /* ── Toast / Loading ───────────────────────────────────── */
  let toastTimer;
  function toast(msg, type = '') {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.className   = 'toast ' + type;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.add('hidden'), 3500);
  }
  function loading(on) {
    document.getElementById('loading').classList.toggle('hidden', !on);
  }

  /* ── Screen router ─────────────────────────────────────── */
  function showScreen(name) {
    // Hide all screens first
    document.querySelectorAll('.screen').forEach(s => {
      s.classList.remove('active');
      s.classList.add('hidden');
    });
    // Deactivate all nav tabs
    document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
    // Show the requested screen
    var target = document.getElementById('screen-' + name);
    if (target) { target.classList.add('active'); target.classList.remove('hidden'); }
    // Activate matching nav tab
    var tab = document.querySelector('[data-screen="' + name + '"]');
    if (tab) tab.classList.add('active');
    // Screen-specific side effects
    if (name === 'preview') renderPreview();
    if (name === 'admin')   loadAdminData();
    if (name === 'entries') loadEntries();
  }

  /* ══════════════════════════════════════════════════════════
     INIT
  ══════════════════════════════════════════════════════════ */
  async function init() {
    loading(true);
    try {
      // Load dropdowns AND entries at the same time for speed
      const results = await Promise.all([
        gasGet('getDropdownData'),
        gasGet('getAllEstimates')
      ]);
      // Store dropdowns — critical for the edit screen to work
      state.dropdowns = results[0] || { catA:[], catB:[], additionalWorks:[], gypsum:[], kitchenAcc:[] };
      // Store entries
      state.estimates = results[1] || [];
      // Render the entries list
      renderEntries(state.estimates);
    } catch (e) {
      toast('Init error: ' + e.message, 'err');
    } finally {
      loading(false);
    }
  }

  /* ══════════════════════════════════════════════════════════
     SCREEN 1 – ENTRIES
  ══════════════════════════════════════════════════════════ */
  async function loadEntries() {
    loading(true);
    try {
      state.estimates = await gasGet('getAllEstimates') || [];
      renderEntries(state.estimates);
    } catch (e) {
      toast('Load error: ' + e.message, 'err');
    } finally {
      loading(false);
    }
  }

  function renderEntries(list) {
    const tbody = document.getElementById('entriesBody');
    if (!list || !list.length) {
      tbody.innerHTML = `<tr><td colspan="9" class="empty">No estimates yet. <a href="#" onclick="APP.startNew()">Create first →</a></td></tr>`;
      return;
    }
    tbody.innerHTML = list.map((e, i) => `
      <tr>
        <td style="color:var(--muted);font-size:11px;">${i + 1}</td>
        <td><strong>${e.estimationNo || e.id}</strong></td>
        <td>${e.clientName || '—'}</td>
        <td style="color:var(--muted);font-size:12px;">${e.siteAddress || '—'}</td>
        <td>${e.mobile || '—'}</td>
        <td style="font-size:12px;">${e.estimateDate || fmtDate(e.dateCreated)}</td>
        <td>${e.preparedBy || '—'}</td>
        <td><span class="badge-active">Active</span></td>
        <td style="white-space:nowrap">
          <button class="btn btn-outline btn-sm" onclick="APP.editEstimate('${e.id}')">✏️ Edit</button>
          <button class="btn btn-outline btn-sm" style="margin-left:4px" onclick="APP.previewById('${e.id}')">👁 View</button>
          <button class="btn btn-danger  btn-sm" style="margin-left:4px" onclick="APP.showDelModal('${e.id}')">🗑</button>
        </td>
      </tr>`).join('');
  }

  function filterEntries() {
    const q = document.getElementById('searchInput').value.toLowerCase();
    renderEntries(state.estimates.filter(e =>
      (e.clientName  || '').toLowerCase().includes(q) ||
      (e.estimationNo|| '').toLowerCase().includes(q) ||
      (e.id          || '').toLowerCase().includes(q) ||
      (e.siteAddress || '').toLowerCase().includes(q)
    ));
  }

  function fmtDate(d) {
    if (!d) return '—';
    try { return new Date(d).toLocaleDateString('en-IN'); } catch { return d; }
  }

  /* Delete */
  function showDelModal(id) {
    state.delTargetId = id;
    document.getElementById('delModal').classList.remove('hidden');
  }
  async function confirmDelete() {
    if (!state.delTargetId) return;
    loading(true);
    try {
      await gasPost('deleteEstimate', { id: state.delTargetId });
      closeModal('delModal');
      toast('Estimate deleted', 'ok');
      await loadEntries();
    } catch (e) { toast('Delete failed: ' + e.message, 'err'); }
    finally { loading(false); state.delTargetId = null; }
  }

  /* ══════════════════════════════════════════════════════════
     SCREEN 2 – EDIT / NEW
  ══════════════════════════════════════════════════════════ */
  function startNew() {
    state.currentId = null;
    clearForm();
    document.getElementById('editTitle').textContent = '✏️ New Estimate';
    document.getElementById('editSub').textContent   = 'Fill client details and select tables';
    document.getElementById('f_estimateDate').value  = new Date().toISOString().split('T')[0];
    document.getElementById('f_estimationNo').value  = 'DOC-' + Date.now().toString().slice(-6);
    // Make sure dropdowns are loaded before showing the edit screen
    if (!state.dropdowns || !state.dropdowns.catA || state.dropdowns.catA.length === 0) {
      loading(true);
      gasGet('getDropdownData').then(function(d) {
        state.dropdowns = d || { catA:[], catB:[], additionalWorks:[], gypsum:[], kitchenAcc:[] };
        loading(false);
        showScreen('edit');
      }).catch(function(e) {
        loading(false);
        toast('Could not load dropdowns: ' + e.message, 'err');
        showScreen('edit');
      });
    } else {
      showScreen('edit');
    }
  }

  function editEstimate(id) {
    const est = state.estimates.find(e => e.id === id);
    if (!est) { toast('Not found', 'err'); return; }
    state.currentId = id;

    document.getElementById('editTitle').textContent = '✏️ Edit – ' + (est.estimationNo || id);
    document.getElementById('editSub').textContent   = 'Client: ' + (est.clientName || '');

    const fields = { clientName:'', estimationNo:'', estimateDate:'', siteAddress:'', mobile:'',
                     contactNumber:'', email:'', bdo:'', interiorDesigner:'', customerCare:'', preparedBy:'' };
    Object.keys(fields).forEach(k => {
      const el = document.getElementById('f_' + k);
      if (el) el.value = est[k] || '';
    });

    const tables = est.data || {};
    Object.keys(state.active).forEach(k => {
      state.active[k] = !!(tables[k] && tables[k].length);
      state.rows[k]   = tables[k] || [];
      tog(k, state.active[k]);
    });
    renderAll();
    showScreen('edit');
  }

  function clearForm() {
    ['clientName','estimationNo','estimateDate','siteAddress','mobile',
     'contactNumber','email','bdo','interiorDesigner','customerCare','preparedBy']
      .forEach(f => { const el = document.getElementById('f_' + f); if (el) el.value = ''; });
    Object.keys(state.active).forEach(k => { state.active[k] = false; state.rows[k] = []; tog(k, false); });
    renderAll(); updateGrand();
  }

  function tog(key, on) {
    const btn = document.getElementById('tog-' + key);
    const blk = document.getElementById('blk-' + key);
    if (btn) btn.classList.toggle('on', on);
    if (blk) blk.classList.toggle('hidden', !on);
  }

  function toggleTable(key) {
    state.active[key] = !state.active[key];
    tog(key, state.active[key]);
    if (state.active[key] && !state.rows[key].length) addRow(key);
    updateGrand();
  }

  /* ── Rows ──────────────────────────────────────────────── */
  const blankRow = {
    catA    : () => ({ category:'', item:'', shutterFinish:'', material:'', unit:'', size:'', sizeX:0, sizeY:0, qtySqft:0, qty:0, price:0, total:0 }),
    kitAcc  : () => ({ product:'', brand:'', articleNo:'', pictureUrl:'', qty:0, price:0, total:0 }),
    addWorks: () => ({ category:'', item:'', shutterFinish:'', material:'', unit:'', size:'', qty:0, price:0, total:0 }),
    catB    : () => ({ category:'', item:'', shutterFinish:'', material:'', unit:'', size:'', qty:0, price:0, total:0 }),
    gypsum  : () => ({ category:'', item:'', finish:'', material:'', unit:'', size:'', price:0, total:0 })
  };

  function addRow(key) {
    state.rows[key].push(blankRow[key]());
    renderTable(key);
  }
  function delRow(key, i) {
    state.rows[key].splice(i, 1);
    renderTable(key);
    updateGrand();
  }

  function renderAll() {
    ['catA','kitAcc','addWorks','catB','gypsum'].forEach(renderTable);
  }

  /* ── Dropdown helpers ──────────────────────────────────── */
  function uniq(arr, k) { return [...new Set(arr.map(r => r[k]).filter(Boolean))]; }
  function sel(opts, cur, fn, w = '') {
    const o = ['<option value="">— Select —</option>',
      ...opts.map(v => `<option value="${esc(v)}"${v === cur ? ' selected' : ''}>${esc(v)}</option>`)
    ].join('');
    return `<select class="ts" style="${w}" onchange="${fn}">${o}</select>`;
  }
  function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;'); }

  /* ── Render functions ──────────────────────────────────── */
  function renderTable(key) {
    switch (key) {
      case 'catA':     renderCatA();     break;
      case 'kitAcc':   renderKitAcc();   break;
      case 'addWorks': renderSimple('addWorks'); break;
      case 'catB':     renderSimple('catB');     break;
      case 'gypsum':   renderGypsum();   break;
    }
    updateTableTot(key);
  }

  /* Category-A */
  function renderCatA() {
    const dd   = state.dropdowns.catA || [];
    const cats = uniq(dd, 'category');
    const shts = uniq(dd, 'shutterFinish');
    const mats = uniq(dd, 'material');
    const uts  = uniq(dd, 'unit');
    document.getElementById('tb-catA').innerHTML = state.rows.catA.map((r, i) => {
      const items = uniq(dd.filter(d => !r.category || d.category === r.category), 'item');
      return `<tr>
        <td>${i+1}</td>
        <td>${sel(cats, r.category, `APP._setCatA(${i},'category',this.value)`)}</td>
        <td>${sel(items, r.item, `APP._setCatA(${i},'item',this.value)`)}</td>
        <td>${sel(shts, r.shutterFinish, `APP._setCatA(${i},'shutterFinish',this.value)`)}</td>
        <td>${sel(mats, r.material,      `APP._setCatA(${i},'material',this.value)`)}</td>
        <td>${sel(uts,  r.unit,          `APP._setCatA(${i},'unit',this.value)`, 'width:70px')}</td>
        <td><input class="ti" type="text"   value="${esc(r.size)}"  oninput="APP._setCatA(${i},'size',this.value)"  style="width:80px" placeholder="e.g.150×235"></td>
        <td><input class="ti" type="number" value="${r.sizeX||0}"   oninput="APP._setCatA(${i},'sizeX',+this.value)" style="width:60px"></td>
        <td><input class="ti" type="number" value="${r.sizeY||0}"   oninput="APP._setCatA(${i},'sizeY',+this.value)" style="width:60px"></td>
        <td><span class="computed">${r.qtySqft||0}</span></td>
        <td><input class="ti" type="number" value="${r.qty||0}"     oninput="APP._setCatA(${i},'qty',+this.value)"   style="width:55px"></td>
        <td><input class="ti" type="number" value="${r.price||0}"   oninput="APP._setCatA(${i},'price',+this.value)" style="width:75px"></td>
        <td><span class="computed">₹${fmt(r.total||0)}</span></td>
        <td><button class="btn-del-row" onclick="APP.delRow('catA',${i})">✕</button></td>
      </tr>`;
    }).join('');
  }

  function _setCatA(i, f, v) {
    const r = state.rows.catA[i];
    r[f] = v;
    if (f === 'item') {
      const m = (state.dropdowns.catA || []).find(d => d.item === v);
      if (m) {
        r.shutterFinish = m.shutterFinish || r.shutterFinish;
        r.material      = m.material      || r.material;
        r.unit          = m.unit          || r.unit;
        if (!r.price) r.price = +m.defaultPrice || 0;
      }
    }
    if (f === 'sizeX' || f === 'sizeY') {
      r.qtySqft = Math.round((r.sizeX * r.sizeY) / 950);
    }
    r.total = r.qtySqft ? r.qtySqft * r.price : r.qty * r.price;
    renderCatA();
  }

  /* Kitchen Accessories */
  function renderKitAcc() {
    const dd   = state.dropdowns.kitchenAcc || [];
    const prods = uniq(dd, 'product');
    document.getElementById('tb-kitAcc').innerHTML = state.rows.kitAcc.map((r, i) => `<tr>
      <td>${i+1}</td>
      <td>${sel(prods, r.product, `APP._setKit(${i},'product',this.value)`, 'min-width:160px')}</td>
      <td><input class="ti" type="text" value="${esc(r.brand)}"    readonly style="background:var(--bg);width:80px"></td>
      <td><input class="ti" type="text" value="${esc(r.articleNo)}" readonly style="background:var(--bg);width:90px"></td>
      <td><input class="ti" type="text" value="${esc(r.pictureUrl)}" oninput="APP._setKit(${i},'pictureUrl',this.value)" style="width:110px" placeholder="URL"></td>
      <td><input class="ti" type="number" value="${r.qty  ||0}" oninput="APP._setKit(${i},'qty',+this.value)"   style="width:55px"></td>
      <td><input class="ti" type="number" value="${r.price||0}" oninput="APP._setKit(${i},'price',+this.value)" style="width:80px"></td>
      <td><span class="computed">₹${fmt(r.total||0)}</span></td>
      <td><button class="btn-del-row" onclick="APP.delRow('kitAcc',${i})">✕</button></td>
    </tr>`).join('');
  }

  function _setKit(i, f, v) {
    const r = state.rows.kitAcc[i];
    r[f] = v;
    if (f === 'product') {
      const m = (state.dropdowns.kitchenAcc || []).find(d => d.product === v);
      if (m) { r.brand = m.brand||''; r.articleNo = m.articleNo||''; r.pictureUrl = m.pictureUrl||''; if (!r.price) r.price = +m.defaultPrice||0; }
    }
    r.total = (r.qty||0) * (r.price||0);
    renderKitAcc();
  }

  /* Simple (addWorks / catB) */
  function renderSimple(key) {
    const ddKey = key === 'addWorks' ? 'additionalWorks' : 'catB';
    const dd    = state.dropdowns[ddKey] || [];
    const cats  = uniq(dd, 'category');
    const shts  = uniq(dd, 'shutterFinish');
    const mats  = uniq(dd, 'material');
    const uts   = uniq(dd, 'unit');
    document.getElementById('tb-' + key).innerHTML = state.rows[key].map((r, i) => {
      const items = uniq(dd.filter(d => !r.category || d.category === r.category), 'item');
      return `<tr>
        <td>${i+1}</td>
        <td>${sel(cats, r.category,     `APP._setSimple('${key}',${i},'category',this.value)`)}</td>
        <td>${sel(items,r.item,          `APP._setSimple('${key}',${i},'item',this.value)`)}</td>
        <td>${sel(shts, r.shutterFinish, `APP._setSimple('${key}',${i},'shutterFinish',this.value)`)}</td>
        <td>${sel(mats, r.material,      `APP._setSimple('${key}',${i},'material',this.value)`)}</td>
        <td>${sel(uts,  r.unit,          `APP._setSimple('${key}',${i},'unit',this.value)`, 'width:70px')}</td>
        <td><input class="ti" type="text"   value="${esc(r.size)}"  oninput="APP._setSimple('${key}',${i},'size',this.value)"  style="width:90px"></td>
        <td><input class="ti" type="number" value="${r.qty  ||0}" oninput="APP._setSimple('${key}',${i},'qty',+this.value)"   style="width:60px"></td>
        <td><input class="ti" type="number" value="${r.price||0}" oninput="APP._setSimple('${key}',${i},'price',+this.value)" style="width:80px"></td>
        <td><span class="computed">₹${fmt(r.total||0)}</span></td>
        <td><button class="btn-del-row" onclick="APP.delRow('${key}',${i})">✕</button></td>
      </tr>`;
    }).join('');
  }

  function _setSimple(key, i, f, v) {
    const ddKey = key === 'addWorks' ? 'additionalWorks' : 'catB';
    const r = state.rows[key][i];
    r[f] = v;
    if (f === 'item') {
      const m = (state.dropdowns[ddKey] || []).find(d => d.item === v);
      if (m && !r.price) r.price = +m.defaultPrice || 0;
    }
    r.total = (r.qty||0) * (r.price||0);
    renderSimple(key);
  }

  /* Gypsum */
  function renderGypsum() {
    const dd   = state.dropdowns.gypsum || [];
    const cats = uniq(dd, 'category');
    const fins = uniq(dd, 'finish');
    const mats = uniq(dd, 'material');
    const uts  = uniq(dd, 'unit');
    document.getElementById('tb-gypsum').innerHTML = state.rows.gypsum.map((r, i) => {
      const items = uniq(dd.filter(d => !r.category || d.category === r.category), 'item');
      return `<tr>
        <td>${i+1}</td>
        <td>${sel(cats, r.category, `APP._setGypsum(${i},'category',this.value)`)}</td>
        <td>${sel(items,r.item,     `APP._setGypsum(${i},'item',this.value)`)}</td>
        <td>${sel(fins, r.finish,   `APP._setGypsum(${i},'finish',this.value)`)}</td>
        <td>${sel(mats, r.material, `APP._setGypsum(${i},'material',this.value)`)}</td>
        <td>${sel(uts,  r.unit,     `APP._setGypsum(${i},'unit',this.value)`, 'width:70px')}</td>
        <td><input class="ti" type="text"   value="${esc(r.size)}"  oninput="APP._setGypsum(${i},'size',this.value)"  style="width:90px"></td>
        <td><input class="ti" type="number" value="${r.price||0}" oninput="APP._setGypsum(${i},'price',+this.value)" style="width:80px"></td>
        <td><span class="computed">₹${fmt(r.total||0)}</span></td>
        <td><button class="btn-del-row" onclick="APP.delRow('gypsum',${i})">✕</button></td>
      </tr>`;
    }).join('');
  }

  function _setGypsum(i, f, v) {
    const r = state.rows.gypsum[i];
    r[f] = v;
    if (f === 'item') {
      const m = (state.dropdowns.gypsum || []).find(d => d.item === v);
      if (m && !r.price) r.price = +m.defaultPrice || 0;
    }
    r.total = r.price || 0;
    renderGypsum();
  }

  /* ── Totals ────────────────────────────────────────────── */
  function updateTableTot(key) {
    const tot = state.rows[key].reduce((s, r) => s + (r.total||0), 0);
    const el  = document.getElementById('tot-' + key);
    if (el) el.textContent = '₹' + fmt(tot);
    updateGrand();
  }
  function updateGrand() {
    let g = 0;
    Object.keys(state.active).forEach(k => { if (state.active[k]) g += state.rows[k].reduce((s,r)=>s+(r.total||0),0); });
    document.getElementById('grandTotal').textContent = '₹' + fmt(g);
  }
  function fmt(n) { return Number(n).toLocaleString('en-IN', { maximumFractionDigits:0 }); }

  /* ── Save Estimate ─────────────────────────────────────── */
  async function saveEstimate() {
    const clientName = document.getElementById('f_clientName').value.trim();
    if (!clientName) { toast('Client name is required', 'err'); return; }

    const payload = {
      id    : state.currentId,
      client: {
        clientName,
        estimationNo    : document.getElementById('f_estimationNo').value,
        estimateDate    : document.getElementById('f_estimateDate').value,
        siteAddress     : document.getElementById('f_siteAddress').value,
        mobile          : document.getElementById('f_mobile').value,
        contactNumber   : document.getElementById('f_contactNumber').value,
        email           : document.getElementById('f_email').value,
        bdo             : document.getElementById('f_bdo').value,
        interiorDesigner: document.getElementById('f_interiorDesigner').value,
        customerCare    : document.getElementById('f_customerCare').value,
        preparedBy      : document.getElementById('f_preparedBy').value
      },
      tables: {}
    };
    Object.keys(state.active).forEach(k => { if (state.active[k]) payload.tables[k] = state.rows[k]; });

    loading(true);
    try {
      const res = await gasPost('saveEstimate', payload);
      state.currentId = res.id;
      toast('Estimate saved!', 'ok');
      await loadEntries();
    } catch (e) { toast('Save failed: ' + e.message, 'err'); }
    finally { loading(false); }
  }

  /* ══════════════════════════════════════════════════════════
     SCREEN 3 – PREVIEW
  ══════════════════════════════════════════════════════════ */
  function previewById(id) {
    const est = state.estimates.find(e => e.id === id);
    if (!est) { toast('Not found', 'err'); return; }
    // Populate client fields silently
    state.currentId = id;
    const tables = est.data || {};
    Object.keys(state.active).forEach(k => { state.active[k] = !!(tables[k] && tables[k].length); state.rows[k] = tables[k] || []; });
    const cf = ['clientName','estimationNo','estimateDate','siteAddress','mobile','contactNumber','email','bdo','interiorDesigner','customerCare','preparedBy'];
    cf.forEach(k => { const el = document.getElementById('f_' + k); if (el) el.value = est[k] || ''; });
    showScreen('preview');
  }

  function renderPreview() {
    const c = {
      clientName      : document.getElementById('f_clientName').value      || '—',
      estimationNo    : document.getElementById('f_estimationNo').value    || '—',
      estimateDate    : document.getElementById('f_estimateDate').value    || '—',
      siteAddress     : document.getElementById('f_siteAddress').value     || '—',
      mobile          : document.getElementById('f_mobile').value          || '—',
      contactNumber   : document.getElementById('f_contactNumber').value   || '—',
      email           : document.getElementById('f_email').value           || '—',
      bdo             : document.getElementById('f_bdo').value             || '—',
      interiorDesigner: document.getElementById('f_interiorDesigner').value|| '—',
      customerCare    : document.getElementById('f_customerCare').value    || '+91 8943807777',
      preparedBy      : document.getElementById('f_preparedBy').value      || '—'
    };

    const { sectionsHtml, subtotals, grandTotal } = buildEstSections();

    document.getElementById('previewWrap').innerHTML = `

      <!-- PAGE 1: COVER -->
      <div class="a4">
        <div class="p1-content"><div class="p1-title">ESTIMATE</div></div>
        <div class="top-hdr">
          <div class="bar-left"></div>
          <div class="logo-box">
            <div style="width:75px;height:36px;background:var(--gold);border-radius:4px;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:14px;">DOC</div>
            <div class="logo-name">DOC INTERIORS</div>
          </div>
          <div class="bar-right"></div>
        </div>
        <div class="hero-img">
          <div class="hero-overlay">
            <div class="hero-txt"><span class="fl">D</span>EPTH<br><span class="fl">O</span>F<br><span class="fl">C</span>REATION</div>
          </div>
        </div>
        <div class="btm-band">
          <div class="btm-bar"></div>
          <div class="btm-cnt">
            <div class="btm-tl">100% CUSTOMIZED CONTEMPORARY</div>
            <div class="btm-cats">KITCHEN <span>|</span> BEDROOM <span>|</span> LIVING <span>|</span> DINING</div>
          </div>
          <div class="site-url">www.docinteriors.com<div class="underline"></div></div>
        </div>
      </div>

      <!-- PAGE 2: INTRO -->
      <div class="a4">
        <div style="height:150mm;background:linear-gradient(135deg,#0d0d18,#2a2a40);display:flex;align-items:center;justify-content:center;">
          <div style="color:var(--gold);font-size:26px;font-family:serif;letter-spacing:4px;text-align:center;">DOC INTERIORS<br><span style="font-size:13px;letter-spacing:6px;opacity:.7;">DEPTH OF CREATION</span></div>
        </div>
        <div style="padding:10mm 14mm;font-family:Arial;">
          <h3 style="font-size:15px;margin-bottom:8px;">Thank you for considering DOC Interiors.</h3>
          <p style="font-size:12px;color:#444;line-height:1.6;">We can't wait to bring your vision to life. Our team will deliver quality craftsmanship within 35–45 working days.</p>
        </div>
        <div style="padding:0 10mm;display:flex;flex-wrap:wrap;gap:6px;justify-content:center;">
          ${['Consultation','Drawing & Approval','Factory Production','Delivery','Installation','Handover'].map((s,i)=>`<div style="text-align:center;width:75px;padding:6px;"><div style="width:34px;height:34px;background:var(--gold);border-radius:50%;display:inline-flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:13px;margin-bottom:4px;">${i+1}</div><p style="font-size:9px;color:#555;">${s}</p></div>`).join('')}
        </div>
        <div style="text-align:right;padding:5mm 12mm;font-size:12px;font-family:Arial;">www.docinteriors.com<div class="underline" style="margin-left:auto;"></div></div>
      </div>

      <!-- PAGE 3: CLIENT INFO -->
      <div class="a4">
        <div class="pg3-hero"><div class="pg3-hero-txt">depth of<br>creation</div></div>
        <div class="pg-inner">
          <table class="info-tbl">
            <tr><td>DATE<br><span>${c.estimateDate}</span></td>       <td>ESTIMATION NO.<br><span>${c.estimationNo}</span></td></tr>
            <tr><td>CLIENT<br><span>${c.clientName}</span></td>        <td>BDO<br><span>${c.bdo}</span></td></tr>
            <tr><td>SITE ADDRESS<br><span>${c.siteAddress}</span></td> <td>MOBILE<br><span>${c.mobile}</span></td></tr>
            <tr><td>E-MAIL<br><span>${c.email}</span></td>             <td>CONTACT<br><span>${c.contactNumber}</span></td></tr>
            <tr><td>INTERIOR DESIGNER<br><span>${c.interiorDesigner}</span></td><td>PREPARED BY<br><span>${c.preparedBy}</span></td></tr>
            <tr><td class="hl">CUSTOMER CARE<br><span>${c.customerCare}</span></td><td></td></tr>
          </table>
        </div>
        <div style="text-align:right;padding:2mm 12mm;font-size:12px;font-family:Arial;">www.docinteriors.com<div class="underline" style="margin-left:auto;"></div></div>
      </div>

      ${sectionsHtml}

      <!-- SUMMARY PAGE -->
      <div class="a4" style="padding:10mm;">
        <div class="sub-hdr" style="margin:-10mm -10mm 8mm -10mm;padding:7px 10mm;">ESTIMATE SUMMARY</div>
        <table class="summary-tbl">
          <thead><tr><th>Section</th><th style="text-align:right">Sub-Total</th></tr></thead>
          <tbody>
            ${subtotals.map(s=>`<tr><td>${s.label}</td><td style="text-align:right;font-weight:600">₹${fmt(s.value)}</td></tr>`).join('')}
          </tbody>
        </table>
        <div class="grand-bar">GRAND TOTAL : ₹${fmt(grandTotal)}</div>
        <div style="margin-top:8mm;font-size:11px;color:#666;font-family:'Times New Roman',serif;line-height:1.8;">
          <p>This estimate is valid for <strong>30 days</strong> from the date of issue.</p>
          <p style="margin-top:6mm;">Authorized Signatory : ___________________________</p>
        </div>
      </div>`;
  }

  function buildEstSections() {
    const TABLES = [
      { key:'catA',     label:'CATEGORY – A',         cols:['SL','ITEM','SHUTTER FINISH','MATERIAL','UNIT','SIZE','QTY SQFT','QTY','PRICE','TOTAL'] },
      { key:'kitAcc',   label:'KITCHEN ACCESSORIES',   cols:['SL','PRODUCT','BRAND','ARTICLE NO','QTY','PRICE','TOTAL'] },
      { key:'addWorks', label:'ADDITIONAL WORKS',       cols:['SL','ITEM','SHUTTER FINISH','MATERIAL','UNIT','SIZE','QTY','PRICE','TOTAL'] },
      { key:'catB',     label:'CATEGORY – B',          cols:['SL','ITEM','SHUTTER FINISH','MATERIAL','UNIT','SIZE','QTY','PRICE','TOTAL'] },
      { key:'gypsum',   label:'GYPSUM CEILING WORK',   cols:['SL','ITEM','FINISH','MATERIAL','UNIT','SIZE','PRICE','TOTAL'] }
    ];
    let html = '', subtotals = [], grand = 0;

    TABLES.forEach(t => {
      if (!state.active[t.key] || !state.rows[t.key].length) return;
      const rows = state.rows[t.key];
      const sub  = rows.reduce((s, r) => s + (r.total||0), 0);
      grand += sub;
      subtotals.push({ label: t.label, value: sub });

      const bodyRows = rows.map((r, i) => {
        let cells = `<td>${i+1}</td>`;
        if      (t.key === 'catA')   cells += `<td style="text-align:left">${r.item||'—'}</td><td>${r.shutterFinish||'—'}</td><td>${r.material||'—'}</td><td>${r.unit||'—'}</td><td>${r.size||'—'}</td><td>${r.qtySqft||0}</td><td>${r.qty||0}</td><td>₹${fmt(r.price||0)}</td>`;
        else if (t.key === 'kitAcc') cells += `<td style="text-align:left">${r.product||'—'}</td><td>${r.brand||'—'}</td><td>${r.articleNo||'—'}</td><td>${r.qty||0}</td><td>₹${fmt(r.price||0)}</td>`;
        else if (t.key === 'gypsum') cells += `<td style="text-align:left">${r.item||'—'}</td><td>${r.finish||'—'}</td><td>${r.material||'—'}</td><td>${r.unit||'—'}</td><td>${r.size||'—'}</td><td>₹${fmt(r.price||0)}</td>`;
        else                         cells += `<td style="text-align:left">${r.item||'—'}</td><td>${r.shutterFinish||'—'}</td><td>${r.material||'—'}</td><td>${r.unit||'—'}</td><td>${r.size||'—'}</td><td>${r.qty||0}</td><td>₹${fmt(r.price||0)}</td>`;
        cells += `<td style="font-weight:600">₹${fmt(r.total||0)}</td>`;
        return `<tr>${cells}</tr>`;
      }).join('');

      html += `
        <div class="a4">
          <div class="sub-hdr">SUB: ESTIMATE FOR WOOD WORK, ACCESSORIES AND BEAUTIFICATION</div>
          <div class="pg4-hero"></div>
          <div class="pg4-wrap">
            <span class="cat-badge">${t.label}</span>
            <table class="prev-tbl">
              <thead><tr>${t.cols.map(c=>`<th>${c}</th>`).join('')}</tr></thead>
              <tbody>${bodyRows}</tbody>
              <tfoot><tr class="prev-tot-row">
                <td colspan="${t.cols.length-1}" style="text-align:right;padding-right:12px;letter-spacing:1px;">${t.label} TOTAL</td>
                <td>₹${fmt(sub)}</td>
              </tr></tfoot>
            </table>
          </div>
          <div style="text-align:right;padding:2mm 12mm;font-size:12px;font-family:Arial;">www.docinteriors.com<div class="underline" style="margin-left:auto;"></div></div>
        </div>`;
    });

    return { sectionsHtml: html, subtotals, grandTotal: grand };
  }

  /* ══════════════════════════════════════════════════════════
     SCREEN 4 – ADMIN
  ══════════════════════════════════════════════════════════ */
  async function loadAdminData() {
    loading(true);
    try {
      state.adminData = await gasGet('getAdminDropdownData') || {};
      renderAdminTable(state.adminTab);
    } catch (e) { toast('Admin load error: ' + e.message, 'err'); }
    finally { loading(false); }
  }

  function switchAdminTab(key) {
    state.adminTab = key;
    document.querySelectorAll('.atab').forEach(t => t.classList.remove('active'));
    document.getElementById('atab-' + key).classList.add('active');
    renderAdminTable(key);
  }

  function renderAdminTable(key) {
    const section = state.adminData[key];
    const el = document.getElementById('adminContent');
    if (!section || !section.data) { el.innerHTML = '<div class="empty">No data</div>'; return; }
    const { cols, data } = section;
    if (!data.length) { el.innerHTML = `<div class="empty">No items. <button class="btn btn-primary btn-sm" onclick="APP.showAddModal()">+ Add First</button></div>`; return; }
    const dataKeys = Object.keys(data[0]).filter(k => k !== '_rowIndex');
    el.innerHTML = `
      <div class="table-scroll">
        <table class="admin-tbl">
          <thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}<th>Actions</th></tr></thead>
          <tbody>
            ${data.map((row, i) => `<tr>
              ${dataKeys.map(k=>`<td><input class="ai" value="${esc(row[k]||'')}" onchange="APP._adminCell('${key}',${i},'${k}',this.value)"></td>`).join('')}
              <td style="white-space:nowrap">
                <button class="btn btn-success btn-sm" onclick="APP.saveAdminRow('${key}',${i})">💾</button>
                <button class="btn btn-danger  btn-sm" style="margin-left:4px" onclick="APP.showAdminDel('${key}',${i})">🗑</button>
              </td>
            </tr>`).join('')}
          </tbody>
        </table>
      </div>`;
  }

  function _adminCell(key, i, field, val) { state.adminData[key].data[i][field] = val; }

  async function saveAdminRow(key, i) {
    const section  = state.adminData[key];
    const row      = section.data[i];
    const rowIndex = row._rowIndex;
    const item     = Object.fromEntries(Object.entries(row).filter(([k]) => k !== '_rowIndex'));
    loading(true);
    try {
      await gasPost('updateDropdownItem', { sheetName: section.sheetName, rowIndex, item });
      toast('Row saved!', 'ok');
      await refreshDropdowns();
    } catch (e) { toast('Save failed: ' + e.message, 'err'); }
    finally { loading(false); }
  }

  /* Add/Edit modal */
  function showAddModal(editKey, editIdx) {
    const key = editKey || state.adminTab;
    const section = state.adminData[key];
    if (!section) return;
    state.modalEdit = editIdx !== undefined ? { key, idx: editIdx } : null;

    document.getElementById('addModalTitle').textContent  = state.modalEdit ? 'Edit Item' : 'Add New Item';
    document.getElementById('addModalSubmit').textContent = state.modalEdit ? 'Save Changes' : 'Add Item';

    const dataKeys = section.data.length ? Object.keys(section.data[0]).filter(k => k !== '_rowIndex') : section.cols.map((_,i)=>i);
    const editRow  = state.modalEdit ? section.data[editIdx] : null;
    document.getElementById('addModalFields').innerHTML = section.cols.map((col, i) => {
      const k   = dataKeys[i] || i;
      const val = editRow ? (editRow[k]||'') : '';
      return `<div class="fg"><label>${col}</label><input class="fc" id="addFld_${k}" value="${esc(val)}" placeholder="${col}"></div>`;
    }).join('');
    document.getElementById('addModal').classList.remove('hidden');
  }

  async function submitModal() {
    const key     = state.modalEdit ? state.modalEdit.key : state.adminTab;
    const section = state.adminData[key];
    const dataKeys = section.data.length ? Object.keys(section.data[0]).filter(k => k !== '_rowIndex') : section.cols.map((_,i)=>i);
    const item    = {};
    dataKeys.forEach(k => { const el = document.getElementById('addFld_' + k); if (el) item[k] = el.value; });
    if (!Object.values(item)[0]) { toast('First field is required', 'err'); return; }

    loading(true);
    try {
      if (state.modalEdit) {
        const rowIndex = section.data[state.modalEdit.idx]._rowIndex;
        await gasPost('updateDropdownItem', { sheetName: section.sheetName, rowIndex, item });
        toast('Item updated!', 'ok');
      } else {
        await gasPost('addDropdownItem', { sheetName: section.sheetName, item });
        toast('Item added!', 'ok');
      }
      closeModal('addModal');
      await loadAdminData();
      await refreshDropdowns();
    } catch (e) { toast('Error: ' + e.message, 'err'); }
    finally { loading(false); }
  }

  /* Admin delete */
  function showAdminDel(key, i) {
    state.adminDel = { key, i };
    document.getElementById('adminDelModal').classList.remove('hidden');
  }
  async function confirmAdminDelete() {
    if (!state.adminDel) return;
    const { key, i }   = state.adminDel;
    const section       = state.adminData[key];
    const rowIndex      = section.data[i]._rowIndex;
    loading(true);
    try {
      await gasPost('deleteDropdownItem', { sheetName: section.sheetName, rowIndex });
      toast('Item deleted', 'ok');
      closeModal('adminDelModal');
      await loadAdminData();
      await refreshDropdowns();
    } catch (e) { toast('Delete failed: ' + e.message, 'err'); }
    finally { loading(false); state.adminDel = null; }
  }

  async function refreshDropdowns() {
    try { state.dropdowns = await gasGet('getDropdownData'); } catch {}
  }

  /* ── Shared ────────────────────────────────────────────── */
  function closeModal(id) { document.getElementById(id).classList.add('hidden'); }

  /* ══════════════════════════════════════════════════════════
     BOOTSTRAP
     ─────────────────────────────────────────────────────────
     On every page load:
       • If no GAS URL in localStorage → show setup banner
       • If GAS URL exists → call getConfig to get/refresh the
         Spreadsheet ID from Code.gs, then load the app.
     The user only ever enters the GAS Web App URL once.
     The Spreadsheet ID always comes from Code.gs automatically.
  ══════════════════════════════════════════════════════════ */
  window.addEventListener('DOMContentLoaded', async () => {
    document.getElementById('f_estimateDate').value = new Date().toISOString().split('T')[0];

    const gasUrl = getGasUrl();

    if (!gasUrl) {
      // First time — show the setup banner
      document.getElementById('configBanner').classList.remove('hidden');
      return;
    }

    // GAS URL exists — silently fetch config to get Spreadsheet ID
    loading(true);
    try {
      const cfg = await gasGet('getConfig');
      localStorage.setItem(KEY_SS_ID,  cfg.spreadsheetId);
      localStorage.setItem(KEY_SS_URL, cfg.spreadsheetUrl);
      updateStatusBar(cfg.spreadsheetId, cfg.spreadsheetUrl);
      await init();
    } catch (err) {
      loading(false);
      // Config fetch failed — show the banner so user can re-enter URL
      document.getElementById('configBanner').classList.remove('hidden');
      document.getElementById('gasUrlInput').value = gasUrl;
      toast('Could not reach backend: ' + err.message + ' — please check your GAS URL', 'err');
    }
  });

  /* Public API (called from HTML onclick attributes) */
  return {
    showScreen, saveGasUrl, openConfig,
    loadEntries, filterEntries, startNew, editEstimate, saveEstimate, previewById,
    showDelModal, confirmDelete, closeModal,
    toggleTable, addRow, delRow,
    switchAdminTab, showAddModal, saveAdminRow, submitModal,
    showAdminDel, confirmAdminDelete,
    // internal mutators exposed for inline onchange handlers
    _setCatA   : _setCatA,
    _setKit    : _setKit,
    _setSimple : _setSimple,
    _setGypsum : _setGypsum,
    _adminCell : _adminCell
  };

})();
