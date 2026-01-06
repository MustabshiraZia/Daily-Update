// script.js — cleaned and fixed version (localStorage persistence removed, robust save, delete tab)
const firebaseConfig = {
  apiKey: "AIzaSyBrjcpI8EMT8tnfSCS82PP9FmFw1A7hv6Y",
  authDomain: "uoodailyupdates.firebaseapp.com",
  databaseURL: "https://uoodailyupdates-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "uoodailyupdates",
  storageBucket: "uoodailyupdates.firebasestorage.app",
  messagingSenderId: "978347462324",
  appId: "1:978347462324:web:770a4de66d0c98d7d8d1e1",
  measurementId: "G-KZ866SZ4H4"
};

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.6.0/firebase-app.js";
import {
  getDatabase,
  ref,
  push,
  set,
  update,
  remove,
  onValue,
  get
} from "https://www.gstatic.com/firebasejs/10.6.0/firebase-database.js";
import { getAuth, signInAnonymously, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/10.6.0/firebase-auth.js";

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);
const auth = getAuth(app);
let headersUnsub = null;
let rowsUnsub = null;

// DOM refs
const sheetsListEl = document.getElementById('sheetsList');
const addTabBtn = document.getElementById('addTab');
const thead = document.getElementById('peopleHead');
const tbody = document.getElementById('peopleBody');
const addRowBtn = document.getElementById('addRow');
const addColumnBtn = document.getElementById('addColumn');
const newColumnTypeSelect = document.getElementById('newColumnType');
const clearTableBtn = document.getElementById('clearTable');
const exportCsvBtn = document.getElementById('exportCsv');
const pageTitle = document.getElementById('pageTitle');
const toastEl = document.getElementById('toast');

function toast(msg, t=2000){ if (!toastEl) return; toastEl.textContent = msg; toastEl.classList.remove('hidden'); setTimeout(()=> toastEl.classList.add('hidden'), t); }
const esc = s => String(s||'').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
function dbg(...a){ try{ console.log('[app]', ...a); } catch(e){} }
function debounce(fn, wait=600){ let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), wait); } }

// state
let sheets = [];
let activeSheetId = null;
let headers = []; // {id, title, order, type, options}
const rowsMap = new Map();

const SHEETS_ROOT = 'sheets';
const sheetHeadersPath = sid => `${SHEETS_ROOT}/${sid}/headers`;
const sheetRowsPath = sid => `${SHEETS_ROOT}/${sid}/rows`;

// defensive
if (!sheetsListEl) dbg('Warning: sheets list element not found (id=sheetsList)');
if (!addColumnBtn) dbg('Warning: addColumn button not found (id=addColumn)');
if (!thead || !tbody) dbg('Warning: table elements missing.');

// --- localStorage persistence DISABLED to avoid phantom duplicate rows ---
const LOCAL_KEY_PREFIX = 'uoo_sheet_locals_';
(function clearLocalCacheOnLoad() {
  try {
    for (let i = localStorage.length - 1; i >= 0; i--) {
      const key = localStorage.key(i);
      if (key && key.startsWith(LOCAL_KEY_PREFIX)) {
        localStorage.removeItem(key);
      }
    }
    dbg('Cleared sheet-local cache on load');
  } catch (e) {
    console.warn('Failed to clear local cache on load', e);
  }
})();

function localStorageKeyForSheet(sid) {
  return LOCAL_KEY_PREFIX + (sid || 'no_sheet');
}
function collectLocalRowsFromDOM() { return []; }
function saveLocalRowsToStorage(sid) { /* no-op */ }
function loadLocalRowsFromStorage(sid) { return []; }
const debouncedPersistLocal = () => { /* no-op */ };

// ---------- save helpers ----------
async function saveRowElementIfDirty(tr) {
  if (!tr) return false;
  if (!tr.dataset.dirty) return false;
  if (typeof tr._saveRow === 'function') {
    return await tr._saveRow();
  }
  return false;
}
async function saveAllDirtyRows() {
  if (!activeSheetId) return;
  const trs = Array.from(document.querySelectorAll('tbody tr'));
  const dirty = trs.filter(t => t.dataset.dirty);
  if (!dirty.length) return;
  for (const tr of dirty) {
    if (tr._saving) { dbg('saveAllDirtyRows: skipping row (already saving)'); continue; }
    try { await saveRowElementIfDirty(tr); } catch (err) { console.error('saveAllDirtyRows: failed to save row', err); }
  }
}

// autosave
const AUTOSAVE_INTERVAL_MS = 30_000;
setInterval(()=> { if (activeSheetId) saveAllDirtyRows().catch(e=>console.error('Autosave err',e)); }, AUTOSAVE_INTERVAL_MS);
window.addEventListener('beforeunload', ()=> { try { saveAllDirtyRows(); } catch(_){} });

// ---------- default sheet ----------
async function createDefaultSheet(){
  dbg('createDefaultSheet called');
  const newSheetRef = push(ref(db, SHEETS_ROOT));
  const sid = newSheetRef.key;
  const updates = {};
  updates[`${SHEETS_ROOT}/${sid}/meta/title`] = 'People';
  const h1 = push(ref(db, `${SHEETS_ROOT}/${sid}/headers`));
  const h2 = push(ref(db, `${SHEETS_ROOT}/${sid}/headers`));
  updates[`${SHEETS_ROOT}/${sid}/headers/${h1.key}/title`] = 'Name';
  updates[`${SHEETS_ROOT}/${sid}/headers/${h1.key}/order`] = 0;
  updates[`${SHEETS_ROOT}/${sid}/headers/${h1.key}/type`] = 'string';
  updates[`${SHEETS_ROOT}/${sid}/headers/${h2.key}/title`] = 'Status';
  updates[`${SHEETS_ROOT}/${sid}/headers/${h2.key}/order`] = 1;
  updates[`${SHEETS_ROOT}/${sid}/headers/${h2.key}/type`] = 'string';
  const r1 = push(ref(db, `${SHEETS_ROOT}/${sid}/rows`));
  const r2 = push(ref(db, `${SHEETS_ROOT}/${sid}/rows`));
  updates[`${SHEETS_ROOT}/${sid}/rows/${r1.key}/cells/${h1.key}`] = 'Alice';
  updates[`${SHEETS_ROOT}/${sid}/rows/${r1.key}/cells/${h2.key}`] = 'Working';
  updates[`${SHEETS_ROOT}/${sid}/rows/${r1.key}/updatedAt`] = new Date().toISOString();
  updates[`${SHEETS_ROOT}/${sid}/rows/${r2.key}/cells/${h1.key}`] = 'Bob';
  updates[`${SHEETS_ROOT}/${sid}/rows/${r2.key}/cells/${h2.key}`] = 'Remote';
  updates[`${SHEETS_ROOT}/${sid}/rows/${r2.key}/updatedAt`] = new Date().toISOString();
  await update(ref(db), updates);
  dbg('Default sheet created', sid);
  return sid;
}

// ---------- subscriptions ----------
function renderSheetsList(){
  if (!sheetsListEl) return;
  sheetsListEl.innerHTML = '';
  sheets.forEach(s=>{
    // wrapper button element for sheet selection
    const wrapper = document.createElement('div');
    wrapper.className = 'sheet-row';
    wrapper.style.display = 'flex';
    wrapper.style.alignItems = 'center';
    wrapper.style.gap = '8px';
    wrapper.style.marginBottom = '6px';

    const btn = document.createElement('button');
    btn.className = 'sheet-btn' + (s.id===activeSheetId ? ' active' : '');
    btn.textContent = s.title || 'Untitled';
    btn.dataset.sid = s.id;
    btn.style.flex = '1';
    btn.style.textAlign = 'left';
    btn.addEventListener('click', ()=> selectSheet(s.id));

    // delete icon
    const del = document.createElement('button');
    del.title = 'Delete sheet';
    del.innerText = '🗑';
    del.style.border = 'none';
    del.style.background = 'transparent';
    del.style.cursor = 'pointer';
    del.style.padding = '4px 6px';
    del.style.fontSize = '14px';
    del.addEventListener('click', async (ev) => {
      ev.stopPropagation();
      const sid = s.id;
      if (!confirm(`Delete sheet "${s.title || sid}" and all its data? This cannot be undone.`)) return;
      try {
        // If deleting active, clear local active to avoid UI race
        const wasActive = (activeSheetId === sid);
        await remove(ref(db, `${SHEETS_ROOT}/${sid}`));
        toast('Sheet deleted');
        dbg('Deleted sheet', sid);
        if (wasActive) {
          activeSheetId = null; // onValue will choose another or create default
          // clear UI immediately
          if (thead) thead.innerHTML = '';
          if (tbody) tbody.innerHTML = '';
        }
      } catch (err) {
        console.error('Failed to delete sheet', err);
        toast('Delete failed (see console)');
      }
    });

    wrapper.appendChild(btn);
    wrapper.appendChild(del);
    sheetsListEl.appendChild(wrapper);
  });
}

function subscribeSheetsList(){
  dbg('subscribeSheetsList registering onValue for', SHEETS_ROOT);
  onValue(ref(db, SHEETS_ROOT), snap => {
    const val = snap.val() || {};
    sheets = Object.entries(val).map(([id,obj])=>({ id, title: obj?.meta?.title || obj?.meta || ('Sheet ' + id) }));
    dbg('sheets snapshot', sheets);
    if (!sheets.length){
      createDefaultSheet().catch(err => { console.error('create default sheet failed', err); toast('Unable to create default sheet'); });
      return;
    }
    // choose active sheet if none or if active was deleted
    if (!activeSheetId) selectSheet(sheets[0].id);
    else if (!sheets.find(s=>s.id===activeSheetId)) selectSheet(sheets[0].id);
    else renderSheetsList();
  }, err => {
    console.error('subscribeSheetsList onValue error', err);
    toast('Database error (see console)');
  });
}

function unsubscribeActiveSheet(){
  try {
    if (typeof headersUnsub === 'function') { headersUnsub(); headersUnsub = null; }
    if (typeof rowsUnsub === 'function') { rowsUnsub(); rowsUnsub = null; }
  } catch(e){ console.warn('Error unsubscribing', e); }

  headers = [];
  rowsMap.clear();
  if (thead) thead.innerHTML = '';
  if (tbody) tbody.innerHTML = '';
}

function subscribeToActiveSheet(){
  if (!activeSheetId) { dbg('subscribeToActiveSheet: no activeSheetId'); return; }
  unsubscribeActiveSheet();
  dbg('subscribeToActiveSheet for', activeSheetId);

  headersUnsub = onValue(ref(db, sheetHeadersPath(activeSheetId)), snap => {
    const raw = snap.val() || {};
    headers = Object.entries(raw).map(([id,obj])=>({
      id,
      title: obj?.title || '',
      order: (obj?.order ?? 0),
      type: obj?.type || 'string',
      options: Array.isArray(obj?.options) ? obj.options : (obj?.options ? Object.values(obj.options) : [])
    }));
    headers.sort((a,b)=> a.order - b.order);
    dbg('headers loaded', headers);
    renderHead();
  }, err => { console.error('headers onValue err', err); });

  rowsUnsub = onValue(ref(db, sheetRowsPath(activeSheetId)), snap => {
    const raw = snap.val() || {};
    const arr = Object.entries(raw).map(([k,v])=>({ key:k, value:v }));
    // ascending by updatedAt so new rows appear at the bottom
    arr.sort((a,b)=>{
      const ta = a.value?.updatedAt ? new Date(a.value.updatedAt).getTime() : 0;
      const tb = b.value?.updatedAt ? new Date(b.value.updatedAt).getTime() : 0;
      return ta - tb;
    });
    renderBody(arr);
  }, err => { console.error('rows onValue err', err); });
}

// ---------- rendering head ----------
function renderHead() {
  thead.innerHTML = '';
  const tr = document.createElement('tr');

  const thLabel = document.createElement('th');
  thLabel.className = 'row-label';
  thLabel.innerText = 'Rows';
  tr.appendChild(thLabel);

  headers.forEach(h => {
    const th = document.createElement('th');
    th.dataset.hid = h.id;
    th.style.padding = '6px 8px';
    th.innerHTML = `<div style="display:flex;flex-direction:column;gap:2px"><div style="font-weight:600;font-size:13px">${esc(h.title)}</div></div>`;
    th.addEventListener('contextmenu', async (ev) => {
      ev.preventDefault();
      showHeaderContextMenu(ev.pageX, ev.pageY, h);
    });
    tr.appendChild(th);
  });

  // Inline add-column cell (at end)
  const thAdd = document.createElement('th');
  thAdd.style.padding = '6px 8px';
  const wrapper = document.createElement('div');
  wrapper.style.display = 'flex';
  wrapper.style.gap = '6px';
  wrapper.style.alignItems = 'center';

  const select = document.createElement('select');
  select.title = 'Column type';
  select.style.padding = '6px';
  select.style.borderRadius = '6px';
  select.style.border = '1px solid var(--cell-border)';
  select.style.background = 'white';
  select.id = 'inlineNewColumnType';
  select.innerHTML = `
    <option value="string">String</option>
    <option value="int">Integer</option>
    <option value="datetime">Datetime</option>
    <option value="dropdown">Dropdown</option>
  `;

  const btn = document.createElement('button');
  btn.title = 'Add column';
  btn.style.padding = '6px 8px';
  btn.style.borderRadius = '6px';
  btn.style.border = '1px solid var(--border)';
  btn.style.background = 'white';
  btn.style.cursor = 'pointer';
  btn.textContent = '＋';
  btn.id = 'inlineAddColumnBtn';

  btn.addEventListener('click', async (ev) => {
    ev.stopPropagation();
    if (!activeSheetId) { toast('No active sheet'); return; }
    const type = (select && select.value) ? select.value : 'string';
    let title = prompt('Column title', 'Column');
    if (title == null) return;
    title = title.trim() || 'Column';
    try {
      const hRef = push(ref(db, sheetHeadersPath(activeSheetId)));
      const payload = { title, order: headers.length, type };
      if (type === 'dropdown') payload.options = ['Option 1','Option 2'];
      const updates = {};
      updates[`${sheetHeadersPath(activeSheetId)}/${hRef.key}`] = payload;
      await update(ref(db), updates);
      toast('Column added');
      setTimeout(()=> subscribeToActiveSheet(), 200);
    } catch (err) {
      console.error('Inline add column error', err);
      toast('Add column failed (see console)');
    }
  });

  wrapper.appendChild(select);
  wrapper.appendChild(btn);
  thAdd.appendChild(wrapper);
  tr.appendChild(thAdd);

  thead.appendChild(tr);
}

// ---------- header context menu ----------
function showHeaderContextMenu(x, y, header) {
  const prev = document.getElementById('hdr-context-menu'); if (prev) prev.remove();
  const menu = document.createElement('div');
  menu.id = 'hdr-context-menu';
  menu.style.position = 'absolute';
  menu.style.left = x + 'px';
  menu.style.top = y + 'px';
  menu.style.background = 'white';
  menu.style.border = '1px solid var(--border)';
  menu.style.boxShadow = '0 6px 18px rgba(2,6,23,0.08)';
  menu.style.borderRadius = '6px';
  menu.style.padding = '6px';
  menu.style.zIndex = 9999;
  menu.style.minWidth = '180px';
  menu.innerHTML = `
    <div style="padding:6px 8px;cursor:pointer" data-action="rename">Rename column</div>
    <div style="padding:6px 8px;cursor:pointer" data-action="type">Change type</div>
    <div style="padding:6px 8px;cursor:pointer" data-action="options">Edit dropdown options</div>
    <div style="padding:6px 8px;color:#c11;cursor:pointer" data-action="delete">Delete column</div>
  `;
  document.body.appendChild(menu);

  menu.querySelectorAll('div[data-action]').forEach(el=>{
    el.addEventListener('click', async ()=>{
      const action = el.dataset.action;
      menu.remove();
      if (action === 'rename') {
        const newTitle = prompt('New column title', header.title || '');
        if (newTitle == null) return;
        await update(ref(db), { [`${sheetHeadersPath(activeSheetId)}/${header.id}/title`]: newTitle });
        toast('Renamed');
      } else if (action === 'type') {
        const newType = prompt('Type (string / int / datetime /dropdown)', header.type || 'string');
        if (newType == null) return;
        const finalType = ['string','int','datetime','dropdown'].includes(newType) ? newType : (header.type||'string');
        const updates = {};
        updates[`${sheetHeadersPath(activeSheetId)}/${header.id}/type`] = finalType;
        if (finalType === 'dropdown') updates[`${sheetHeadersPath(activeSheetId)}/${header.id}/options`] = ['Option 1','Option 2'];
        await update(ref(db), updates);
        toast('Type updated');
      } else if (action === 'options') {
        const optsCsv = prompt('Comma-separated options', (header.options||[]).join(','));
        if (optsCsv == null) return;
        const opts = optsCsv.split(',').map(s=>s.trim()).filter(Boolean);
        await update(ref(db), { [`${sheetHeadersPath(activeSheetId)}/${header.id}/options`]: opts });
        toast('Options updated');
      } else if (action === 'delete') {
        if (!confirm('Delete column and its values?')) return;
        const updates = {};
        updates[`${sheetHeadersPath(activeSheetId)}/${header.id}`] = null;
        const rowsSnap = await get(ref(db, sheetRowsPath(activeSheetId)));
        const rows = rowsSnap.val() || {};
        for (const rk of Object.keys(rows)) updates[`${sheetRowsPath(activeSheetId)}/${rk}/cells/${header.id}`] = null;
        await update(ref(db), updates);
        toast('Deleted');
      }
    });
  });

  const closer = (ev) => { if (!menu.contains(ev.target)) menu.remove(); };
  setTimeout(()=> window.addEventListener('mousedown', closer, { once:true }), 0);
}

// ---------- row DOM builder ----------
function makeTr(rowKey, rowData, rowIndex) {
  const tr = document.createElement('tr');
  tr.dataset.key = rowKey || `local-${Date.now()}-${Math.floor(Math.random()*1000)}`;
  tr._suppress = false;

  const tdLabel = document.createElement('td');
  tdLabel.className = 'row-index';
  tdLabel.innerHTML = `<div style="font-size:12px;color:var(--muted)">${rowIndex || ''}</div>`;
  tr.appendChild(tdLabel);

  headers.forEach(h => {
    const td = document.createElement('td');
    td.dataset.hid = h.id;
    const value = rowData?.cells?.[h.id] ?? '';
    const type = h.type || 'string';
    const opts = h.options || [];
    let inner = '';

    if (type === 'string') inner = `<input type="text" data-field="${h.id}" value="${esc(value)}" />`;
    else if (type === 'int') inner = `<input type="number" step="1" data-field="${h.id}" value="${esc(value)}" />`;
    else if (type === 'datetime') {
      let localVal = '';
      if (value) {
        try {
          const dt = new Date(value);
          const pad = n => String(n).padStart(2,'0');
          localVal = `${dt.getFullYear()}-${pad(dt.getMonth()+1)}-${pad(dt.getDate())}T${pad(dt.getHours())}:${pad(dt.getMinutes())}`;
        } catch(e){}
      }
      inner = `<input type="datetime-local" data-field="${h.id}" value="${esc(localVal)}" />`;
    } else if (type === 'dropdown') {
      if (opts && opts.length) {
        inner = `<select data-field="${h.id}">
      <option value=""></option>
      ${opts.map(opt => `<option value="${esc(opt)}"${opt === value ? ' selected' : ''}>${esc(opt)}</option>`).join('')}
    </select>`;
      } else {
        inner = `<select data-field="${h.id}"><option value="">No options — set options</option></select>
                 <button type="button" class="set-opts-btn" data-hid="${h.id}" title="Set dropdown options" style="margin-left:6px;padding:4px;border-radius:4px;border:1px solid var(--border);background:white;cursor:pointer">⚙</button>`;
      }
    } else inner = `<input type="text" data-field="${h.id}" value="${esc(value)}" />`;

    td.innerHTML = inner;
    tr.appendChild(td);
  });

  const tdActions = document.createElement('td');
  tdActions.className = 'row-actions';
  tdActions.innerHTML = `
    <div style="display:flex;flex-direction:column;gap:6px;align-items:center">
      <div style="display:flex;gap:6px">
        <button class="small-icon-btn save-row" title="Save">💾</button>
        <button class="small-icon-btn cancel-row" title="Revert">↺</button>
        <button class="small-icon-btn delete-row" title="Delete">🗑</button>
      </div>
    </div>
  `;
  tr.appendChild(tdActions);

  if (String(tr.dataset.key).startsWith('local-')) tr.classList.add('local');

  attachRowHandlers(tr);
  return tr;
}

// ---------- attach handlers ----------
function attachRowHandlers(tr) {
  if (tr._handlersAttached) return;
  tr._handlersAttached = true;

  function markDirty(rowEl) { rowEl.dataset.dirty = '1'; rowEl.classList.add('unsaved'); }
  function clearDirty(rowEl) { delete rowEl.dataset.dirty; rowEl.classList.remove('unsaved'); }

  function collectAndCoerceCellsFromRowEl(rowEl) {
    const cells = {};
    rowEl.querySelectorAll('input[data-field], select[data-field]').forEach(el => {
      const hid = el.dataset.field;
      const header = headers.find(h => h.id === hid) || { type: 'string' };
      const raw = el.value;
      if (header.type === 'int') {
        const parsed = parseInt(raw, 10);
        cells[hid] = Number.isFinite(parsed) ? parsed : null;
      } else if (header.type === 'datetime') {
        if (raw) { try { cells[hid] = new Date(raw).toISOString(); } catch(e){ cells[hid] = ''; } }
        else cells[hid] = '';
      } else {
        cells[hid] = raw;
      }
    });
    return cells;
  }

  async function saveRow(rowEl) {
    if (!rowEl) return false;
    if (!rowEl.dataset.dirty) return false;
    if (rowEl._saving) { dbg('saveRow: already saving, skipping'); return false; }
    rowEl._saving = true;

    try {
      let rawKey = rowEl.dataset.key;
      const cells = collectAndCoerceCellsFromRowEl(rowEl);

      // If placeholder and all cells empty => remove and skip pushing
      const hasValue = Object.values(cells).some(v => v !== '' && v !== null && v !== undefined);
      if ((!rawKey || rawKey.startsWith('local-')) && !hasValue) {
        dbg('saveRow: empty local placeholder — removing and skipping save');
        const oldKey = rowEl.dataset.key;
        rowEl.remove();
        for (const [k, v] of Array.from(rowsMap.entries())) { if (v.tr === rowEl) rowsMap.delete(k); }
        return false;
      }

      const payload = { cells, updatedAt: new Date().toISOString(), owner: auth.currentUser?.uid || null };

      if (!rawKey || rawKey.startsWith('local-')) {
        rawKey = rowEl.dataset.key;
        if (!rawKey || rawKey.startsWith('local-')) {
          const newRef = push(ref(db, sheetRowsPath(activeSheetId)));
          await set(newRef, payload);
          dbg('saveRow: pushed new row', newRef.key);
          const oldKey = rowEl.dataset.key;
          rowEl.dataset.key = newRef.key;
          rowEl.classList.remove('local');
          for (const [k, v] of Array.from(rowsMap.entries())) {
            if (v.tr === rowEl) { rowsMap.delete(k); rowsMap.set(newRef.key, { tr: rowEl }); break; }
          }
        } else {
          await set(ref(db, `${sheetRowsPath(activeSheetId)}/${rawKey}`), payload);
        }
      } else {
        await set(ref(db, `${sheetRowsPath(activeSheetId)}/${rawKey}`), payload);
      }

      delete rowEl.dataset.dirty;
      rowEl.classList.remove('unsaved');
      return true;
    } catch (err) {
      console.error('Save row error', err);
      return false;
    } finally {
      rowEl._saving = false;
    }
  }

  async function cancelRow(rowEl) {
    const rawKey = rowEl.dataset.key;
    if (!rawKey || rawKey.startsWith('local-')) {
      rowEl.remove();
      for (const [k, v] of rowsMap.entries()) { if (v.tr === rowEl) rowsMap.delete(k); }
      return;
    }
    try {
      const snap = await get(ref(db, `${sheetRowsPath(activeSheetId)}/${rawKey}`));
      const remote = snap.val() || { cells: {} };
      rowEl._suppress = true;
      rowEl.querySelectorAll('input[data-field], select[data-field]').forEach(inp => {
        const hid = inp.dataset.field;
        const val = remote?.cells?.[hid];
        const header = headers.find(h => h.id === hid) || { type: 'string' };
        if (header.type === 'datetime' && typeof val === 'string' && val) {
          try {
            const dt = new Date(val);
            const pad = n => String(n).padStart(2,'0');
            inp.value = `${dt.getFullYear()}-${pad(dt.getMonth()+1)}-${pad(dt.getDate())}T${pad(dt.getHours())}:${pad(dt.getMinutes())}`;
          } catch (e) { inp.value = ''; }
        } else {
          inp.value = (val !== undefined && val !== null) ? String(val) : '';
        }
      });
      setTimeout(()=> rowEl._suppress = false, 50);
      clearDirty(rowEl);
    } catch (err) { console.error('Cancel reload error', err); }
  }

  async function deleteRow(rowEl) {
    const rawKey = rowEl.dataset.key;
    if (!rawKey || rawKey.startsWith('local-')) {
      rowEl.remove();
      for (const [k, v] of rowsMap.entries()) { if (v.tr === rowEl) rowsMap.delete(k); }
      return;
    }
    if (!confirm('Delete this row?')) return;
    try { await remove(ref(db, `${sheetRowsPath(activeSheetId)}/${rawKey}`)); } catch (err) { console.error('Delete row error', err); }
  }

  // wire buttons
  const saveBtn = tr.querySelector('.save-row');
  if (saveBtn) saveBtn.addEventListener('click', async ()=> { if (!tr._saving) { saveBtn.disabled = true; try { await saveRow(tr); } finally { saveBtn.disabled = false; } } });
  const cancelBtn = tr.querySelector('.cancel-row'); if (cancelBtn) cancelBtn.addEventListener('click', async ()=> { await cancelRow(tr); });
  const delBtn = tr.querySelector('.delete-row'); if (delBtn) delBtn.addEventListener('click', async ()=> { await deleteRow(tr); });

  tr.querySelectorAll('input[data-field], select[data-field]').forEach(inp => {
    inp.addEventListener('input', ()=> {
      if (tr._suppress) return;
      markDirty(tr);
      debouncedPersistLocal();
    });

    inp.addEventListener('focus', async (e) => {
      try {
        const hid = inp.dataset.field;
        const header = headers.find(h=>h.id===hid);
        if (header && header.type === 'dropdown' && (!header.options || header.options.length === 0)) {
          const resp = prompt('This dropdown has no options yet. Enter comma-separated options (or Cancel):', '');
          if (resp == null) return;
          const opts = resp.split(',').map(s=>s.trim()).filter(Boolean);
          if (!opts.length) return;
          await update(ref(db), { [`${sheetHeadersPath(activeSheetId)}/${hid}/options`]: opts });
          toast('Options saved');
          header.options = opts;
          const sel = inp;
          sel.innerHTML = opts.map(opt=>`<option value="${esc(opt)}">${esc(opt)}</option>`).join('');
          markDirty(tr);
        }
      } catch (err) {
        console.error('Setting options on focus failed', err);
      }
    });

    inp.addEventListener('keydown', (e)=> {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); saveRow(tr); }
    });
  });

  tr.querySelectorAll('.set-opts-btn').forEach(btn => {
    btn.addEventListener('click', async (ev) => {
      ev.preventDefault();
      const hid = btn.dataset.hid;
      const header = headers.find(h=>h.id===hid);
      if (!header) { toast('Header not found'); return; }
      const optsCsv = prompt('Enter comma-separated options for this column', (header.options||[]).join(','));
      if (optsCsv == null) return;
      const opts = optsCsv.split(',').map(s=>s.trim()).filter(Boolean);
      try {
        await update(ref(db), { [`${sheetHeadersPath(activeSheetId)}/${hid}/options`]: opts });
        toast('Options updated');
        header.options = opts;
        const sel = tr.querySelector(`select[data-field="${hid}"]`);
        if (sel) sel.innerHTML = opts.map(opt=>`<option value="${esc(opt)}">${esc(opt)}</option>`).join('');
        markDirty(tr);
      } catch (err) {
        console.error('Set options error', err);
        toast('Failed to save options (see console)');
      }
    });
  });

  tr._saveRow = async () => await saveRow(tr);

  function markDirty(rowEl) { rowEl.dataset.dirty = '1'; rowEl.classList.add('unsaved'); }
  function clearDirty(rowEl) { delete rowEl.dataset.dirty; rowEl.classList.remove('unsaved'); }
}

// ---------- create local placeholder row ----------
function createLocalRow(rowIndex = null) {
  const tempKey = 'local-' + Date.now() + '-' + Math.floor(Math.random()*1000);
  const placeholder = { cells: {} };
  if (headers && headers.length) { headers.forEach(h => placeholder.cells[h.id] = ''); }
  const tr = makeTr(tempKey, placeholder, rowIndex);
  if (tbody) {
    const bottom = tbody.querySelector('tr.bottom-add-row');
    if (bottom) tbody.insertBefore(tr, bottom);
    else tbody.appendChild(tr);
  }
  rowsMap.set(tempKey, { tr });
  setTimeout(()=> {
    const t = rowsMap.get(tempKey).tr;
    if (!t) return;
    const first = t.querySelector('input[data-field], select[data-field]');
    if (first) first.focus();
  }, 10);
  return tr;
}

// ---------- upsert row ----------
function upsertRow(key, rowData, index){
  if (rowsMap.has(key)){
    const { tr } = rowsMap.get(key);
    tr._suppress = true;
    tr.dataset.key = key;
    tr.querySelectorAll('input[data-field], select[data-field]').forEach(inp => {
      const hid = inp.dataset.field;
      const header = headers.find(h => h.id === hid) || { type: 'string' };
      const val = rowData?.cells?.[hid];
      if (header.type === 'datetime' && typeof val === 'string' && val) {
        try {
          const dt = new Date(val); const pad = n => String(n).padStart(2,'0');
          inp.value = `${dt.getFullYear()}-${pad(dt.getMonth()+1)}-${pad(dt.getDate())}T${pad(dt.getHours())}:${pad(dt.getMinutes())}`;
        } catch(e){ inp.value = ''; }
      } else {
        inp.value = (val !== undefined && val !== null) ? String(val) : '';
      }
    });
    setTimeout(()=> tr._suppress = false, 50);
    return tr;
  }
  const tr = makeTr(key, rowData, index);
  tbody.appendChild(tr);
  rowsMap.set(key, { tr });
  return tr;
}

// ---------- render bottom add-row control ----------
function renderBottomAddRow() {
  try {
    if (!tbody) { console.warn('renderBottomAddRow: tbody not found'); return; }
    const existing = tbody.querySelector('tr.bottom-add-row');
    if (existing) existing.remove();

    const tr = document.createElement('tr');
    tr.className = 'bottom-add-row';
    const tdLeft = document.createElement('td');
    tdLeft.className = 'row-index';
    tdLeft.style.padding = '12px';
    tdLeft.style.verticalAlign = 'middle';
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = '＋';
    btn.title = 'Add row';
    btn.style.padding = '8px';
    btn.style.borderRadius = '8px';
    btn.style.border = '1px solid var(--border)';
    btn.style.background = 'white';
    btn.style.cursor = 'pointer';
    btn.style.minWidth = '36px';
    btn.style.minHeight = '36px';
    btn.addEventListener('click', () => {
      createLocalRow((tbody.querySelectorAll('tr:not(.bottom-add-row)')?.length || 0) + 1);
      const lastRow = tbody.querySelector('tr:not(.bottom-add-row):last-child');
      if (lastRow) {
        const first = lastRow.querySelector('input[data-field], select[data-field]');
        if (first) first.focus();
        lastRow.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    });
    tdLeft.appendChild(btn);
    tr.appendChild(tdLeft);
    const tdRight = document.createElement('td');
    tdRight.colSpan = Math.max((headers?.length || 0) + 1, 3);
    tdRight.style.padding = '8px';
    tr.appendChild(tdRight);
    tbody.appendChild(tr);
    dbg('renderBottomAddRow appended');
  } catch (err) { console.error('renderBottomAddRow error', err); }
}

// ---------- render body ----------
function renderBody(rowsArray){
  tbody.innerHTML = '';
  rowsMap.clear();

  if (!headers || headers.length === 0) {
    const tr = document.createElement('tr');
    const tdLabel = document.createElement('td');
    tdLabel.className = 'row-index';
    tdLabel.innerHTML = `<div style="font-size:12px;color:var(--muted)">1</div>`;
    tr.appendChild(tdLabel);
    const td = document.createElement('td');
    td.colSpan = 6;
    td.style.padding = '18px';
    td.style.color = 'var(--muted)';
    td.innerHTML = `<div style="opacity:0.9">No columns yet — click the <strong>＋</strong> button in the header to add a column, or use the top controls.</div>`;
    tr.appendChild(td);
    tbody.appendChild(tr);
    renderBottomAddRow();
    return;
  }

  // render DB rows first (ascending updatedAt => newest at bottom)
  rowsArray.forEach((item, idx) => upsertRow(item.key, item.value, idx+1));

  // NOTE: localStorage restoration REMOVED to avoid phantom duplicate rows

  // ensure there is at least one local placeholder
  const anyLocal = tbody.querySelector('tr.local');
  if (!anyLocal) createLocalRow(rowsArray.length + 1);

  // finally add the bottom "+" row
  renderBottomAddRow();
}

// ---------- UI actions wiring ----------
// Add Tab
if (addTabBtn) {
  addTabBtn.addEventListener('click', async ()=>{
    dbg('Add Tab clicked');
    const title = prompt('New sheet name', 'Sheet');
    if (title == null) { dbg('Add Tab cancelled'); return; }
    try {
      const newRef = push(ref(db, SHEETS_ROOT));
      const sid = newRef.key;
      const updates = {};
      updates[`${SHEETS_ROOT}/${sid}/meta/title`] = title;
      const h1 = push(ref(db, `${SHEETS_ROOT}/${sid}/headers`));
      const h2 = push(ref(db, `${SHEETS_ROOT}/${sid}/headers`));
      updates[`${SHEETS_ROOT}/${sid}/headers/${h1.key}/title`] = 'Column 1';
      updates[`${SHEETS_ROOT}/${sid}/headers/${h1.key}/order`] = 0;
      updates[`${SHEETS_ROOT}/${sid}/headers/${h1.key}/type`] = 'string';
      updates[`${SHEETS_ROOT}/${sid}/headers/${h2.key}/title`] = 'Column 2';
      updates[`${SHEETS_ROOT}/${sid}/headers/${h2.key}/order`] = 1;
      updates[`${SHEETS_ROOT}/${sid}/headers/${h2.key}/type`] = 'string';
      const r1 = push(ref(db, `${SHEETS_ROOT}/${sid}/rows`));
      const r2 = push(ref(db, `${SHEETS_ROOT}/${sid}/rows`));
      updates[`${SHEETS_ROOT}/${sid}/rows/${r1.key}/cells/${h1.key}`] = '';
      updates[`${SHEETS_ROOT}/${sid}/rows/${r1.key}/cells/${h2.key}`] = '';
      updates[`${SHEETS_ROOT}/${sid}/rows/${r1.key}/updatedAt`] = new Date().toISOString();
      updates[`${SHEETS_ROOT}/${sid}/rows/${r2.key}/cells/${h1.key}`] = '';
      updates[`${SHEETS_ROOT}/${sid}/rows/${r2.key}/cells/${h2.key}`] = '';
      updates[`${SHEETS_ROOT}/${sid}/rows/${r2.key}/updatedAt`] = new Date().toISOString();
      await update(ref(db), updates);
      dbg('Sheet created', sid);
      selectSheet(sid);
      toast('Sheet created');
    } catch (err){
      console.error('add tab error', err);
      toast('Failed to create sheet (see console)');
    }
  });
} else dbg('addTabBtn not found');

// Add Row (top control)
if (addRowBtn) {
  addRowBtn.addEventListener('click', async ()=>{
    dbg('Add Row clicked, activeSheetId=', activeSheetId);
    if (!activeSheetId){ toast('No active sheet'); return; }
    try {
      const rRef = push(ref(db, sheetRowsPath(activeSheetId)));
      const cells = {};
      headers.forEach(h => cells[h.id] = '');
      const payload = { cells, updatedAt: new Date().toISOString(), owner: auth.currentUser?.uid || null };
      await set(rRef, payload);
      dbg('row created', rRef.key);
      toast('Row added');
    } catch (err) {
      console.error('add row error', err);
      toast('Failed to add row (see console)');
    }
  });
} else dbg('addRowBtn not found');

// SINGLE add-column handler — safe: writes header only
if (addColumnBtn) {
  try { addColumnBtn.replaceWith(addColumnBtn.cloneNode(true)); } catch(e){ dbg('replace clone failed', e); }
  const btn = document.getElementById('addColumn');
  btn.addEventListener('click', async (ev) => {
    ev.stopPropagation();
    if (!activeSheetId) { toast('No active sheet'); return; }
    const sel = document.getElementById('newColumnType');
    const type = (sel && sel.value) ? sel.value : 'string';
    let title = prompt('Column title', 'Column');
    if (title == null) return;
    title = title.trim() || 'Column';
    try {
      const hRef = push(ref(db, sheetHeadersPath(activeSheetId)));
      const payload = { title, order: headers.length, type };
      if (type === 'dropdown') payload.options = ['Option 1','Option 2'];
      const updates = {};
      updates[`${sheetHeadersPath(activeSheetId)}/${hRef.key}`] = payload;
      await update(ref(db), updates);
      setTimeout(()=> subscribeToActiveSheet(), 200);
      toast('Column added');
    } catch (err) {
      console.error('Add column error', err);
      toast('Add column failed (see console)');
    }
  });
} else dbg('addColumnBtn not found');

// Clear table
if (clearTableBtn) {
  clearTableBtn.addEventListener('click', async ()=>{
    if (!activeSheetId){ toast('No active sheet'); return; }
    if (!confirm('Clear headers and rows for this sheet?')) return;
    try {
      const updates = {};
      updates[`${SHEETS_ROOT}/${activeSheetId}/headers`] = null;
      updates[`${SHEETS_ROOT}/${activeSheetId}/rows`] = null;
      await update(ref(db), updates);
      toast('Cleared sheet');
    } catch (err) { console.error(err); toast('Clear failed'); }
  });
}

// Export CSV
if (exportCsvBtn) {
  exportCsvBtn.addEventListener('click', ()=>{
    if (!activeSheetId){ toast('No active sheet'); return; }
    const headerTitles = headers.map(h=>h.title);
    const rows = [];
    tbody.querySelectorAll('tr').forEach(tr=>{
      const key = tr.dataset.key || '';
      const cells = headers.map(h=>{
        const el = tr.querySelector(`[data-field="${h.id}"]`);
        const v = el ? el.value : '';
        return (v+'').replaceAll('"','""');
      });
      rows.push(`"${(key||'').replaceAll('"','""')}",${cells.map(v=>`"${v}"`).join(',')}`);
    });
    const csv = [`"RowID",${headerTitles.map(t=>`"${(t||'').replaceAll('"','""')}"`).join(',')}`].concat(rows).join('\n');
    const blob = new Blob([csv], { type:'text/csv' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${(pageTitle.textContent||'sheet').replaceAll(' ','_')}.csv`; document.body.appendChild(a); a.click(); a.remove();
    toast('CSV exported');
  });
}

// selectSheet (save dirty rows first)
async function selectSheet(sid) {
  if (!sid) return;
  if (activeSheetId && activeSheetId !== sid) {
    try { await saveAllDirtyRows(); } catch (err) { console.warn('Error saving before sheet switch', err); }
  }
  activeSheetId = sid;
  const s = sheets.find(x => x.id === sid);
  pageTitle.textContent = s?.title || 'Sheet';
  renderSheetsList();
  subscribeToActiveSheet();
}

// auth + start
async function start(){
  try {
    dbg('Signing in anonymously...');
    await signInAnonymously(auth);
  } catch (err) {
    console.error('signInAnonymously failed', err);
    toast('Auth failed (see console)');
    return;
  }
  onAuthStateChanged(auth, user => {
    if (user) { dbg('Signed in uid=', user.uid); subscribeSheetsList(); }
    else dbg('Auth state: signed out');
  });
}
start();
