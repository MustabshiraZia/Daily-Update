/* People page table with add-row, editable cells, broadcast sync and localStorage persistence */

const CHANNEL = "daily-update-people-v1";
const channel = new BroadcastChannel(CHANNEL);
const stateKey = "daily-update-people-sheet";
const COLS = ["Date","Atika","Ben","Bert","Dave","Isabelle","Louis","Mustabshira","Stephen","Note Any Upcoming Availability"];

const sampleRows = [
  ["6/12/23","","","Happy","","","Very Happy","Alwyn to block days out for diff. centre. Weds&Thurs for WACC."],
  ["19/06/2023","","","Very Happy","","","Okay",""],
  ["22/6/23","","","Very Happy","","","Very Happy",""],
  ["26/6/23","","","Happy","","","Happy",""],
  ["29/6/2023","","","","","","Very Happy",""]
];

// ensure consistent number of columns (10 data cells after Date)
function normalizeRow(arr){
  const out = new Array(COLS.length).fill("");
  for(let i=0;i<Math.min(arr.length, COLS.length); i++) out[i]=arr[i]??"";
  return out;
}

// DOM references
const peopleHead = document.getElementById("peopleHead");
const peopleBody = document.getElementById("peopleBody");
const addRowBtn = document.getElementById("addRow");
const clearTableBtn = document.getElementById("clearTable");
const exportBtn = document.getElementById("exportCsv");
const toast = document.getElementById("toast");
const pageTitle = document.getElementById("pageTitle");

// app state
let rows = []; // array of arrays
let editingCell = null;

function showToast(msg, ms=1200){
  toast.textContent = msg;
  toast.classList.remove("hidden");
  clearTimeout(toast._t);
  toast._t = setTimeout(()=> toast.classList.add("hidden"), ms);
}

function save(){
  localStorage.setItem(stateKey, JSON.stringify(rows));
  channel.postMessage({type:"sync", rows});
}

function load(){
  const raw = localStorage.getItem(stateKey);
  if(raw){
    try{ rows = JSON.parse(raw); return; } catch(e){}
  }
  // fallback sample (normalize)
  rows = sampleRows.map(r => normalizeRow(r));
  save();
}

function buildHead(){
  const tr = document.createElement("tr");
  const thLeft = document.createElement("th");
  thLeft.className = "row-label";
  thLeft.textContent = "";
  tr.appendChild(thLeft);

  COLS.forEach(c=>{
    const th = document.createElement("th");
    th.textContent = c;
    tr.appendChild(th);
  });
  peopleHead.innerHTML = "";
  peopleHead.appendChild(tr);
}

function buildBody(){
  peopleBody.innerHTML = "";
  rows.forEach((r, rowIndex) => {
    const tr = document.createElement("tr");
    const idx = document.createElement("td");
    idx.className = "row-index";
    idx.textContent = rowIndex+1;
    tr.appendChild(idx);

    for(let c=0;c<COLS.length;c++){
      const td = document.createElement("td");
      td.dataset.r = rowIndex;
      td.dataset.c = c;
      td.tabIndex = 0;
      td.textContent = r[c] ?? "";
      td.addEventListener("dblclick", enterEdit);
      td.addEventListener("keydown", (e)=> {
        if(e.key === "Enter") { enterEdit.call(td, e); e.preventDefault(); }
      });
      tr.appendChild(td);
    }

    // delete button cell appended (rightmost)
    const del = document.createElement("td");
    del.innerHTML = `<button class="small" data-r="${rowIndex}">Delete</button>`;
    del.style.minWidth = "90px";
    del.querySelector("button").addEventListener("click", (e)=>{
      const idx = Number(e.currentTarget.dataset.r);
      if(!confirm(`Delete row ${idx+1}?`)) return;
      rows.splice(idx,1);
      save();
      render();
      showToast("Row deleted");
    });
    tr.appendChild(del);

    peopleBody.appendChild(tr);
  });
  // add an empty footer row showing add-button hint
  const footer = document.createElement("tr");
  const ftd = document.createElement("td");
  ftd.colSpan = COLS.length + 2;
  ftd.style.padding = "10px";
  ftd.style.background = "#fbfdff";
  ftd.textContent = "Double-click any cell to edit. Use + Add row to append a new row.";
  footer.appendChild(ftd);
  peopleBody.appendChild(footer);
}

function enterEdit(e){
  const td = (this && this.tagName==="TD") ? this : e.currentTarget;
  const r = Number(td.dataset.r);
  const c = Number(td.dataset.c);
  if(Number.isNaN(r) || Number.isNaN(c)) return;
  if(editingCell) return;
  editingCell = {r,c,td};
  const input = document.createElement("input");
  input.type = "text";
  input.value = rows[r][c] ?? "";
  input.style.width = "100%";
  td.classList.add("editing");
  td.innerHTML = "";
  td.appendChild(input);
  input.focus();
  input.select();

  function finish(saveVal){
    td.classList.remove("editing");
    if(saveVal !== undefined){
      rows[r][c] = saveVal;
      save();
      showToast("Saved");
    }
    editingCell = null;
    render();
  }

  input.addEventListener("blur", ()=> finish(input.value));
  input.addEventListener("keydown", ev=>{
    if(ev.key === "Enter"){ ev.preventDefault(); input.blur(); }
    if(ev.key === "Escape"){ editingCell = null; render(); }
  });
}

function addRow(newRow){
  rows.push(newRow || new Array(COLS.length).fill(""));
  save();
  render();
  showToast("Row added");
}

function clearTable(){
  if(!confirm("Clear all rows? This cannot be undone.")) return;
  rows = [];
  save();
  render();
  showToast("Cleared");
}

function exportCsv(){
  const header = COLS.slice();
  const rowsOut = [header];
  rows.forEach(r=>{
    rowsOut.push(r.map(v => (v||"").toString()));
  });
  const csv = rowsOut.map(r => r.map(x=> `"${(x||"").replace(/"/g,'""')}"`).join(",")).join("\n");
  const blob = new Blob([csv], {type:"text/csv"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "people.csv";
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
  showToast("CSV exported");
}

function render(){
  buildHead();
  buildBody();
  pageTitle.textContent = "People";
}

// BroadcastChannel handler
channel.onmessage = (ev) => {
  const msg = ev.data;
  if(!msg) return;
  if(msg.type === "sync"){
    // naive approach: accept remote rows
    rows = msg.rows || [];
    localStorage.setItem(stateKey, JSON.stringify(rows));
    render();
    showToast("Synced from collaborator", 900);
  }
};

// UI wiring
addRowBtn.addEventListener("click", ()=> addRow());
clearTableBtn.addEventListener("click", ()=> clearTable());
exportBtn.addEventListener("click", ()=> exportCsv());

// menu wiring: page switching
document.querySelectorAll(".menu-btn").forEach(b=>{
  b.addEventListener("click", ()=>{
    document.querySelectorAll(".menu-btn").forEach(x=>x.classList.remove("active"));
    b.classList.add("active");
    const page = b.dataset.page;
    document.querySelectorAll(".page").forEach(p=>p.classList.add("hidden"));
    document.getElementById(page + "Page").classList.remove("hidden");
    document.getElementById("pageTitle").textContent = b.textContent;
  });
});

// initial load
load();
render();

// post initial sync so new tabs get current content
channel.postMessage({type:"sync", rows});
