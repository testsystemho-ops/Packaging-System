/* ===================== js/utils.js ===================== */
// ============================================================
// utils.js — ฟังก์ชันช่วยทั่วไป: Toast, Modal, วันที่, Firebase, Excel
// ============================================================

/* ---------- Date helpers ---------- */
function pad2(n){ return String(n).padStart(2,'0'); }

function todayStr(){
  const d = new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
}

function thaiDate(dateStr){
  // 'YYYY-MM-DD' -> 'DD/MM/YYYY'
  if(!dateStr) return '-';
  const [y,m,d] = dateStr.split('-');
  return `${d}/${m}/${y}`;
}

function daysInMonth(yyyyMm){
  // 'YYYY-MM' -> number of days
  const [y,m] = yyyyMm.split('-').map(Number);
  return new Date(y, m, 0).getDate();
}

function dateRange(from, to){
  const out = [];
  let cur = new Date(from + 'T00:00:00');
  const end = new Date(to + 'T00:00:00');
  if(isNaN(cur) || isNaN(end) || cur > end) return out;
  while(cur <= end){
    out.push(`${cur.getFullYear()}-${pad2(cur.getMonth()+1)}-${pad2(cur.getDate())}`);
    cur.setDate(cur.getDate()+1);
  }
  return out;
}

function fmtDateTime(ts){
  if(!ts) return '-';
  const d = new Date(ts);
  return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
}

/* ---------- Number / formatting helpers ---------- */
function fmtNum(n, digits=0){
  n = Number(n)||0;
  return n.toLocaleString('th-TH', {minimumFractionDigits:digits, maximumFractionDigits:digits});
}
function fmtMoney(n){ return fmtNum(n, 2); }

function escapeHtml(str){
  if(str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

/* ---------- Toast ---------- */
let toastTimer = null;
function toast(msg, type='default'){
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'show' + (type!=='default' ? ' '+type : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(()=>{ t.className = ''; }, 3200);
}

/* ---------- Modal ---------- */
function showModal(innerHtml, onMount){
  const root = document.getElementById('modalRoot');
  root.innerHTML = `<div class="modal-backdrop" id="modalBackdrop"><div class="modal">${innerHtml}</div></div>`;
  document.getElementById('modalBackdrop').addEventListener('click', (e)=>{
    if(e.target.id === 'modalBackdrop') closeModal();
  });
  if(onMount) onMount(root);
}
function closeModal(){
  document.getElementById('modalRoot').innerHTML = '';
}

/* ---------- Loading state on buttons ---------- */
function setLoading(btn, loading, loadingText='กำลังโหลด...'){
  if(!btn) return;
  if(loading){
    btn.dataset.origText = btn.innerHTML;
    btn.innerHTML = loadingText;
    btn.disabled = true;
  } else {
    if(btn.dataset.origText) btn.innerHTML = btn.dataset.origText;
    btn.disabled = false;
  }
}

/* ---------- Firebase helpers ---------- */
async function dbGetOnce(path){
  const snap = await db.ref(path).once('value');
  return snap.val();
}
async function dbUpdate(updates){
  return db.ref().update(updates);
}
async function dbSet(path, val){
  return db.ref(path).set(val);
}
async function dbPush(path, val){
  return db.ref(path).push(val);
}
async function dbRemove(path){
  return db.ref(path).remove();
}

/* ---------- Item helpers ---------- */
const CAT_LABELS = { FRESH: 'FRESH FOOD', TRANSFER: 'TRANSFER', NONFRESH: 'NON FRESH' };
const CAT_FIELD_LABELS = { FV: 'F&V', BUT: 'BUT', FISH: 'FISH', QTY: 'จำนวน' };

function itemUnitCost(item){
  const pack = Number(item.packCount) || 1;
  const price = Number(item.price) || 0;
  return price / pack;
}
function recordTotal(item, rec){
  if(!rec) return 0;
  return (item.subFields||[]).reduce((s,f)=> s + (Number(rec[f])||0), 0);
}
function recordAmount(item, rec){
  return recordTotal(item, rec) * itemUnitCost(item);
}

/* ---------- Excel export ---------- */
function exportRowsToExcel(sheets, filename){
  // sheets: { 'ชื่อชีท': [ {col:val,...}, ... ] }
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows])=>{
    const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ 'ไม่มีข้อมูล': '' }]);
    XLSX.utils.book_append_sheet(wb, ws, name.substring(0,31));
  });
  XLSX.writeFile(wb, filename);
}

/* ---------- Store helpers ---------- */
function getStoreByCode(code){
  return STORES_DATA.find(s=>s.username === code);
}
function storeLabel(code){
  const s = getStoreByCode(code);
  return s ? `${s.locNo} - ${s.name}` : code;
}

/* ---------- Export row builder (shared by store + admin export) ---------- */
function buildExportRow(dateStr, storeCode, cat, item, rec){
  const s = getStoreByCode(storeCode);
  const total = recordTotal(item, rec);
  const amount = recordAmount(item, rec);
  return {
    'วันที่': thaiDate(dateStr),
    'รหัสผู้ใช้สาขา': storeCode,
    'เลขที่สาขา (Loc)': s ? s.locNo : '',
    'ชื่อสาขา': s ? s.name : '',
    'หมวดหมู่': CAT_LABELS[cat] || cat,
    'รหัสสินค้า': item.code,
    'รายการสินค้า': item.desc,
    'ผู้ขาย/Supplier': item.supplier,
    'หน่วยนับ': item.uomCount,
    'F&V': cat === 'NONFRESH' ? '' : (Number(rec.FV)||0),
    'BUT': cat === 'NONFRESH' ? '' : (Number(rec.BUT)||0),
    'FISH': cat === 'NONFRESH' ? '' : (Number(rec.FISH)||0),
    'จำนวน (Non Fresh)': cat === 'NONFRESH' ? (Number(rec.QTY)||0) : '',
    'รวมจำนวนตรวจนับ': total,
    'ราคาต่อหน่วยนับ (บาท)': Math.round(itemUnitCost(item)*100)/100,
    'มูลค่ารวม (บาท)': Math.round(amount*100)/100
  };
}

/* ---------- Dashboard helpers: colors + donut chart ---------- */
const CAT_COLORS = {
  FRESH: '#1FA97C',     // เขียว
  TRANSFER: '#0B5FB4',  // น้ำเงิน (Makro)
  NONFRESH: '#F5A623'   // ส้ม (Makro)
};

/**
 * สร้าง SVG Donut Chart
 * segments: [{ label, value, color }]
 * คืนค่า HTML string ของ <svg>
 */
function buildDonutChart(segments, opts={}){
  const size = opts.size || 200;
  const stroke = opts.stroke || 28;
  const radius = (size - stroke) / 2;
  const circumference = 2 * Math.PI * radius;
  const total = segments.reduce((s,x)=> s + (Number(x.value)||0), 0);

  let offset = 0;
  const arcs = segments.map(seg=>{
    const value = Number(seg.value) || 0;
    const frac = total > 0 ? value / total : 0;
    const dash = frac * circumference;
    const gap = circumference - dash;
    const arc = `<circle cx="${size/2}" cy="${size/2}" r="${radius}" fill="none"
        stroke="${seg.color}" stroke-width="${stroke}"
        stroke-dasharray="${dash.toFixed(3)} ${gap.toFixed(3)}"
        stroke-dashoffset="${(-offset).toFixed(3)}"
        transform="rotate(-90 ${size/2} ${size/2})"
        stroke-linecap="${segments.length>1?'butt':'round'}"></circle>`;
    offset += dash;
    return arc;
  }).join('');

  const centerLabel = opts.centerLabel || 'มูลค่ารวม';
  const centerValue = opts.centerValue !== undefined ? opts.centerValue : fmtMoney(total);

  return `
  <svg viewBox="0 0 ${size} ${size}" width="${size}" height="${size}" style="display:block;">
    <circle cx="${size/2}" cy="${size/2}" r="${radius}" fill="none" stroke="var(--border-soft)" stroke-width="${stroke}"></circle>
    ${total > 0 ? arcs : ''}
    <text x="50%" y="46%" text-anchor="middle" dominant-baseline="middle"
      font-family="Inter,'Noto Sans Thai',sans-serif" font-size="12" fill="var(--text-soft)" font-weight="600">${escapeHtml(centerLabel)}</text>
    <text x="50%" y="60%" text-anchor="middle" dominant-baseline="middle"
      font-family="'Roboto Mono',monospace" font-size="17" fill="var(--text)" font-weight="800">${escapeHtml(String(centerValue))}</text>
  </svg>`;
}

/* ===================== js/auth.js ===================== */
// ============================================================
// auth.js — เข้าสู่ระบบ / จัดการ Session
// ============================================================

const SESSION_KEY = 'pc_session_v1';

let SESSION = null; // { role:'store'|'admin', username, storeCode, storeName, locNo }

function findAccount(username, password){
  username = (username||'').trim();
  password = (password||'').trim();
  if(!username || !password) return null;

  if(username === ADMIN_ACCOUNT.username && password === ADMIN_ACCOUNT.password){
    return { role:'admin', username, name: ADMIN_ACCOUNT.name };
  }
  const s = STORES_DATA.find(s=> s.username === username && s.password === password);
  if(s){
    return { role:'store', username: s.username, storeCode: s.username, storeName: s.name, locNo: s.locNo };
  }
  return null;
}

function restoreSession(){
  try{
    const raw = localStorage.getItem(SESSION_KEY);
    if(!raw) return null;
    return JSON.parse(raw);
  }catch(e){ return null; }
}

function saveSession(session){
  localStorage.setItem(SESSION_KEY, JSON.stringify(session));
}

function clearSession(){
  localStorage.removeItem(SESSION_KEY);
}

function initLoginForm(){
  const form = document.getElementById('loginForm');
  const errBox = document.getElementById('loginError');
  form.addEventListener('submit', (e)=>{
    e.preventDefault();
    const u = document.getElementById('loginUser').value;
    const p = document.getElementById('loginPass').value;
    const account = findAccount(u, p);
    if(!account){
      errBox.style.display = 'block';
      return;
    }
    errBox.style.display = 'none';
    SESSION = account;
    saveSession(account);
    startApp();
  });

  document.getElementById('logoutBtn').addEventListener('click', ()=>{
    clearSession();
    SESSION = null;
    location.reload();
  });
}

/* ===================== js/store-view.js ===================== */
// ============================================================
// store-view.js — หน้าจอสาขา: บันทึกตรวจนับรายวัน + ประวัติ/Export
// ============================================================

let CURRENT_DATE = todayStr();
let CURRENT_CAT  = 'FRESH';
let CURRENT_DATA = { FRESH:{}, TRANSFER:{}, NONFRESH:{} };
let CURRENT_META = null;
let DIRTY = { FRESH:false, TRANSFER:false, NONFRESH:false };

/* ============================================================
   ENTRY VIEW
============================================================ */
async function renderEntryView(){
  setTopbar('บันทึกการตรวจนับ Packing', `${SESSION.locNo} - ${SESSION.storeName}`);
  const content = document.getElementById('content');

  content.innerHTML = `
    <div class="card">
      <div class="card-head">
        <div>
          <div class="card-title">บันทึกการตรวจนับ Packing</div>
          <div class="muted">เลือกวันที่เพื่อบันทึก หรือแก้ไขข้อมูลย้อนหลัง ก่อนกด "บันทึกข้อมูล"</div>
        </div>
        <div class="form-row" style="align-items:flex-end;">
          <div class="form-group">
            <label>วันที่ตรวจนับ</label>
            <input type="date" id="entryDate" value="${CURRENT_DATE}" max="${todayStr()}">
          </div>
          <button class="btn btn-secondary" id="loadDateBtn">โหลดข้อมูล</button>
          <button class="btn btn-accent" id="exportTemplateBtn">📥 Export Excel</button>
        </div>
      </div>

      <div class="tabs" id="catTabs">
        ${['FRESH'].map(c=>`
          <div class="tab ${c===CURRENT_CAT?'active':''} ${DIRTY[c]?'dirty':''}" data-cat="${c}">
            ${CAT_LABELS[c]} <span class="dot"></span>
          </div>`).join('')}
      </div>

      <div class="mt-12" id="entryMeta"></div>
      <div class="table-wrap mt-12" id="entryTableWrap"></div>

      <div class="flex justify-between items-center mt-16" style="flex-wrap:wrap;gap:12px;">
        <div class="text-soft" id="entryTotals"></div>
        <button class="btn btn-primary" id="saveEntryBtn">บันทึกข้อมูล (${CAT_LABELS[CURRENT_CAT]})</button>
      </div>
    </div>
  `;

  document.getElementById('entryDate').addEventListener('change', e=>{
    CURRENT_DATE = e.target.value;
    loadDateData();
  });
  document.getElementById('loadDateBtn').addEventListener('click', ()=> loadDateData());

  content.querySelectorAll('#catTabs .tab').forEach(t=>{
    t.addEventListener('click', ()=>{
      CURRENT_CAT = t.dataset.cat;
      content.querySelectorAll('#catTabs .tab').forEach(x=> x.classList.toggle('active', x===t));
      renderEntryTable();
      document.getElementById('saveEntryBtn').textContent = `บันทึกข้อมูล (${CAT_LABELS[CURRENT_CAT]})`;
    });
  });

  document.getElementById('saveEntryBtn').addEventListener('click', saveCategory);
  document.getElementById('exportTemplateBtn').addEventListener('click', exportEntryTemplate);

  await loadDateData();
}

async function loadDateData(){
  const wrap = document.getElementById('entryTableWrap');
  wrap.innerHTML = '<div style="padding:48px;text-align:center;color:var(--text-soft)">กำลังโหลดข้อมูล...</div>';

  const data = await dbGetOnce(`counts/${CURRENT_DATE}/${SESSION.storeCode}`) || {};
  CURRENT_DATA = { FRESH: data.FRESH||{}, TRANSFER: data.TRANSFER||{}, NONFRESH: data.NONFRESH||{} };
  CURRENT_META = data._meta || null;
  DIRTY = { FRESH:false, TRANSFER:false, NONFRESH:false };

  updateTabDots();
  renderEntryMeta();
  renderEntryTable();
}

function renderEntryMeta(){
  const el = document.getElementById('entryMeta');
  if(!el) return;
  if(CURRENT_META){
    el.innerHTML = `<span class="pill pill-success">✔ มีข้อมูลของวันที่ ${thaiDate(CURRENT_DATE)} แล้ว</span>
      <span class="text-faint" style="margin-left:8px;font-size:12px;">
        ปรับปรุงล่าสุด ${fmtDateTime(CURRENT_META.updatedAt)} โดย ${escapeHtml(CURRENT_META.updatedBy||'-')}
      </span>`;
  } else {
    el.innerHTML = `<span class="pill pill-warning">⚠ ยังไม่มีข้อมูลของวันที่ ${thaiDate(CURRENT_DATE)}</span>`;
  }
}

function renderEntryTable(){
  const cat = CURRENT_CAT;
  const items = ITEMS_BY_CAT[cat] || [];
  const fields = cat === 'NONFRESH' ? ['QTY'] : ['FV','BUT','FISH'];
  const wrap = document.getElementById('entryTableWrap');

  if(items.length === 0){
    wrap.innerHTML = '<div style="padding:48px;text-align:center;color:var(--text-soft)">ไม่มีรายการสินค้าในหมวดนี้</div>';
    return;
  }

  const headCols = fields.map(f=>`<th class="text-right">${CAT_FIELD_LABELS[f]}</th>`).join('');

  const rows = items.map((item, idx)=>{
    const rec = CURRENT_DATA[cat][item.code];
    const inputCols = fields.map(f=>{
      const val = (rec && rec[f]!==undefined) ? rec[f] : '';
      return `<td class="text-right">
        <input class="qty-input" type="number" min="0" step="any" inputmode="decimal"
          id="inp_${cat}_${item.code}_${f}" data-cat="${cat}" data-code="${item.code}" data-field="${f}"
          value="${val===''?'':val}">
      </td>`;
    }).join('');
    const total = recordTotal(item, rec);
    const amount = recordAmount(item, rec);
    return `<tr>
      <td class="num text-soft">${idx+1}</td>
      <td class="desc"><div>${escapeHtml(item.desc)}</div><div class="code">${escapeHtml(item.code)}</div></td>
      <td class="nowrap text-soft">${escapeHtml(item.uomCount)}</td>
      ${inputCols}
      <td class="text-right num" id="total_${cat}_${item.code}">${fmtNum(total,2)}</td>
      <td class="text-right num" id="amount_${cat}_${item.code}">${fmtMoney(amount)}</td>
    </tr>`;
  }).join('');

  wrap.innerHTML = `
    <table class="dtable">
      <thead><tr>
        <th>#</th><th>รายการสินค้า</th><th>หน่วยนับ</th>${headCols}<th class="text-right">รวม</th><th class="text-right">มูลค่า (บาท)</th>
      </tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr>
        <td colspan="${3+fields.length}" class="text-right">รวมมูลค่าทั้งหมวด</td>
        <td class="text-right num" id="grandTotalQty">0</td>
        <td class="text-right num" id="grandTotalAmount">0.00</td>
      </tr></tfoot>
    </table>
  `;

  wrap.querySelectorAll('.qty-input').forEach(inp=> inp.addEventListener('input', onQtyInput));
  updateGrandTotals();
}

function onQtyInput(e){
  const inp = e.target;
  const { cat, code, field } = inp.dataset;
  const item = ITEM_MAP[cat][code];
  if(!item) return;

  const rec = {};
  item.subFields.forEach(f=>{
    const v = document.getElementById(`inp_${cat}_${code}_${f}`).value;
    rec[f] = v===''?0:(parseFloat(v)||0);
  });

  const total = recordTotal(item, rec);
  const amount = recordAmount(item, rec);
  const totalEl = document.getElementById(`total_${cat}_${code}`);
  const amountEl = document.getElementById(`amount_${cat}_${code}`);
  if(totalEl) totalEl.textContent = fmtNum(total, 2);
  if(amountEl) amountEl.textContent = fmtMoney(amount);

  const baseline = CURRENT_DATA[cat][code];
  const baseVal = (baseline && baseline[field]!==undefined) ? Number(baseline[field]) : 0;
  const curVal = inp.value===''?0:(parseFloat(inp.value)||0);
  inp.classList.toggle('changed', curVal !== baseVal);

  DIRTY[cat] = true;
  updateTabDots();
  updateGrandTotals();
}

function updateGrandTotals(){
  const cat = CURRENT_CAT;
  const items = ITEMS_BY_CAT[cat] || [];
  const fields = cat === 'NONFRESH' ? ['QTY'] : ['FV','BUT','FISH'];
  let totalQty = 0, totalAmount = 0, filledCount = 0;

  items.forEach(item=>{
    const rec = {};
    let has = false;
    fields.forEach(f=>{
      const inp = document.getElementById(`inp_${cat}_${item.code}_${f}`);
      const v = inp ? inp.value : '';
      if(v !== '') has = true;
      rec[f] = v===''?0:(parseFloat(v)||0);
    });
    if(has) filledCount++;
    totalQty += recordTotal(item, rec);
    totalAmount += recordAmount(item, rec);
  });

  const qtyEl = document.getElementById('grandTotalQty');
  const amtEl = document.getElementById('grandTotalAmount');
  if(qtyEl) qtyEl.textContent = fmtNum(totalQty, 2);
  if(amtEl) amtEl.textContent = fmtMoney(totalAmount);

  const totalsEl = document.getElementById('entryTotals');
  if(totalsEl){
    totalsEl.innerHTML = `กรอกแล้ว <b class="num">${filledCount}</b> / ${items.length} รายการ
      &nbsp;&middot;&nbsp; มูลค่ารวม <b class="num">${fmtMoney(totalAmount)}</b> บาท`;
  }
}

function updateTabDots(){
  document.querySelectorAll('#catTabs .tab').forEach(t=>{
    t.classList.toggle('dirty', !!DIRTY[t.dataset.cat]);
  });
}

/* ============================================================
   SAVE
============================================================ */
async function saveCategory(){
  const cat = CURRENT_CAT;
  const items = ITEMS_BY_CAT[cat] || [];
  const updates = {};
  const changesLog = [];
  const newDataCat = {};

  items.forEach(item=>{
    const fields = item.subFields;
    let hasInput = false;
    const rec = {};
    fields.forEach(f=>{
      const inp = document.getElementById(`inp_${cat}_${item.code}_${f}`);
      const raw = inp ? inp.value : '';
      if(raw !== '') hasInput = true;
      rec[f] = raw===''?0:(parseFloat(raw)||0);
    });

    const baseline = CURRENT_DATA[cat][item.code];
    const recsEqual = (a,b)=>{
      if(!a && !b) return true;
      if(!a || !b) return false;
      return fields.every(f=> (Number(a[f])||0) === (Number(b[f])||0));
    };

    if(hasInput){
      if(!recsEqual(baseline, rec)){
        updates[`counts/${CURRENT_DATE}/${SESSION.storeCode}/${cat}/${item.code}`] = rec;
        fields.forEach(f=>{
          const ov = (baseline && Number(baseline[f])) || 0;
          const nv = Number(rec[f]) || 0;
          if(ov !== nv) changesLog.push({ itemCode:item.code, itemDesc:item.desc, field:f, oldVal:ov, newVal:nv });
        });
      }
      newDataCat[item.code] = rec;
    } else if(baseline){
      updates[`counts/${CURRENT_DATE}/${SESSION.storeCode}/${cat}/${item.code}`] = null;
      fields.forEach(f=>{
        const ov = Number(baseline[f]) || 0;
        if(ov !== 0) changesLog.push({ itemCode:item.code, itemDesc:item.desc, field:f, oldVal:ov, newVal:0 });
      });
    }
  });

  if(changesLog.length === 0){
    toast('ไม่มีข้อมูลที่เปลี่ยนแปลง');
    return;
  }

  const btn = document.getElementById('saveEntryBtn');
  setLoading(btn, true, 'กำลังบันทึก...');
  try{
    updates[`counts/${CURRENT_DATE}/${SESSION.storeCode}/_meta`] = {
      storeName: SESSION.storeName,
      locNo: SESSION.locNo,
      updatedAt: Date.now(),
      updatedBy: SESSION.username
    };
    await dbUpdate(updates);
    await dbPush('logs', {
      ts: Date.now(),
      date: CURRENT_DATE,
      store: SESSION.storeCode,
      storeName: SESSION.storeName,
      user: SESSION.username,
      action: 'SAVE',
      category: cat,
      changes: changesLog
    });

    CURRENT_DATA[cat] = newDataCat;
    DIRTY[cat] = false;
    updateTabDots();
    renderEntryMeta();
    document.querySelectorAll('#entryTableWrap .qty-input.changed').forEach(i=> i.classList.remove('changed'));
    toast(`บันทึกข้อมูลเรียบร้อย (${changesLog.length} การเปลี่ยนแปลง)`, 'success');
  }catch(err){
    console.error(err);
    toast('เกิดข้อผิดพลาดในการบันทึก: ' + err.message, 'error');
  }finally{
    setLoading(btn, false);
  }
}

/* ============================================================
   EXPORT TEMPLATE — ส่งออกรายการสินค้าสำหรับใช้ตรวจนับ
============================================================ */
function exportEntryTemplate(){
  const cat = CURRENT_CAT;
  const items = ITEMS_BY_CAT[cat] || [];
  if(items.length === 0){ toast('ไม่มีรายการสินค้าในหมวดนี้', 'error'); return; }

  const fields = cat === 'NONFRESH' ? ['QTY'] : ['F&V', 'BUT', 'FISH'];
  const dateLabel = CURRENT_DATE;

  // สร้าง header row
  const headerRow = {
    '#':              '#',
    'รหัสสินค้า':     'รหัสสินค้า',
    'รายการสินค้า':   'รายการสินค้า',
    'หน่วยนับ':       'หน่วยนับ',
  };
  fields.forEach(f => { headerRow[f] = f; });
  headerRow['รวม']        = 'รวม';
  headerRow['หมายเหตุ']   = 'หมายเหตุ';

  // สร้างแถวข้อมูลสินค้า (ช่องตัวเลขว่างเปล่า)
  const rows = items.map((item, idx) => {
    const row = {
      '#':              idx + 1,
      'รหัสสินค้า':     item.code,
      'รายการสินค้า':   item.desc,
      'หน่วยนับ':       item.uomCount,
    };
    fields.forEach(f => { row[f] = ''; });
    row['รวม']      = '';
    row['หมายเหตุ'] = '';
    return row;
  });

  // Sheet ชื่อตามวันที่และสาขา
  const sheetName = `${SESSION.locNo}_${dateLabel}`.substring(0, 31);
  const filename  = `Template_${SESSION.locNo}_${SESSION.storeName.replace(/[^a-zA-Z0-9ก-๙]/g,'_').substring(0,20)}_${dateLabel}.xlsx`;

  // สร้าง workbook พร้อม title row บนสุด
  const wb = XLSX.utils.book_new();

  // แปลง rows → worksheet แล้วเพิ่ม title 2 แถวด้านบน
  const titleRows = [
    [`แบบฟอร์มตรวจนับ Packing — ${CAT_LABELS[cat]}`],
    [`สาขา: ${SESSION.locNo} - ${SESSION.storeName}  |  วันที่: ${thaiDate(dateLabel)}`],
    [], // แถวว่างคั่น
  ];

  // รวม title + header + data เข้าด้วยกันเป็น array of arrays
  const colKeys = ['#', 'รหัสสินค้า', 'รายการสินค้า', 'หน่วยนับ', ...fields, 'รวม', 'หมายเหตุ'];
  const headerArr = colKeys;
  const dataArrs  = rows.map(r => colKeys.map(k => r[k]));

  const allArrs = [...titleRows, headerArr, ...dataArrs];

  const ws = XLSX.utils.aoa_to_sheet(allArrs);

  // กำหนดความกว้างคอลัมน์
  ws['!cols'] = [
    { wch: 5  },   // #
    { wch: 18 },   // รหัสสินค้า
    { wch: 52 },   // รายการสินค้า
    { wch: 14 },   // หน่วยนับ
    ...fields.map(() => ({ wch: 10 })),
    { wch: 10 },   // รวม
    { wch: 20 },   // หมายเหตุ
  ];

  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
  toast(`Export แบบฟอร์มตรวจนับเรียบร้อย (${items.length} รายการ)`, 'success');
}

/* ============================================================
   HISTORY / EXPORT VIEW
============================================================ */
async function renderHistoryView(){
  setTopbar('ประวัติการตรวจนับ / Export Excel', `${SESSION.locNo} - ${SESSION.storeName}`);
  const content = document.getElementById('content');
  const thisMonth = todayStr().slice(0,7);

  content.innerHTML = `
    <div class="card">
      <div class="card-head">
        <div>
          <div class="card-title">ค้นหาข้อมูลรายเดือน</div>
          <div class="muted">เลือกเดือนเพื่อดูวันที่มีการบันทึกข้อมูลแล้ว และเปิดแก้ไขย้อนหลัง</div>
        </div>
        <div class="form-row">
          <div class="form-group"><label>เดือน</label><input type="month" id="histMonth" value="${thisMonth}"></div>
          <button class="btn btn-secondary" id="histSearchBtn">ค้นหา</button>
          <button class="btn btn-accent" id="histExportBtn">Export เดือนนี้</button>
        </div>
      </div>
      <div id="histResult"><div class="text-soft" style="padding:32px;text-align:center;">เลือกเดือนแล้วกด "ค้นหา"</div></div>
    </div>

    <div class="card">
      <div class="card-head"><div><div class="card-title">Export ตามช่วงวันที่</div>
        <div class="muted">เลือกช่วงวันที่และหมวดหมู่ที่ต้องการ Export เป็นไฟล์ Excel</div></div></div>
      <div class="form-row">
        <div class="form-group"><label>จากวันที่</label><input type="date" id="expFrom" max="${todayStr()}"></div>
        <div class="form-group"><label>ถึงวันที่</label><input type="date" id="expTo" max="${todayStr()}"></div>
        <div class="form-group"><label>หมวดหมู่</label>
          <select id="expCat">
            <option value="ALL">ทั้งหมด</option>
            <option value="FRESH">FRESH FOOD</option>
            <option value="TRANSFER">TRANSFER</option>
            <option value="NONFRESH">NON FRESH</option>
          </select>
        </div>
        <button class="btn btn-accent" id="rangeExportBtn">Export Excel</button>
      </div>
    </div>
  `;

  document.getElementById('histSearchBtn').addEventListener('click', searchHistoryMonth);
  document.getElementById('histExportBtn').addEventListener('click', ()=>{
    const month = document.getElementById('histMonth').value;
    if(!month){ toast('กรุณาเลือกเดือน','error'); return; }
    const days = daysInMonth(month);
    exportStoreDateRange(`${month}-01`, `${month}-${pad2(days)}`, 'ALL', `PackagingCount_${storeFileTag()}_${month}.xlsx`);
  });
  document.getElementById('rangeExportBtn').addEventListener('click', ()=>{
    const from = document.getElementById('expFrom').value;
    const to = document.getElementById('expTo').value;
    const cat = document.getElementById('expCat').value;
    if(!from || !to){ toast('กรุณาเลือกช่วงวันที่','error'); return; }
    if(from > to){ toast('วันที่เริ่มต้องไม่เกินวันที่สิ้นสุด','error'); return; }
    exportStoreDateRange(from, to, cat, `PackagingCount_${storeFileTag()}_${from}_to_${to}.xlsx`);
  });
}

async function searchHistoryMonth(){
  const month = document.getElementById('histMonth').value;
  if(!month) return;
  const resultEl = document.getElementById('histResult');
  resultEl.innerHTML = '<div class="text-soft" style="padding:32px;text-align:center;">กำลังค้นหา...</div>';

  const days = daysInMonth(month);
  const promises = [];
  for(let d=1; d<=days; d++){
    const dateStr = `${month}-${pad2(d)}`;
    promises.push(dbGetOnce(`counts/${dateStr}/${SESSION.storeCode}/_meta`).then(meta=>({dateStr, meta})));
  }
  const results = await Promise.all(promises);
  const found = results.filter(r=>r.meta);

  if(found.length === 0){
    resultEl.innerHTML = '<div class="text-soft" style="padding:32px;text-align:center;">ไม่พบข้อมูลในเดือนที่เลือก</div>';
    return;
  }

  resultEl.innerHTML = `
    <div class="table-wrap mt-12">
      <table class="dtable">
        <thead><tr><th>วันที่</th><th>ปรับปรุงล่าสุด</th><th>บันทึกโดย</th><th></th></tr></thead>
        <tbody>
          ${found.map(r=>`
            <tr>
              <td class="num">${thaiDate(r.dateStr)}</td>
              <td class="num">${fmtDateTime(r.meta.updatedAt)}</td>
              <td>${escapeHtml(r.meta.updatedBy || '-')}</td>
              <td><button class="btn btn-secondary btn-sm" data-jump="${r.dateStr}">เปิด / แก้ไข</button></td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>
  `;
  resultEl.querySelectorAll('[data-jump]').forEach(b=>{
    b.addEventListener('click', ()=>{
      CURRENT_DATE = b.dataset.jump;
      navigateTo('entry');
    });
  });
}

function storeFileTag(){
  return `${SESSION.locNo}_${SESSION.storeCode}`;
}

async function exportStoreDateRange(from, to, catFilter, filename){
  const dates = dateRange(from, to);
  if(dates.length === 0){ toast('ช่วงวันที่ไม่ถูกต้อง','error'); return; }
  if(dates.length > 370){ toast('ช่วงวันที่กว้างเกินไป (เกิน 1 ปี)','error'); return; }

  toast('กำลังเตรียมไฟล์ Export...');
  const rows = [];
  for(const dateStr of dates){
    const data = await dbGetOnce(`counts/${dateStr}/${SESSION.storeCode}`);
    if(!data) continue;
    ['FRESH','TRANSFER','NONFRESH'].forEach(cat=>{
      if(catFilter !== 'ALL' && catFilter !== cat) return;
      const catData = data[cat];
      if(!catData) return;
      Object.entries(catData).forEach(([code, rec])=>{
        const item = ITEM_MAP[cat][code];
        if(!item) return;
        rows.push(buildExportRow(dateStr, SESSION.storeCode, cat, item, rec));
      });
    });
  }

  if(rows.length === 0){ toast('ไม่พบข้อมูลในช่วงที่เลือก','error'); return; }
  exportRowsToExcel({ 'ข้อมูลตรวจนับ': rows }, filename);
  toast('Export ไฟล์เรียบร้อย', 'success');
}

/* ===================== js/admin-view.js ===================== */
// ============================================================
// admin-view.js — หน้าจอผู้ดูแลระบบ (Admin)
// ============================================================

let ADMIN_DATA_ROWS = []; // cache of last search results for edit/delete

/* ============================================================
   OVERVIEW
============================================================ */
async function renderAdminOverview(){
  setTopbar('ภาพรวมระบบ', 'Packaging Count — CP Axtra / Makro');
  const content = document.getElementById('content');
  const today = todayStr();

  content.innerHTML = `
    <div class="card">
      <div class="card-head">
        <div>
          <div class="card-title">ภาพรวมการบันทึกข้อมูล</div>
          <div class="muted">เลือกวันที่เพื่อดูสถานะการบันทึกของแต่ละสาขา</div>
        </div>
        <div class="form-row">
          <div class="form-group"><label>วันที่</label><input type="date" id="ovDate" value="${today}" max="${today}"></div>
          <button class="btn btn-secondary" id="ovRefreshBtn">รีเฟรช</button>
        </div>
      </div>
      <div class="stat-grid" id="ovStats">
        <div class="stat-card"><div class="label">กำลังโหลด...</div></div>
      </div>
    </div>

    <div class="card">
      <div class="card-head">
        <div class="card-title">รายการสินค้าในระบบ</div>
      </div>
      <div class="stat-grid">
        <div class="stat-card"><div class="label">FRESH FOOD</div><div class="value">${ITEMS_BY_CAT.FRESH.length}</div><div class="sub">รายการ &middot; ตรวจนับ F&amp;V / BUT / FISH</div></div>
        <div class="stat-card"><div class="label">สาขาทั้งหมด</div><div class="value">${STORES_DATA.length}</div><div class="sub">สาขาในระบบ</div></div>
      </div>
    </div>
  `;

  document.getElementById('ovRefreshBtn').addEventListener('click', loadOverviewStats);
  document.getElementById('ovDate').addEventListener('change', loadOverviewStats);
  await loadOverviewStats();
}

async function loadOverviewStats(){
  const date = document.getElementById('ovDate').value;
  const statsEl = document.getElementById('ovStats');
  statsEl.innerHTML = `<div class="stat-card"><div class="label">กำลังโหลด...</div></div>`;

  const data = await dbGetOnce(`counts/${date}`) || {};
  const storeCodes = Object.keys(data);
  const submitted = storeCodes.length;
  const total = STORES_DATA.length;
  const pct = total ? Math.round((submitted/total)*100) : 0;

  let grandAmount = 0;
  storeCodes.forEach(code=>{
    const rec = data[code];
    ['FRESH','TRANSFER','NONFRESH'].forEach(cat=>{
      const catData = rec[cat];
      if(!catData) return;
      Object.entries(catData).forEach(([itemCode, r])=>{
        const item = ITEM_MAP[cat][itemCode];
        if(item) grandAmount += recordAmount(item, r);
      });
    });
  });

  statsEl.innerHTML = `
    <div class="stat-card">
      <div class="label">สาขาที่บันทึกแล้ว (${thaiDate(date)})</div>
      <div class="value">${fmtNum(submitted)} / ${fmtNum(total)}</div>
      <div class="bar"><div style="width:${pct}%"></div></div>
      <div class="sub">${pct}% ของสาขาทั้งหมด</div>
    </div>
    <div class="stat-card">
      <div class="label">สาขาที่ยังไม่บันทึก</div>
      <div class="value">${fmtNum(total - submitted)}</div>
      <div class="sub">สาขา</div>
    </div>
    <div class="stat-card">
      <div class="label">มูลค่ารวมที่บันทึกของวันนี้</div>
      <div class="value">${fmtMoney(grandAmount)}</div>
      <div class="sub">บาท (ทุกสาขา / ทุกหมวด)</div>
    </div>
  `;
}

/* ============================================================
   DATA BROWSER
============================================================ */
async function renderAdminData(){
  setTopbar('ข้อมูลการตรวจนับ', 'ค้นหา / แก้ไข / ลบ ข้อมูลรายการตรวจนับ');
  const content = document.getElementById('content');
  const today = todayStr();

  const storeOptions = STORES_DATA
    .slice()
    .sort((a,b)=> Number(a.locNo) - Number(b.locNo))
    .map(s=>`<option value="${s.username}">${s.locNo} - ${escapeHtml(s.name)}</option>`).join('');

  content.innerHTML = `
    <div class="card">
      <div class="card-head">
        <div>
          <div class="card-title">ค้นหาข้อมูลการตรวจนับ</div>
          <div class="muted">เลือกสาขา / ช่วงวันที่ / หมวดหมู่ แล้วกด "ค้นหา"</div>
        </div>
      </div>
      <div class="form-row">
        <div class="form-group" style="min-width:220px;">
          <label>สาขา</label>
          <select id="dataStore">
            <option value="ALL">-- ทุกสาขา --</option>
            ${storeOptions}
          </select>
        </div>
        <div class="form-group"><label>จากวันที่</label><input type="date" id="dataFrom" value="${today}" max="${today}"></div>
        <div class="form-group"><label>ถึงวันที่</label><input type="date" id="dataTo" value="${today}" max="${today}"></div>
        <div class="form-group">
          <label>หมวดหมู่</label>
          <select id="dataCat">
            <option value="ALL">ทั้งหมด</option>
            <option value="FRESH">FRESH FOOD</option>
            <option value="TRANSFER">TRANSFER</option>
            <option value="NONFRESH">NON FRESH</option>
          </select>
        </div>
        <button class="btn btn-primary" id="dataSearchBtn">ค้นหา</button>
      </div>
      <div class="mt-12 text-faint" style="font-size:12px;">
        คำแนะนำ: หากเลือก "ทุกสาขา" ในช่วงวันที่ยาว ระบบจะอ่านข้อมูล 1 ครั้งต่อวัน (ครอบคลุมทุกสาขาในการอ่านครั้งเดียว) จึงรองรับข้อมูลจำนวนมากได้รวดเร็ว
      </div>
      <div class="mt-12" id="dataResultInfo"></div>
      <div class="table-wrap mt-12" id="dataResultWrap"></div>
    </div>
  `;

  document.getElementById('dataSearchBtn').addEventListener('click', searchAdminData);
}

async function searchAdminData(){
  const storeFilter = document.getElementById('dataStore').value;
  const from = document.getElementById('dataFrom').value;
  const to = document.getElementById('dataTo').value;
  const catFilter = document.getElementById('dataCat').value;
  const infoEl = document.getElementById('dataResultInfo');
  const wrapEl = document.getElementById('dataResultWrap');

  if(!from || !to || from > to){
    toast('กรุณาเลือกช่วงวันที่ให้ถูกต้อง', 'error');
    return;
  }
  const dates = dateRange(from, to);
  if(dates.length > 370){
    toast('ช่วงวันที่กว้างเกินไป (เกิน 1 ปี)', 'error');
    return;
  }

  infoEl.innerHTML = `<div class="text-soft">กำลังค้นหา... (${dates.length} วัน)</div>`;
  wrapEl.innerHTML = '';

  ADMIN_DATA_ROWS = [];

  for(const dateStr of dates){
    const data = await dbGetOnce(`counts/${dateStr}`);
    if(!data) continue;

    const storeCodes = storeFilter === 'ALL' ? Object.keys(data) : (data[storeFilter] ? [storeFilter] : []);
    storeCodes.forEach(storeCode=>{
      const rec = data[storeCode];
      if(!rec) return;
      ['FRESH','TRANSFER','NONFRESH'].forEach(cat=>{
        if(catFilter !== 'ALL' && catFilter !== cat) return;
        const catData = rec[cat];
        if(!catData) return;
        Object.entries(catData).forEach(([itemCode, r])=>{
          const item = ITEM_MAP[cat][itemCode] || placeholderItem(cat, itemCode);
          ADMIN_DATA_ROWS.push({ dateStr, storeCode, cat, itemCode, item, rec: r });
        });
      });
    });
  }

  if(ADMIN_DATA_ROWS.length === 0){
    infoEl.innerHTML = '<div class="text-soft">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</div>';
    return;
  }

  const MAX_SHOW = 1000;
  const showRows = ADMIN_DATA_ROWS.slice(0, MAX_SHOW);
  infoEl.innerHTML = `พบ <b class="num">${fmtNum(ADMIN_DATA_ROWS.length)}</b> รายการ
    ${ADMIN_DATA_ROWS.length > MAX_SHOW ? `<span class="pill pill-warning" style="margin-left:8px;">แสดงเฉพาะ ${fmtNum(MAX_SHOW)} รายการแรก — กรุณาเลือกช่วงให้แคบลงหากต้องการดูทั้งหมด</span>` : ''}`;

  renderAdminDataTable(showRows);
}

function placeholderItem(cat, code){
  return { category: cat, code, desc:'(ไม่พบในรายการสินค้า — อาจถูกลบแล้ว)', supplier:'', price:0, uomCount:'', packCount:1,
    subFields: cat === 'NONFRESH' ? ['QTY'] : ['FV','BUT','FISH'] };
}

function renderAdminDataTable(rows){
  const wrapEl = document.getElementById('dataResultWrap');
  const trs = rows.map((row, idx)=>{
    const { dateStr, storeCode, cat, item, rec } = row;
    const total = recordTotal(item, rec);
    const amount = recordAmount(item, rec);
    const fields = item.subFields;
    const fieldVals = fields.map(f=>`<span class="text-soft">${CAT_FIELD_LABELS[f]}:</span> <b class="num">${fmtNum(rec[f]||0,2)}</b>`).join('&nbsp;&nbsp;');
    return `<tr>
      <td class="num nowrap">${thaiDate(dateStr)}</td>
      <td class="nowrap">${escapeHtml(storeLabel(storeCode))}</td>
      <td><span class="pill pill-info">${CAT_LABELS[cat]}</span></td>
      <td class="desc"><div>${escapeHtml(item.desc)}</div><div class="code">${escapeHtml(item.code)}</div></td>
      <td class="nowrap">${fieldVals}</td>
      <td class="text-right num">${fmtNum(total,2)}</td>
      <td class="text-right num">${fmtMoney(amount)}</td>
      <td class="nowrap">
        <button class="btn btn-secondary btn-sm" data-edit="${idx}">แก้ไข</button>
        <button class="btn btn-danger btn-sm" data-del="${idx}">ลบ</button>
      </td>
    </tr>`;
  }).join('');

  wrapEl.innerHTML = `
    <table class="dtable">
      <thead><tr>
        <th>วันที่</th><th>สาขา</th><th>หมวดหมู่</th><th>รายการสินค้า</th><th>จำนวนที่บันทึก</th><th class="text-right">รวม</th><th class="text-right">มูลค่า</th><th>การจัดการ</th>
      </tr></thead>
      <tbody>${trs}</tbody>
    </table>
  `;

  wrapEl.querySelectorAll('[data-edit]').forEach(b=> b.addEventListener('click', ()=> openEditRecordModal(Number(b.dataset.edit))));
  wrapEl.querySelectorAll('[data-del]').forEach(b=> b.addEventListener('click', ()=> confirmDeleteRecord(Number(b.dataset.del))));
}

function openEditRecordModal(idx){
  const row = ADMIN_DATA_ROWS[idx];
  const { dateStr, storeCode, cat, item, rec } = row;
  const fields = item.subFields;

  const inputsHtml = fields.map(f=>`
    <div class="field">
      <label>${CAT_FIELD_LABELS[f]}</label>
      <input type="number" min="0" step="any" id="editField_${f}" value="${rec[f]!==undefined?rec[f]:0}">
    </div>
  `).join('');

  showModal(`
    <h3>แก้ไขข้อมูลการตรวจนับ</h3>
    <div class="text-soft" style="font-size:13px;margin-bottom:14px;">
      <div><b>วันที่:</b> ${thaiDate(dateStr)}</div>
      <div><b>สาขา:</b> ${escapeHtml(storeLabel(storeCode))}</div>
      <div><b>รายการ:</b> ${escapeHtml(item.desc)} <span class="code">(${escapeHtml(item.code)})</span></div>
    </div>
    ${inputsHtml}
    <div class="modal-actions">
      <button class="btn btn-secondary" id="editCancelBtn">ยกเลิก</button>
      <button class="btn btn-primary" id="editSaveBtn">บันทึก</button>
    </div>
  `, ()=>{
    document.getElementById('editCancelBtn').addEventListener('click', closeModal);
    document.getElementById('editSaveBtn').addEventListener('click', ()=> saveEditedRecord(idx));
  });
}

async function saveEditedRecord(idx){
  const row = ADMIN_DATA_ROWS[idx];
  const { dateStr, storeCode, cat, item, rec } = row;
  const fields = item.subFields;
  const newRec = {};
  const changes = [];

  fields.forEach(f=>{
    const inp = document.getElementById(`editField_${f}`);
    const nv = parseFloat(inp.value) || 0;
    const ov = Number(rec[f]) || 0;
    newRec[f] = nv;
    if(nv !== ov) changes.push({ itemCode:item.code, itemDesc:item.desc, field:f, oldVal:ov, newVal:nv });
  });

  if(changes.length === 0){
    toast('ไม่มีการเปลี่ยนแปลง');
    closeModal();
    return;
  }

  try{
    const updates = {};
    updates[`counts/${dateStr}/${storeCode}/${cat}/${item.code}`] = newRec;
    updates[`counts/${dateStr}/${storeCode}/_meta/updatedAt`] = Date.now();
    updates[`counts/${dateStr}/${storeCode}/_meta/updatedBy`] = `${SESSION.username} (admin)`;
    await dbUpdate(updates);
    await dbPush('logs', {
      ts: Date.now(), date: dateStr, store: storeCode, storeName: getStoreByCode(storeCode)?.name || storeCode,
      user: SESSION.username, action: 'UPDATE', category: cat, changes
    });

    row.rec = newRec;
    closeModal();
    toast('แก้ไขข้อมูลเรียบร้อย', 'success');
    renderAdminDataTable(ADMIN_DATA_ROWS.slice(0, 1000));
  }catch(err){
    console.error(err);
    toast('เกิดข้อผิดพลาด: ' + err.message, 'error');
  }
}

function confirmDeleteRecord(idx){
  const row = ADMIN_DATA_ROWS[idx];
  const { dateStr, storeCode, cat, item } = row;
  showModal(`
    <h3>ยืนยันการลบข้อมูล</h3>
    <div class="text-soft" style="font-size:13px;margin-bottom:14px;">
      ต้องการลบข้อมูลตรวจนับของรายการนี้ใช่หรือไม่?<br><br>
      <b>วันที่:</b> ${thaiDate(dateStr)}<br>
      <b>สาขา:</b> ${escapeHtml(storeLabel(storeCode))}<br>
      <b>รายการ:</b> ${escapeHtml(item.desc)} <span class="code">(${escapeHtml(item.code)})</span>
    </div>
    <div class="modal-actions">
      <button class="btn btn-secondary" id="delCancelBtn">ยกเลิก</button>
      <button class="btn btn-danger" id="delConfirmBtn">ลบข้อมูล</button>
    </div>
  `, ()=>{
    document.getElementById('delCancelBtn').addEventListener('click', closeModal);
    document.getElementById('delConfirmBtn').addEventListener('click', ()=> deleteRecord(idx));
  });
}

async function deleteRecord(idx){
  const row = ADMIN_DATA_ROWS[idx];
  const { dateStr, storeCode, cat, item, rec } = row;
  const fields = item.subFields;
  const changes = fields.map(f=>({ itemCode:item.code, itemDesc:item.desc, field:f, oldVal:Number(rec[f])||0, newVal:0 }));

  try{
    const updates = {};
    updates[`counts/${dateStr}/${storeCode}/${cat}/${item.code}`] = null;
    updates[`counts/${dateStr}/${storeCode}/_meta/updatedAt`] = Date.now();
    updates[`counts/${dateStr}/${storeCode}/_meta/updatedBy`] = `${SESSION.username} (admin)`;
    await dbUpdate(updates);
    await dbPush('logs', {
      ts: Date.now(), date: dateStr, store: storeCode, storeName: getStoreByCode(storeCode)?.name || storeCode,
      user: SESSION.username, action: 'DELETE', category: cat, changes
    });

    ADMIN_DATA_ROWS.splice(idx, 1);
    closeModal();
    toast('ลบข้อมูลเรียบร้อย', 'success');
    renderAdminDataTable(ADMIN_DATA_ROWS.slice(0, 1000));
  }catch(err){
    console.error(err);
    toast('เกิดข้อผิดพลาด: ' + err.message, 'error');
  }
}

/* ============================================================
   ITEM MASTER
============================================================ */
let ADMIN_ITEMS_CAT = 'FRESH';

async function renderAdminItems(){
  ADMIN_ITEMS_CAT = 'FRESH'; // รีเซ็ตเสมอ เนื่องจากแสดงเฉพาะ FRESH FOOD
  setTopbar('รายการสินค้า (Item Master)', 'เพิ่ม / แก้ไข / ลบ รายการสินค้าในระบบ');
  const content = document.getElementById('content');

  content.innerHTML = `
    <div class="card">
      <div class="card-head">
        <div>
          <div class="card-title">รายการสินค้า</div>
          <div class="muted">รายการที่เพิ่มหรือแก้ไขที่นี่ จะมีผลกับหน้าบันทึกของทุกสาขาทันที</div>
        </div>
        <button class="btn btn-primary" id="addItemBtn">+ เพิ่มรายการสินค้า</button>
      </div>
      <div class="tabs" id="itemCatTabs">
        ${['FRESH'].map(c=>`<div class="tab ${c===ADMIN_ITEMS_CAT?'active':''}" data-cat="${c}">${CAT_LABELS[c]} (${ITEMS_BY_CAT[c].length})</div>`).join('')}
      </div>
      <div class="table-wrap mt-12" id="itemsTableWrap"></div>
    </div>
  `;

  content.querySelectorAll('#itemCatTabs .tab').forEach(t=>{
    t.addEventListener('click', ()=>{
      ADMIN_ITEMS_CAT = t.dataset.cat;
      content.querySelectorAll('#itemCatTabs .tab').forEach(x=> x.classList.toggle('active', x===t));
      renderItemsTable();
    });
  });
  document.getElementById('addItemBtn').addEventListener('click', ()=> openItemModal(null));

  renderItemsTable();
}

function renderItemsTable(){
  const cat = ADMIN_ITEMS_CAT;
  const items = ITEMS_BY_CAT[cat] || [];
  const wrap = document.getElementById('itemsTableWrap');

  const rows = items.map((item, idx)=>`
    <tr>
      <td class="num text-soft">${idx+1}</td>
      <td class="code">${escapeHtml(item.code)}</td>
      <td class="desc">${escapeHtml(item.desc)}</td>
      <td>${escapeHtml(item.supplier||'')}</td>
      <td class="text-right num">${fmtMoney(item.price)}</td>
      <td class="nowrap">${escapeHtml(item.uomRec||'')}</td>
      <td class="text-right num">${fmtNum(item.packRec,2)}</td>
      <td class="nowrap">${escapeHtml(item.uomCount||'')}</td>
      <td class="text-right num">${fmtNum(item.packCount,2)}</td>
      <td class="nowrap">
        <button class="btn btn-secondary btn-sm" data-edit-item="${item.code}">แก้ไข</button>
        <button class="btn btn-danger btn-sm" data-del-item="${item.code}">ลบ</button>
      </td>
    </tr>
  `).join('');

  wrap.innerHTML = `
    <table class="dtable">
      <thead><tr>
        <th>#</th><th>รหัสสินค้า</th><th>รายการสินค้า</th><th>Supplier</th>
        <th class="text-right">ราคา (บาท)</th><th>หน่วยรับ</th><th class="text-right">Pack รับ</th>
        <th>หน่วยตรวจนับ</th><th class="text-right">Pack ตรวจนับ</th><th>การจัดการ</th>
      </tr></thead>
      <tbody>${rows || `<tr><td colspan="10" class="text-center text-soft" style="padding:24px;">ไม่มีรายการ</td></tr>`}</tbody>
    </table>
  `;

  wrap.querySelectorAll('[data-edit-item]').forEach(b=> b.addEventListener('click', ()=> openItemModal(b.dataset.editItem)));
  wrap.querySelectorAll('[data-del-item]').forEach(b=> b.addEventListener('click', ()=> confirmDeleteItem(b.dataset.delItem)));
}

function openItemModal(code){
  const cat = ADMIN_ITEMS_CAT;
  const isEdit = !!code;
  const item = isEdit ? ITEM_MAP[cat][code] : null;

  showModal(`
    <h3>${isEdit ? 'แก้ไขรายการสินค้า' : 'เพิ่มรายการสินค้าใหม่'}</h3>
    <div class="field"><label>หมวดหมู่</label>
      <input type="text" value="${CAT_LABELS[cat]}" disabled>
    </div>
    <div class="field"><label>รหัสสินค้า (Item Code)</label>
      <input type="text" id="itCode" value="${isEdit ? escapeHtml(code) : ''}" ${isEdit?'disabled':''} placeholder="เช่น 0240080001Y">
    </div>
    <div class="field"><label>รายการสินค้า</label>
      <input type="text" id="itDesc" value="${isEdit ? escapeHtml(item.desc) : ''}">
    </div>
    <div class="field"><label>Supplier</label>
      <input type="text" id="itSupplier" value="${isEdit ? escapeHtml(item.supplier||'') : ''}">
    </div>
    <div class="form-row">
      <div class="form-group flex-1"><label>ราคา (บาท)</label>
        <input type="number" step="any" id="itPrice" value="${isEdit ? item.price : 0}"></div>
      <div class="form-group flex-1"><label>หน่วยรับ (UOM รับ)</label>
        <input type="text" id="itUomRec" value="${isEdit ? escapeHtml(item.uomRec||'') : ''}"></div>
    </div>
    <div class="form-row">
      <div class="form-group flex-1"><label>Pack รับ (Unit Pack Rec)</label>
        <input type="number" step="any" id="itPackRec" value="${isEdit ? item.packRec : 1}"></div>
      <div class="form-group flex-1"><label>หน่วยตรวจนับ</label>
        <input type="text" id="itUomCount" value="${isEdit ? escapeHtml(item.uomCount||'') : ''}"></div>
    </div>
    <div class="form-group"><label>Pack ตรวจนับ (Unit Pack Count)</label>
      <input type="number" step="any" id="itPackCount" value="${isEdit ? item.packCount : 1}"></div>

    <div class="modal-actions">
      <button class="btn btn-secondary" id="itCancelBtn">ยกเลิก</button>
      <button class="btn btn-primary" id="itSaveBtn">บันทึก</button>
    </div>
  `, ()=>{
    document.getElementById('itCancelBtn').addEventListener('click', closeModal);
    document.getElementById('itSaveBtn').addEventListener('click', ()=> saveItem(isEdit, code));
  });
}

async function saveItem(isEdit, oldCode){
  const cat = ADMIN_ITEMS_CAT;
  const code = isEdit ? oldCode : document.getElementById('itCode').value.trim();
  const desc = document.getElementById('itDesc').value.trim();

  if(!code || !desc){ toast('กรุณากรอกรหัสสินค้าและชื่อรายการ', 'error'); return; }
  if(!isEdit && ITEM_MAP[cat][code]){ toast('รหัสสินค้านี้มีอยู่แล้วในหมวดนี้', 'error'); return; }

  const newItem = {
    code,
    desc,
    supplier: document.getElementById('itSupplier').value.trim(),
    price: parseFloat(document.getElementById('itPrice').value) || 0,
    uomRec: document.getElementById('itUomRec').value.trim(),
    packRec: parseFloat(document.getElementById('itPackRec').value) || 1,
    uomCount: document.getElementById('itUomCount').value.trim(),
    packCount: parseFloat(document.getElementById('itPackCount').value) || 1,
    subFields: cat === 'NONFRESH' ? ['QTY'] : ['FV','BUT','FISH'],
    no: isEdit ? ITEM_MAP[cat][code].no : (ITEMS_BY_CAT[cat].length ? Math.max(...ITEMS_BY_CAT[cat].map(i=>i.no||0))+1 : 1)
  };

  try{
    await dbSet(`items/${cat}/${code}`, newItem);
    await dbPush('logs', {
      ts: Date.now(), date: todayStr(), store: '-', storeName: '-',
      user: SESSION.username, action: isEdit ? 'ITEM_EDIT' : 'ITEM_ADD',
      category: cat, changes: [{ itemCode: code, itemDesc: desc, field: '-', oldVal: '-', newVal: '-' }]
    });

    if(isEdit){
      const idx = ITEMS_BY_CAT[cat].findIndex(i=>i.code===code);
      ITEMS_BY_CAT[cat][idx] = newItem;
    } else {
      ITEMS_BY_CAT[cat].push(newItem);
    }
    ITEM_MAP[cat][code] = newItem;

    closeModal();
    toast(isEdit ? 'แก้ไขรายการเรียบร้อย' : 'เพิ่มรายการเรียบร้อย', 'success');
    await renderAdminItems();
  }catch(err){
    console.error(err);
    toast('เกิดข้อผิดพลาด: ' + err.message, 'error');
  }
}

function confirmDeleteItem(code){
  const cat = ADMIN_ITEMS_CAT;
  const item = ITEM_MAP[cat][code];
  showModal(`
    <h3>ยืนยันการลบรายการสินค้า</h3>
    <div class="text-soft" style="font-size:13px;margin-bottom:14px;">
      ต้องการลบ <b>${escapeHtml(item.desc)}</b> (${escapeHtml(code)}) ออกจากรายการสินค้าใช่หรือไม่?<br><br>
      <span class="pill pill-warning">หมายเหตุ: ข้อมูลตรวจนับที่บันทึกไปแล้วในอดีตจะยังคงอยู่ แต่จะไม่แสดงในหน้าบันทึกของสาขาอีกต่อไป</span>
    </div>
    <div class="modal-actions">
      <button class="btn btn-secondary" id="diCancelBtn">ยกเลิก</button>
      <button class="btn btn-danger" id="diConfirmBtn">ลบรายการ</button>
    </div>
  `, ()=>{
    document.getElementById('diCancelBtn').addEventListener('click', closeModal);
    document.getElementById('diConfirmBtn').addEventListener('click', ()=> deleteItem(code));
  });
}

async function deleteItem(code){
  const cat = ADMIN_ITEMS_CAT;
  const item = ITEM_MAP[cat][code];
  try{
    await dbRemove(`items/${cat}/${code}`);
    await dbPush('logs', {
      ts: Date.now(), date: todayStr(), store: '-', storeName: '-',
      user: SESSION.username, action: 'ITEM_DELETE',
      category: cat, changes: [{ itemCode: code, itemDesc: item.desc, field: '-', oldVal: '-', newVal: '-' }]
    });

    ITEMS_BY_CAT[cat] = ITEMS_BY_CAT[cat].filter(i=>i.code!==code);
    delete ITEM_MAP[cat][code];

    closeModal();
    toast('ลบรายการเรียบร้อย', 'success');
    await renderAdminItems();
  }catch(err){
    console.error(err);
    toast('เกิดข้อผิดพลาด: ' + err.message, 'error');
  }
}

/* ============================================================
   LOGS
============================================================ */
async function renderAdminLogs(){
  setTopbar('Log การใช้งาน', 'ประวัติการบันทึก / แก้ไข / ลบ ข้อมูล');
  const content = document.getElementById('content');

  const storeOptions = STORES_DATA
    .slice().sort((a,b)=> Number(a.locNo) - Number(b.locNo))
    .map(s=>`<option value="${s.username}">${s.locNo} - ${escapeHtml(s.name)}</option>`).join('');

  content.innerHTML = `
    <div class="card">
      <div class="card-head">
        <div>
          <div class="card-title">Log การใช้งานระบบ</div>
          <div class="muted">แสดง Log ล่าสุดตามจำนวนที่เลือก สามารถกรองตามสาขาและประเภทการทำงาน</div>
        </div>
      </div>
      <div class="form-row">
        <div class="form-group" style="min-width:220px;">
          <label>สาขา</label>
          <select id="logStore">
            <option value="ALL">-- ทุกสาขา --</option>
            <option value="-">-- ระบบ / Admin (รายการสินค้า) --</option>
            ${storeOptions}
          </select>
        </div>
        <div class="form-group">
          <label>ประเภทการทำงาน</label>
          <select id="logAction">
            <option value="ALL">ทั้งหมด</option>
            <option value="SAVE">SAVE (สาขาบันทึก)</option>
            <option value="UPDATE">UPDATE (แก้ไขโดย Admin)</option>
            <option value="DELETE">DELETE (ลบโดย Admin)</option>
            <option value="ITEM_ADD">ITEM_ADD (เพิ่มสินค้า)</option>
            <option value="ITEM_EDIT">ITEM_EDIT (แก้ไขสินค้า)</option>
            <option value="ITEM_DELETE">ITEM_DELETE (ลบสินค้า)</option>
            <option value="BULK_DELETE">BULK_DELETE (ล้างข้อมูลทั้งเดือน)</option>
          </select>
        </div>
        <div class="form-group">
          <label>จำนวน Log ล่าสุด</label>
          <select id="logLimit">
            <option value="200">200 รายการ</option>
            <option value="500">500 รายการ</option>
            <option value="1000">1,000 รายการ</option>
          </select>
        </div>
        <button class="btn btn-primary" id="logSearchBtn">โหลด Log</button>
      </div>
      <div class="mt-12" id="logResultInfo"></div>
      <div class="table-wrap mt-12" id="logResultWrap"></div>
    </div>
  `;

  document.getElementById('logSearchBtn').addEventListener('click', loadLogs);
  await loadLogs();
}

async function loadLogs(){
  const storeFilter = document.getElementById('logStore').value;
  const actionFilter = document.getElementById('logAction').value;
  const limit = parseInt(document.getElementById('logLimit').value, 10) || 200;
  const infoEl = document.getElementById('logResultInfo');
  const wrapEl = document.getElementById('logResultWrap');

  infoEl.innerHTML = '<div class="text-soft">กำลังโหลด...</div>';
  wrapEl.innerHTML = '';

  let data;
  try{
    const snap = await db.ref('logs').orderByChild('ts').limitToLast(limit).once('value');
    data = snap.val() || {};
  }catch(err){
    console.error(err);
    infoEl.innerHTML = `<div class="pill pill-danger">ไม่สามารถโหลด Log ได้ (ตรวจสอบว่าได้ตั้งค่า .indexOn: ["ts"] ใน Database Rules แล้ว) — ${err.message}</div>`;
    return;
  }

  let entries = Object.entries(data).map(([id, v])=>({ id, ...v }));
  if(storeFilter !== 'ALL') entries = entries.filter(e=> e.store === storeFilter);
  if(actionFilter !== 'ALL') entries = entries.filter(e=> e.action === actionFilter);
  entries.sort((a,b)=> (b.ts||0) - (a.ts||0));

  if(entries.length === 0){
    infoEl.innerHTML = '<div class="text-soft">ไม่พบ Log ตามเงื่อนไขที่เลือก</div>';
    return;
  }

  infoEl.innerHTML = `พบ <b class="num">${fmtNum(entries.length)}</b> รายการ`;

  const rows = entries.map(e=>{
    const changeCount = (e.changes||[]).length;
    const detailId = 'log_' + e.id;
    const changesRows = (e.changes||[]).map(c=>`
      <tr>
        <td class="desc"><div>${escapeHtml(c.itemDesc||'')}</div><div class="code">${escapeHtml(c.itemCode||'')}</div></td>
        <td>${escapeHtml(CAT_FIELD_LABELS[c.field] || c.field || '-')}</td>
        <td class="text-right num">${c.oldVal === '-' ? '-' : fmtNum(c.oldVal,2)}</td>
        <td class="text-right num">${c.newVal === '-' ? '-' : fmtNum(c.newVal,2)}</td>
      </tr>
    `).join('');

    return `
      <tr>
        <td class="num nowrap">${fmtDateTime(e.ts)}</td>
        <td class="nowrap">${e.store==='-' ? '-' : escapeHtml(storeLabel(e.store))}</td>
        <td>${escapeHtml(e.user||'-')}</td>
        <td><span class="pill ${logActionPillClass(e.action)}">${e.action}</span></td>
        <td>${e.category ? `<span class="pill pill-muted">${CAT_LABELS[e.category]||e.category}</span>` : '-'}</td>
        <td class="text-center">
          ${changeCount ? `<button class="btn btn-ghost btn-sm" data-toggle="${detailId}">${changeCount} รายการ ▾</button>` : '-'}
        </td>
      </tr>
      ${changeCount ? `
      <tr class="hidden" id="${detailId}">
        <td colspan="6" style="padding:0;">
          <table class="dtable" style="margin:0;">
            <thead><tr><th>รายการสินค้า</th><th>ฟิลด์</th><th class="text-right">ค่าเดิม</th><th class="text-right">ค่าใหม่</th></tr></thead>
            <tbody>${changesRows}</tbody>
          </table>
        </td>
      </tr>` : ''}
    `;
  }).join('');

  wrapEl.innerHTML = `
    <table class="dtable">
      <thead><tr>
        <th>เวลา</th><th>สาขา</th><th>ผู้ใช้งาน</th><th>การทำงาน</th><th>หมวดหมู่</th><th class="text-center">รายละเอียด</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  wrapEl.querySelectorAll('[data-toggle]').forEach(b=>{
    b.addEventListener('click', ()=>{
      document.getElementById(b.dataset.toggle).classList.toggle('hidden');
    });
  });
}

function logActionPillClass(action){
  switch(action){
    case 'SAVE': return 'pill-success';
    case 'UPDATE': return 'pill-info';
    case 'DELETE': return 'pill-danger';
    case 'ITEM_ADD': return 'pill-success';
    case 'ITEM_EDIT': return 'pill-info';
    case 'ITEM_DELETE': return 'pill-danger';
    case 'BULK_DELETE': return 'pill-danger';
    default: return 'pill-muted';
  }
}

/* ============================================================
   EXPORT (ALL BRANCHES)
============================================================ */
async function renderAdminExport(){
  setTopbar('Export ข้อมูล', 'ส่งออกข้อมูลการตรวจนับเป็นไฟล์ Excel');
  const content = document.getElementById('content');
  const today = todayStr();

  const storeChecks = STORES_DATA
    .slice().sort((a,b)=> Number(a.locNo) - Number(b.locNo))
    .map(s=>`
      <label class="flex items-center gap-8" style="padding:4px 0;font-weight:500;font-size:13px;cursor:pointer;">
        <input type="checkbox" class="exp-store-chk" value="${s.username}" checked>
        ${s.locNo} - ${escapeHtml(s.name)}
      </label>
    `).join('');

  content.innerHTML = `
    <div class="card">
      <div class="card-head">
        <div>
          <div class="card-title">Export ข้อมูลการตรวจนับ</div>
          <div class="muted">เลือกสาขา ช่วงวันที่ และหมวดหมู่ที่ต้องการ Export เป็นไฟล์ Excel (.xlsx)</div>
        </div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>จากวันที่</label><input type="date" id="expFrom" value="${today}" max="${today}"></div>
        <div class="form-group"><label>ถึงวันที่</label><input type="date" id="expTo" value="${today}" max="${today}"></div>
        <div class="form-group">
          <label>หมวดหมู่</label>
          <select id="expCat">
            <option value="ALL">ทั้งหมด</option>
            <option value="FRESH">FRESH FOOD</option>
            <option value="TRANSFER">TRANSFER</option>
            <option value="NONFRESH">NON FRESH</option>
          </select>
        </div>
        <button class="btn btn-primary" id="expRunBtn">Export Excel</button>
      </div>

      <div class="mt-16">
        <div class="flex items-center justify-between" style="margin-bottom:8px;">
          <div class="card-title" style="font-size:13px;">เลือกสาขา (${STORES_DATA.length} สาขา)</div>
          <div class="flex gap-8">
            <button class="btn btn-secondary btn-sm" id="expSelectAll">เลือกทั้งหมด</button>
            <button class="btn btn-secondary btn-sm" id="expSelectNone">ไม่เลือกเลย</button>
          </div>
        </div>
        <input type="text" id="expStoreFilter" placeholder="ค้นหาสาขา..." style="width:100%;padding:9px 12px;border-radius:8px;border:1.5px solid var(--border);background:var(--surface-2);margin-bottom:8px;">
        <div id="expStoreList" style="max-height:260px;overflow:auto;border:1px solid var(--border-soft);border-radius:10px;padding:8px 12px;background:var(--surface-2);">
          ${storeChecks}
        </div>
      </div>

      <div class="mt-12 text-faint" style="font-size:12px;">
        การ Export อ่านข้อมูลจาก Firebase 1 ครั้งต่อวัน (ครอบคลุมทุกสาขาในการอ่านครั้งเดียว) จึงรองรับการ Export ข้อมูลจำนวนมากกว่า 200 สาขาได้อย่างรวดเร็ว
      </div>
    </div>
  `;

  document.getElementById('expSelectAll').addEventListener('click', ()=>{
    content.querySelectorAll('.exp-store-chk').forEach(c=> c.checked = true);
  });
  document.getElementById('expSelectNone').addEventListener('click', ()=>{
    content.querySelectorAll('.exp-store-chk').forEach(c=> c.checked = false);
  });
  document.getElementById('expStoreFilter').addEventListener('input', (e)=>{
    const q = e.target.value.trim().toLowerCase();
    content.querySelectorAll('#expStoreList label').forEach(label=>{
      label.style.display = label.textContent.toLowerCase().includes(q) ? '' : 'none';
    });
  });
  document.getElementById('expRunBtn').addEventListener('click', runAdminExport);
}

async function runAdminExport(){
  const from = document.getElementById('expFrom').value;
  const to = document.getElementById('expTo').value;
  const catFilter = document.getElementById('expCat').value;

  if(!from || !to || from > to){ toast('กรุณาเลือกช่วงวันที่ให้ถูกต้อง', 'error'); return; }
  const dates = dateRange(from, to);
  if(dates.length > 370){ toast('ช่วงวันที่กว้างเกินไป (เกิน 1 ปี)', 'error'); return; }

  const selectedStores = new Set(
    Array.from(document.querySelectorAll('.exp-store-chk:checked')).map(c=>c.value)
  );
  if(selectedStores.size === 0){ toast('กรุณาเลือกสาขาอย่างน้อย 1 สาขา', 'error'); return; }

  const btn = document.getElementById('expRunBtn');
  setLoading(btn, true, 'กำลัง Export...');

  try{
    // ---- รวบรวม detail rows ----
    const rows = [];
    for(const dateStr of dates){
      const data = await dbGetOnce(`counts/${dateStr}`);
      if(!data) continue;
      Object.entries(data).forEach(([storeCode, rec])=>{
        if(!selectedStores.has(storeCode)) return;
        ['FRESH','TRANSFER','NONFRESH'].forEach(cat=>{
          if(catFilter !== 'ALL' && catFilter !== cat) return;
          const catData = rec[cat];
          if(!catData) return;
          Object.entries(catData).forEach(([itemCode, r])=>{
            const item = ITEM_MAP[cat][itemCode] || placeholderItem(cat, itemCode);
            rows.push(buildExportRow(dateStr, storeCode, cat, item, r));
          });
        });
      });
    }

    if(rows.length === 0){ toast('ไม่พบข้อมูลตามเงื่อนไขที่เลือก', 'error'); return; }

    // ---- สร้าง Summary Sheet: รวมรายสาขา + หมวดหมู่ ----
    // key = storeCode + '|' + cat
    const summaryMap = {};
    rows.forEach(r=>{
      const key = r['รหัสผู้ใช้สาขา'] + '|' + r['หมวดหมู่'];
      if(!summaryMap[key]){
        summaryMap[key] = {
          'รหัสผู้ใช้สาขา':     r['รหัสผู้ใช้สาขา'],
          'เลขที่สาขา (Loc)':   r['เลขที่สาขา (Loc)'],
          'ชื่อสาขา':           r['ชื่อสาขา'],
          'หมวดหมู่':           r['หมวดหมู่'],
          'รวมจำนวนตรวจนับ':   0,
          'มูลค่ารวม (บาท)':   0
        };
      }
      summaryMap[key]['รวมจำนวนตรวจนับ'] += Number(r['รวมจำนวนตรวจนับ']) || 0;
      summaryMap[key]['มูลค่ารวม (บาท)'] += Number(r['มูลค่ารวม (บาท)']) || 0;
    });

    // เรียงตาม locNo แล้วตาม category
    const catOrder = { 'FRESH FOOD': 0, 'TRANSFER': 1, 'NON FRESH': 2 };
    const summaryRows = Object.values(summaryMap).sort((a,b)=>{
      const la = Number(a['เลขที่สาขา (Loc)']) || 0;
      const lb = Number(b['เลขที่สาขา (Loc)']) || 0;
      if(la !== lb) return la - lb;
      return (catOrder[a['หมวดหมู่']]||0) - (catOrder[b['หมวดหมู่']]||0);
    });

    // ปัดทศนิยม 2 ตำแหน่งให้ summary
    summaryRows.forEach(r=>{
      r['รวมจำนวนตรวจนับ'] = Math.round(r['รวมจำนวนตรวจนับ'] * 100) / 100;
      r['มูลค่ารวม (บาท)'] = Math.round(r['มูลค่ารวม (บาท)'] * 100) / 100;
    });

    // แถว Grand Total สำหรับ summary
    const grandQty = summaryRows.reduce((s,r)=> s + r['รวมจำนวนตรวจนับ'], 0);
    const grandAmt = summaryRows.reduce((s,r)=> s + r['มูลค่ารวม (บาท)'], 0);
    summaryRows.push({
      'รหัสผู้ใช้สาขา':   '',
      'เลขที่สาขา (Loc)': '',
      'ชื่อสาขา':         'รวมทั้งหมด',
      'หมวดหมู่':         '',
      'รวมจำนวนตรวจนับ': Math.round(grandQty * 100) / 100,
      'มูลค่ารวม (บาท)': Math.round(grandAmt * 100) / 100
    });

    const filename = `PackagingCount_AllBranches_${from}_to_${to}.xlsx`;
    exportRowsToExcel({
      'สรุปรายสาขา': summaryRows,
      'ข้อมูลตรวจนับ': rows
    }, filename);
    toast(`Export สำเร็จ (${fmtNum(rows.length)} รายการ · ${fmtNum(summaryRows.length - 1)} สาขา/หมวด)`, 'success');
  }catch(err){
    console.error(err);
    toast('เกิดข้อผิดพลาด: ' + err.message, 'error');
  }finally{
    setLoading(btn, false);
  }
}

/* ============================================================
   CLEAR ALL DATA (BULK DELETE BY MONTH)
============================================================ */
async function renderAdminClearAll(){
  setTopbar('ล้างข้อมูลทั้งหมด', 'ลบข้อมูลการตรวจนับของทุกสาขาในเดือนที่เลือก');
  const content = document.getElementById('content');
  const thisMonth = todayStr().slice(0,7);

  content.innerHTML = `
    <div class="card">
      <div class="card-head">
        <div>
          <div class="card-title">ล้างข้อมูลการตรวจนับทั้งหมด (ทุกสาขา)</div>
          <div class="muted">ใช้สำหรับล้างข้อมูลตรวจนับของ <b>ทุกสาขา (${STORES_DATA.length} สาขา)</b> ในเดือนที่เลือกออกทั้งหมด</div>
        </div>
      </div>

      <div class="form-row">
        <div class="form-group"><label>เดือนที่ต้องการล้างข้อมูล</label><input type="month" id="clearMonth" value="${thisMonth}"></div>
        <button class="btn btn-danger" id="clearMonthBtn">ลบข้อมูลทั้งหมดของเดือนนี้</button>
      </div>

      <div class="mt-16">
        <span class="pill pill-danger">⚠ คำเตือน</span>
        <div class="text-soft mt-8" style="font-size:13px; line-height:1.7;">
          การลบข้อมูลนี้จะลบ <b>ข้อมูลการตรวจนับ (FRESH FOOD / TRANSFER / NON FRESH)</b> ของ
          <b>ทุกสาขาทั้งหมด</b> ในเดือนที่เลือก ออกจากระบบอย่างถาวร และ<b>ไม่สามารถกู้คืนได้</b><br>
          &middot; รายการสินค้า (Item Master) จะไม่ได้รับผลกระทบ<br>
          &middot; Log การลบนี้จะถูกบันทึกไว้ในหน้า "Log การใช้งาน"<br>
          &middot; แนะนำให้ Export ข้อมูลของเดือนนั้นเก็บไว้ก่อน (ไปที่เมนู "Export ข้อมูล") หากต้องการสำรองข้อมูล
        </div>
      </div>
    </div>
  `;

  document.getElementById('clearMonthBtn').addEventListener('click', confirmClearMonth);
}

function confirmClearMonth(){
  const month = document.getElementById('clearMonth').value;
  if(!month){ toast('กรุณาเลือกเดือน', 'error'); return; }

  showModal(`
    <h3>ยืนยันการล้างข้อมูลทั้งหมด</h3>
    <div class="text-soft" style="font-size:13px;margin-bottom:14px;line-height:1.7;">
      คุณกำลังจะลบข้อมูลการตรวจนับของ <b>ทุกสาขา (${STORES_DATA.length} สาขา)</b>
      ในเดือน <b class="num">${month}</b> ออกทั้งหมด<br><br>
      การลบนี้ <b style="color:var(--danger)">ไม่สามารถย้อนกลับได้</b><br><br>
      กรุณาพิมพ์ <code class="num" style="background:var(--surface-2);padding:2px 6px;border-radius:4px;">${month}</code> ในช่องด้านล่างเพื่อยืนยัน
    </div>
    <div class="field"><input type="text" id="clearConfirmInput" placeholder="พิมพ์ ${month} เพื่อยืนยัน" class="num"></div>
    <div class="modal-actions">
      <button class="btn btn-secondary" id="clearCancelBtn">ยกเลิก</button>
      <button class="btn btn-danger" id="clearConfirmBtn">ลบข้อมูลทั้งหมดของเดือน ${month}</button>
    </div>
  `, ()=>{
    document.getElementById('clearCancelBtn').addEventListener('click', closeModal);
    document.getElementById('clearConfirmBtn').addEventListener('click', ()=>{
      const val = document.getElementById('clearConfirmInput').value.trim();
      if(val !== month){
        toast('ข้อความยืนยันไม่ถูกต้อง กรุณาพิมพ์ให้ตรงกับเดือนที่เลือก', 'error');
        return;
      }
      executeClearMonth(month);
    });
  });
}

async function executeClearMonth(month){
  const btn = document.getElementById('clearConfirmBtn');
  setLoading(btn, true, 'กำลังลบข้อมูล...');

  const days = daysInMonth(month);
  const dates = [];
  for(let d=1; d<=days; d++) dates.push(`${month}-${pad2(d)}`);

  try{
    for(const dateStr of dates){
      await dbRemove(`counts/${dateStr}`);
    }
    await dbPush('logs', {
      ts: Date.now(), date: month, store: '-', storeName: 'ทุกสาขา',
      user: SESSION.username, action: 'BULK_DELETE', category: 'ALL',
      changes: [{ itemCode:'-', itemDesc:`ล้างข้อมูลการตรวจนับทั้งหมดของเดือน ${month} (ทุกสาขา ${STORES_DATA.length} สาขา)`, field:'-', oldVal:'-', newVal:'-' }]
    });

    closeModal();
    toast(`ล้างข้อมูลของเดือน ${month} เรียบร้อยแล้ว (ทุกสาขา)`, 'success');
  }catch(err){
    console.error(err);
    toast('เกิดข้อผิดพลาด: ' + err.message, 'error');
    setLoading(btn, false);
  }
}

/* ===================== js/dashboard.js ===================== */
// ============================================================
// dashboard.js — แดชบอร์ดสรุปข้อมูล (ใช้ร่วมกันทั้งหน้าสาขาและ Admin)
// ============================================================

let DASH_MODE  = 'month';      // 'month' | 'range'
let DASH_MONTH = null;
let DASH_FROM  = null;
let DASH_TO    = null;
let DASH_STORE = 'ALL';        // admin เท่านั้น

async function renderDashboardView(){
  const isAdmin = SESSION.role === 'admin';
  setTopbar('แดชบอร์ดสรุปข้อมูล', isAdmin ? 'สรุปมูลค่าการตรวจนับของทุกสาขา / สาขาที่เลือก' : `${SESSION.locNo} - ${SESSION.storeName}`);
  const content = document.getElementById('content');
  const thisMonth = todayStr().slice(0,7);

  const storeOptions = isAdmin ? STORES_DATA.slice()
    .sort((a,b)=> Number(a.locNo) - Number(b.locNo))
    .map(s=>`<option value="${s.username}" ${DASH_STORE===s.username?'selected':''}>${s.locNo} - ${escapeHtml(s.name)}</option>`).join('') : '';

  content.innerHTML = `
    <div class="card">
      <div class="card-head">
        <div>
          <div class="card-title">แดชบอร์ดสรุปข้อมูลการตรวจนับ</div>
          <div class="muted">${isAdmin ? 'เลือกสาขาและช่วงเวลาที่ต้องการดูสรุปมูลค่า Fresh Food' : 'สรุปมูลค่าการตรวจนับ Fresh Food ของสาขาคุณ'}</div>
        </div>
        <div class="tabs" id="dashModeTabs">
          <div class="tab ${DASH_MODE==='month'?'active':''}" data-mode="month">รายเดือน</div>
          <div class="tab ${DASH_MODE==='range'?'active':''}" data-mode="range">ช่วงวันที่</div>
        </div>
      </div>

      <div class="form-row mt-12">
        ${isAdmin ? `
        <div class="form-group" style="min-width:240px;">
          <label>สาขา</label>
          <select id="dashStore">
            <option value="ALL" ${DASH_STORE==='ALL'?'selected':''}>-- ทุกสาขา (รวม ${STORES_DATA.length} สาขา) --</option>
            ${storeOptions}
          </select>
        </div>` : ''}
        <div id="dashDateFilters" class="flex gap-12" style="flex-wrap:wrap;"></div>
        <button class="btn btn-primary" id="dashLoadBtn">แสดงข้อมูล</button>
      </div>

      <div id="dashBody" class="mt-16">
        <div class="text-soft" style="padding:40px;text-align:center;">กำลังโหลดข้อมูล...</div>
      </div>
    </div>
  `;

  renderDashDateFilters();

  content.querySelectorAll('#dashModeTabs .tab').forEach(t=>{
    t.addEventListener('click', ()=>{
      DASH_MODE = t.dataset.mode;
      content.querySelectorAll('#dashModeTabs .tab').forEach(x=> x.classList.toggle('active', x===t));
      renderDashDateFilters();
    });
  });

  document.getElementById('dashLoadBtn').addEventListener('click', loadDashboard);

  await loadDashboard();
}

function renderDashDateFilters(){
  const el = document.getElementById('dashDateFilters');
  const thisMonth = todayStr().slice(0,7);
  const today = todayStr();
  if(DASH_MODE === 'month'){
    el.innerHTML = `<div class="form-group"><label>เดือน</label><input type="month" id="dashMonth" value="${DASH_MONTH || thisMonth}" max="${thisMonth}"></div>`;
  } else {
    el.innerHTML = `
      <div class="form-group"><label>จากวันที่</label><input type="date" id="dashFrom" value="${DASH_FROM || today}" max="${today}"></div>
      <div class="form-group"><label>ถึงวันที่</label><input type="date" id="dashTo" value="${DASH_TO || today}" max="${today}"></div>
    `;
  }
}

async function loadDashboard(){
  const isAdmin = SESSION.role === 'admin';
  const bodyEl = document.getElementById('dashBody');
  bodyEl.innerHTML = '<div class="text-soft" style="padding:40px;text-align:center;">กำลังโหลดข้อมูล...</div>';

  let dates;
  if(DASH_MODE === 'month'){
    DASH_MONTH = document.getElementById('dashMonth').value;
    if(!DASH_MONTH){ toast('กรุณาเลือกเดือน', 'error'); return; }
    const days = daysInMonth(DASH_MONTH);
    dates = dateRange(`${DASH_MONTH}-01`, `${DASH_MONTH}-${pad2(days)}`);
  } else {
    DASH_FROM = document.getElementById('dashFrom').value;
    DASH_TO = document.getElementById('dashTo').value;
    if(!DASH_FROM || !DASH_TO || DASH_FROM > DASH_TO){ toast('กรุณาเลือกช่วงวันที่ให้ถูกต้อง', 'error'); return; }
    dates = dateRange(DASH_FROM, DASH_TO);
  }
  if(dates.length === 0){ toast('ช่วงเวลาไม่ถูกต้อง', 'error'); return; }
  if(dates.length > 370){ toast('ช่วงเวลากว้างเกินไป (เกิน 1 ปี)', 'error'); return; }

  const storeFilter = isAdmin ? document.getElementById('dashStore').value : SESSION.storeCode;
  DASH_STORE = storeFilter;

  const result = await computeCategorySums(dates, storeFilter, isAdmin);
  renderDashboardResult(result, dates.length, isAdmin, storeFilter);
}

/**
 * รวมมูลค่าตามหมวดหมู่ (FRESH / TRANSFER / NONFRESH)
 * storeFilter: 'ALL' (admin เท่านั้น = ทุกสาขา) หรือ storeCode เฉพาะสาขา
 */
async function computeCategorySums(dates, storeFilter, isAdmin){
  const sums = { FRESH:0, TRANSFER:0, NONFRESH:0 };
  let daysWithData = 0;
  const storesSeen = new Set();

  for(const dateStr of dates){
    let dayHasData = false;

    if(isAdmin && storeFilter === 'ALL'){
      const data = await dbGetOnce(`counts/${dateStr}`);
      if(!data) continue;
      Object.entries(data).forEach(([storeCode, rec])=>{
        ['FRESH','TRANSFER','NONFRESH'].forEach(cat=>{
          const catData = rec[cat];
          if(!catData) return;
          dayHasData = true;
          storesSeen.add(storeCode);
          Object.entries(catData).forEach(([code, r])=>{
            const item = ITEM_MAP[cat][code] || placeholderItem(cat, code);
            sums[cat] += recordAmount(item, r);
          });
        });
      });
    } else {
      const storeCode = storeFilter;
      const data = await dbGetOnce(`counts/${dateStr}/${storeCode}`);
      if(!data) continue;
      ['FRESH','TRANSFER','NONFRESH'].forEach(cat=>{
        const catData = data[cat];
        if(!catData) return;
        dayHasData = true;
        storesSeen.add(storeCode);
        Object.entries(catData).forEach(([code, r])=>{
          const item = ITEM_MAP[cat][code] || placeholderItem(cat, code);
          sums[cat] += recordAmount(item, r);
        });
      });
    }

    if(dayHasData) daysWithData++;
  }

  return { sums, daysWithData, storeCount: storesSeen.size };
}

function renderDashboardResult(result, totalDays, isAdmin, storeFilter){
  const bodyEl = document.getElementById('dashBody');
  const { sums, daysWithData, storeCount } = result;
  // Show only FRESH value
  const freshValue = sums.FRESH;
  const avgPerDay = daysWithData > 0 ? freshValue / daysWithData : 0;

  if(freshValue === 0){
    bodyEl.innerHTML = `
      <div style="padding:56px 24px;text-align:center;">
        <div style="font-size:48px;margin-bottom:12px;">📦</div>
        <div style="font-size:15px;font-weight:700;color:var(--text-soft);">ไม่พบข้อมูลการตรวจนับในช่วงเวลาที่เลือก</div>
        <div style="font-size:12.5px;color:var(--text-faint);margin-top:6px;">ลองเปลี่ยนช่วงวันที่หรือเลือกสาขาอื่น</div>
      </div>`;
    return;
  }

  let extraInfo = '';
  let extraInfoShort = '';
  if(isAdmin && storeFilter === 'ALL'){
    extraInfo = `สาขาที่มีข้อมูล <b class="num">${fmtNum(storeCount)}</b> / ${fmtNum(STORES_DATA.length)} สาขา &nbsp;·&nbsp; วันที่มีข้อมูล <b class="num">${fmtNum(daysWithData)}</b> / ${fmtNum(totalDays)} วัน`;
    extraInfoShort = `${storeCount} สาขา`;
  } else {
    const label = isAdmin ? escapeHtml(storeLabel(storeFilter)) : `${SESSION.locNo} - ${SESSION.storeName}`;
    extraInfo = `สาขา <b>${label}</b> &nbsp;·&nbsp; วันที่มีข้อมูล <b class="num">${fmtNum(daysWithData)}</b> / ${fmtNum(totalDays)} วัน`;
    extraInfoShort = label;
  }

  const freshColor = CAT_COLORS.FRESH || '#1A9E6B';

  bodyEl.innerHTML = `
    <!-- Hero KPI Card -->
    <div class="dash-hero-wrap">
      <div class="dash-hero-card">
        <div class="dash-hero-icon">🥩</div>
        <div class="dash-hero-content">
          <div class="dash-hero-label">มูลค่า FRESH FOOD รวม</div>
          <div class="dash-hero-value num">${fmtMoney(freshValue)}</div>
          <div class="dash-hero-unit">บาท</div>
        </div>
        <div class="dash-hero-badge">FRESH FOOD</div>
      </div>
    </div>

    <!-- KPI Row -->
    <div class="dash-kpi-row">
      <div class="dash-kpi-card">
        <div class="dash-kpi-icon">📅</div>
        <div class="dash-kpi-body">
          <div class="dash-kpi-label">วันที่มีข้อมูล</div>
          <div class="dash-kpi-value num">${fmtNum(daysWithData)} <span class="dash-kpi-unit">/ ${fmtNum(totalDays)} วัน</span></div>
        </div>
      </div>
      <div class="dash-kpi-card">
        <div class="dash-kpi-icon">📊</div>
        <div class="dash-kpi-body">
          <div class="dash-kpi-label">เฉลี่ยต่อวัน</div>
          <div class="dash-kpi-value num">${fmtMoney(avgPerDay)} <span class="dash-kpi-unit">บาท/วัน</span></div>
        </div>
      </div>
      ${isAdmin && storeFilter === 'ALL' ? `
      <div class="dash-kpi-card">
        <div class="dash-kpi-icon">🏪</div>
        <div class="dash-kpi-body">
          <div class="dash-kpi-label">สาขาที่มีข้อมูล</div>
          <div class="dash-kpi-value num">${fmtNum(storeCount)} <span class="dash-kpi-unit">/ ${fmtNum(STORES_DATA.length)} สาขา</span></div>
        </div>
      </div>` : ''}
    </div>

    <!-- Progress Visual -->
    <div class="dash-progress-card">
      <div class="dash-progress-head">
        <div>
          <div class="dash-progress-title">ความครอบคลุมข้อมูล</div>
          <div class="dash-progress-sub">${extraInfo}</div>
        </div>
        <div class="dash-progress-pct num">${totalDays > 0 ? Math.round(daysWithData/totalDays*100) : 0}%</div>
      </div>
      <div class="dash-progress-bar-wrap">
        <div class="dash-progress-bar-track">
          <div class="dash-progress-bar-fill" style="width:${totalDays > 0 ? (daysWithData/totalDays*100).toFixed(1) : 0}%;background:${freshColor};"></div>
        </div>
      </div>
      <div class="dash-progress-labels">
        <span class="dash-progress-label-start">0 วัน</span>
        <span class="dash-progress-label-end">${fmtNum(totalDays)} วัน</span>
      </div>
    </div>
  `;
}

/* ===================== js/app.js ===================== */
// ============================================================
// app.js — Bootstrap, Sidebar, Navigation, Routing
// ============================================================

let ITEMS_BY_CAT = { FRESH: [], TRANSFER: [], NONFRESH: [] };
let ITEM_MAP = { FRESH: {}, TRANSFER: {}, NONFRESH: {} };
let CURRENT_VIEW = null;

const STORE_NAV = [
  { id:'dashboard', label:'แดชบอร์ด',            icon:'📊' },
  { id:'entry',   label:'บันทึกการตรวจนับ',     icon:'📝' },
  { id:'history', label:'ประวัติ / Export Excel', icon:'🗂️' },
];
const ADMIN_NAV = [
  { id:'dashboard', label:'แดชบอร์ด',            icon:'📊' },
  { id:'overview', label:'ภาพรวมระบบ',         icon:'📈' },
  { id:'data',     label:'ข้อมูลการตรวจนับ',     icon:'🧾' },
  { id:'items',    label:'รายการสินค้า',         icon:'📦' },
  { id:'logs',     label:'Log การใช้งาน',         icon:'📜' },
  { id:'export',   label:'Export ข้อมูล (ทุกสาขา)', icon:'📤' },
  { id:'clear',    label:'ล้างข้อมูลทั้งหมด',     icon:'🗑️' },
];

document.addEventListener('DOMContentLoaded', ()=>{
  // รอให้โหลดข้อมูลหลักจาก data.json (ดู firebase.js) เสร็จก่อน แล้วค่อยเริ่มแอป
  DATA_READY.then(()=>{
  initLoginForm();

  // Set Makro logo image (embedded as data URI in data.js)
  document.querySelectorAll('.brand-logo').forEach(img=>{
    img.src = MAKRO_LOGO_DATA_URI;
  });

  // Mobile sidebar toggle
  document.getElementById('menuToggle').addEventListener('click', ()=>{
    document.getElementById('sidebar').classList.add('open');
    document.getElementById('sidebarBackdrop').classList.add('show');
  });
  document.getElementById('sidebarBackdrop').addEventListener('click', closeSidebar);

  const restored = restoreSession();
  if(restored){
    SESSION = restored;
    startApp();
  }
  }); // ปิด DATA_READY.then(...)
});

function closeSidebar(){
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sidebarBackdrop').classList.remove('show');
}

/* ============================================================
   START APP
============================================================ */
async function startApp(){
  document.getElementById('loginScreen').classList.add('hidden');
  document.getElementById('app').classList.remove('hidden');

  // Sidebar header info
  if(SESSION.role === 'store'){
    document.getElementById('sbStoreName').textContent = `${SESSION.locNo} - ${SESSION.storeName}`;
    document.getElementById('sbRole').textContent = 'บัญชีสาขา (Store)';
  } else {
    document.getElementById('sbStoreName').textContent = SESSION.name || 'ผู้ดูแลระบบ';
    document.getElementById('sbRole').textContent = 'ผู้ดูแลระบบ (Admin)';
  }

  renderSidebar();

  toast('กำลังโหลดข้อมูลสินค้า...');
  await ensureItemsSeeded();
  await loadItemsFromDB();

  if(SESSION.role === 'store'){
    navigateTo('dashboard');
  } else {
    navigateTo('dashboard');
  }
}

/* ============================================================
   ITEM MASTER: seed + load
============================================================ */
async function ensureItemsSeeded(){
  const existing = await dbGetOnce('items');
  if(existing && Object.keys(existing).length > 0) return;

  // First run — seed item master data from embedded ITEMS_DATA
  const updates = {};
  let counters = { FRESH:0, TRANSFER:0, NONFRESH:0 };
  ITEMS_DATA.forEach(item=>{
    counters[item.category] += 1;
    updates[`items/${item.category}/${item.code}`] = {
      code: item.code,
      desc: item.desc,
      supplier: item.supplier,
      price: item.price,
      uomRec: item.uomRec,
      packRec: item.packRec,
      uomCount: item.uomCount,
      packCount: item.packCount,
      subFields: item.subFields,
      no: counters[item.category]
    };
  });
  await dbUpdate(updates);
}

async function loadItemsFromDB(){
  const data = await dbGetOnce('items') || {};
  ['FRESH','TRANSFER','NONFRESH'].forEach(cat=>{
    const obj = data[cat] || {};
    const arr = Object.values(obj);
    arr.sort((a,b)=> (a.no||0) - (b.no||0));
    ITEMS_BY_CAT[cat] = arr;
    ITEM_MAP[cat] = {};
    arr.forEach(it=> ITEM_MAP[cat][it.code] = it);
  });
}

/* ============================================================
   SIDEBAR / NAV
============================================================ */
function renderSidebar(){
  const nav = SESSION.role === 'store' ? STORE_NAV : ADMIN_NAV;
  const container = document.getElementById('sidebarNav');
  container.innerHTML = nav.map(item=>`
    <div class="nav-item" data-view="${item.id}">
      <span class="ico">${item.icon}</span>
      <span>${item.label}</span>
    </div>
  `).join('');
  container.querySelectorAll('.nav-item').forEach(el=>{
    el.addEventListener('click', ()=>{
      navigateTo(el.dataset.view);
      closeSidebar();
    });
  });
}

function setActiveNav(viewId){
  document.querySelectorAll('#sidebarNav .nav-item').forEach(el=>{
    el.classList.toggle('active', el.dataset.view === viewId);
  });
}

function setTopbar(title, sub){
  document.getElementById('topbarTitle').textContent = title;
  document.getElementById('topbarSub').textContent = sub || '';
}

/* ============================================================
   ROUTER
============================================================ */
function navigateTo(viewId){
  CURRENT_VIEW = viewId;
  setActiveNav(viewId);

  switch(viewId){
    // Shared
    case 'dashboard': return renderDashboardView();

    // Store views
    case 'entry':   return renderEntryView();
    case 'history': return renderHistoryView();

    // Admin views
    case 'overview': return renderAdminOverview();
    case 'data':     return renderAdminData();
    case 'items':    return renderAdminItems();
    case 'logs':     return renderAdminLogs();
    case 'export':   return renderAdminExport();
    case 'clear':    return renderAdminClearAll();

    default:
      document.getElementById('content').innerHTML = '<div class="card">ไม่พบหน้านี้</div>';
  }
}
