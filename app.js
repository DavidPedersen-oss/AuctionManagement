/* ===================================================================
   CSULB Auction Console — static, server-free.
   Runs on GitHub Pages or by opening index.html. Data in browser.
   =================================================================== */

const STATUSES = ["PREP","LIVE","SOLD","UNSOLD","RELISTED","PAID","PICKED_UP","CLOSED","NEEDS_REVIEW"];
const STATUS_LABEL = { PREP:"Prep", LIVE:"Live", SOLD:"Sold", UNSOLD:"Unsold", RELISTED:"Relisted",
  PAID:"Paid", PICKED_UP:"Picked up", CLOSED:"Closed", NEEDS_REVIEW:"Review" };
const STORE_KEY = "csulb_auction_console_v3";

/* Column definitions drive the header, sorting, filtering, and cells */
const COLUMNS = [
  { key:"survey",       label:"Survey #",     filter:"text",   sort:"text" },
  { key:"description",  label:"Item",         filter:"text",   sort:"text" },
  { key:"tag",          label:"Tag #",        filter:"text",   sort:"text" },
  { key:"dateAssigned", label:"Date Assigned",filter:"text",   sort:"date" },
  { key:"dept",         label:"Department",   filter:"select", sort:"text" },
  { key:"condition",    label:"Cond.",        filter:"select", sort:"text" },
  { key:"status",       label:"Status",       filter:"select", sort:"text" },
  { key:"platforms",    label:"Platforms",    filter:"select", sort:"none" },
  { key:"currentBid",   label:"Current Bid",  filter:"none",   sort:"num", align:"right" },
];

/* ---------- State ---------- */
let items = [];
let selected = new Set();
let globalTerm = "";
let colText = {};          // {colKey: substring}
let colSet = {};           // {colKey: Set(values) | null}  null = all
let sortKey = "survey", sortDir = 1;
let editingId = null, drawerDraft = null;
let pendingRows = null, pendingSheet = null;
let lastDeleted = null;
let popCol = null;         // column whose filter popup is open
let showArchived = false;  // archived items hidden from main view by default

/* ---------- Record ---------- */
function uid(){ return 'i'+Math.random().toString(36).slice(2,9)+Date.now().toString(36).slice(-4); }
function newPlat(){ return { listed:false, releases:0, url:"", price:"", currentBid:"", soldPrice:"" }; }
function blankItem(){ return { id:uid(), survey:"", description:"", serial:"", tag:"", assetId:"", year:"",
  amount:"", currentBid:"", dateAssigned:"", deptId:"", dept:"", drk:"", condition:"", status:"PREP", notes:"",
  archived:false, platforms:{ GD:newPlat(), PS:newPlat() } }; }
function migrate(it){
  it.platforms = it.platforms || {};
  it.platforms.GD = Object.assign(newPlat(), it.platforms.GD||{});
  it.platforms.PS = Object.assign(newPlat(), it.platforms.PS||{});
  if(it.dateAssigned==null) it.dateAssigned="";
  if(it.currentBid==null) it.currentBid="";
  if(typeof it.archived!=="boolean") it.archived=false;
  if(!it.status) it.status="PREP"; if(!it.id) it.id=uid(); return it;
}
/* The current bid shown on a row = highest live bid across listed platforms,
   falling back to the item-level currentBid the parser/scraper set. */
function rowCurrentBid(it){
  const vals=[it.currentBid, it.platforms.GD.currentBid, it.platforms.PS.currentBid]
    .map(v=>{ const n=parseFloat(String(v).replace(/[^0-9.\-]/g,"")); return isNaN(n)?null:n; })
    .filter(v=>v!=null);
  return vals.length?Math.max(...vals):"";
}

/* ---------- Persistence ---------- */
function save(){ try{ localStorage.setItem(STORE_KEY, JSON.stringify({items, savedAt:Date.now()})); }
  catch(e){ toast("Couldn't save to browser storage — export a backup.","warn"); } }
function load(){ try{ const r=localStorage.getItem(STORE_KEY); if(r){ items=(JSON.parse(r).items||[]).map(migrate); } }catch(e){ items=[]; } }

/* ===================================================================
   Excel / CSV parsing
   =================================================================== */
function fmtAmount(v){ if(v===""||v==null) return ""; const n=typeof v==="number"?v:parseFloat(String(v).replace(/[^0-9.\-]/g,"")); return isNaN(n)?"":n; }
function fmtDate(v){
  if(v===""||v==null) return "";
  // Excel serial date (number) -> ISO yyyy-mm-dd
  if(typeof v==="number" && v>20000 && v<80000){
    const d=new Date(Math.round((v-25569)*86400*1000));
    if(!isNaN(d)) return d.toISOString().slice(0,10);
  }
  const s=String(v).trim();
  const d=new Date(s);
  if(!isNaN(d) && /\d{4}/.test(s)) return d.toISOString().slice(0,10);
  return s; // leave as-is if unparseable
}
function normH(h){ return String(h||"").replace(/\s+/g," ").trim().toLowerCase(); }
const HEADER_MAP = [
  [["survey #","survey#","survey number","survey no","survey"],"survey"],
  [["description","item description","item"],"description"],
  [["serial#","serial #","serial number","serial","s/n","sn"],"serial"],
  [["tag #","tag#","tag","property tag","property tag #"],"tag"],
  [["date assigned","assigned date","date surveyed","survey date","date"],"dateAssigned"],
  [["asset id","assetid","asset"],"assetId"],
  [["in service year","service year","acquisition year","year"],"year"],
  [["original amount","original value","est market value","est. market value","value","amount"],"amount"],
  [["dept id","department id","deptid"],"deptId"],
  [["dept name","department name","department","dept"],"dept"],
  [["drk","record keeper","department record keeper"],"drk"],
  [["disposal condition","present condition code","condition","present condition"],"condition"],
  [["disposal action","action","disposition"],"action"],
  [["notes","note","comments","additional comments"],"notes"],
];
function headerIndex(hr){ const idx={}; hr.forEach((h,i)=>{ const n=normH(h); for(const[al,f]of HEADER_MAP){ if(al.includes(n)&&!(f in idx)) idx[f]=i; } }); return idx; }
function findHeaderRow(aoa){ for(let r=0;r<Math.min(aoa.length,8);r++){ const c=aoa[r].map(normH); if(c.some(x=>/survey/.test(x))&&c.some(x=>/description|item/.test(x))) return r; } return 0; }

function parseWorkbook(buf){
  const wb = XLSX.read(buf,{type:"array"});
  const order = wb.SheetNames.slice().sort((a,b)=>(/survey/i.test(a)?0:1)-(/survey/i.test(b)?0:1));
  for(const sn of order){
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:""});
    if(!aoa.length) continue;
    const hr=findHeaderRow(aoa), idx=headerIndex(aoa[hr]);
    if(idx.survey==null&&idx.description==null) continue;
    const rows=[];
    for(let r=hr+1;r<aoa.length;r++){
      const row=aoa[r], get=f=>idx[f]!=null?row[idx[f]]:"";
      const desc=String(get("description")||"").trim(), survey=String(get("survey")||"").trim();
      if(!desc&&!survey) continue;
      rows.push({ survey, description:desc, serial:String(get("serial")||"").trim(),
        tag:String(get("tag")||"").trim(), assetId:String(get("assetId")||"").trim(),
        dateAssigned:fmtDate(get("dateAssigned")),
        year:String(get("year")||"").trim(), amount:fmtAmount(get("amount")),
        deptId:String(get("deptId")||"").trim(), dept:String(get("dept")||"").trim(),
        drk:String(get("drk")||"").trim(), condition:String(get("condition")||"").trim(),
        action:String(get("action")||"").trim().toUpperCase(), notes:String(get("notes")||"").trim() });
    }
    if(rows.length) return { sheet:sn, rows };
  }
  return { sheet:null, rows:[] };
}
function rowsToItems(rows,{auctionOnly,group}){
  let src = auctionOnly ? rows.filter(r=>r.action.includes("AUCTION")) : rows;
  if(!group) return src.map(r=>Object.assign(blankItem(),{survey:r.survey,description:r.description,serial:r.serial,tag:r.tag,assetId:r.assetId,dateAssigned:r.dateAssigned,year:r.year,amount:r.amount,deptId:r.deptId,dept:r.dept,drk:r.drk,condition:r.condition,notes:r.notes}));
  const bySurvey=new Map(), loose=[];
  for(const r of src){ if(r.survey){ if(!bySurvey.has(r.survey)) bySurvey.set(r.survey,[]); bySurvey.get(r.survey).push(r); } else loose.push(r); }
  const out=[];
  for(const [survey,grp] of bySurvey){
    const first=grp[0]; let description=first.description, notes=first.notes;
    if(grp.length>1){ description=`${first.description}  (+${grp.length-1} more)`;
      const extra=grp.slice(1).map(g=>"• "+g.description).join("\n");
      notes=(notes?notes+"\n\n":"")+`Line items in survey ${survey}:\n• ${first.description}\n${extra}`; }
    const best=grp.reduce((a,b)=>(Number(b.amount)||0)>(Number(a.amount)||0)?b:a,first);
    out.push(Object.assign(blankItem(),{survey,description,serial:first.serial,tag:first.tag||best.tag,assetId:first.assetId,dateAssigned:first.dateAssigned,year:first.year,amount:best.amount,deptId:first.deptId,dept:first.dept,drk:first.drk,condition:first.condition,notes}));
  }
  for(const r of loose) out.push(Object.assign(blankItem(),{description:r.description,serial:r.serial,tag:r.tag,assetId:r.assetId,dateAssigned:r.dateAssigned,year:r.year,amount:r.amount,deptId:r.deptId,dept:r.dept,drk:r.drk,condition:r.condition,notes:r.notes}));
  return out;
}
function importItems(newItems,{merge}){
  if(!merge){ items=newItems; return newItems.length; }
  const existing=new Set(items.map(i=>(i.survey||"").toUpperCase()+"|"+i.description.toLowerCase()));
  let added=0;
  for(const ni of newItems){ const k=(ni.survey||"").toUpperCase()+"|"+ni.description.toLowerCase(); if(!existing.has(k)){ items.push(ni); existing.add(k); added++; } }
  return added;
}

/* ===================================================================
   Cell value accessor (used by filters + sorting + rendering)
   =================================================================== */
function cellValue(it,key){
  if(key==="platforms"){
    const a=[]; if(it.platforms.GD.listed) a.push("GovDeals"); if(it.platforms.PS.listed) a.push("Public Surplus");
    return a.length?a.join(", "):"Not listed";
  }
  if(key==="status") return STATUS_LABEL[it.status]||it.status;
  if(key==="currentBid"){ const b=rowCurrentBid(it); return b===""?"":String(b); }
  return it[key]==null?"":String(it[key]);
}
/* distinct values for a select-filter column */
function distinctValues(key){
  const m=new Map();
  for(const it of items){
    let vals;
    if(key==="platforms"){ vals=[]; if(it.platforms.GD.listed) vals.push("GovDeals"); if(it.platforms.PS.listed) vals.push("Public Surplus"); if(!vals.length) vals=["Not listed"]; }
    else vals=[cellValue(it,key)||"(blank)"];
    for(const v of vals) m.set(v,(m.get(v)||0)+1);
  }
  return [...m.entries()].sort((a,b)=>a[0].localeCompare(b[0],undefined,{numeric:true}));
}

/* ===================================================================
   Filtering / sorting
   =================================================================== */
function matchesGlobal(it,term){
  if(!term) return true;
  const hay=[it.survey,it.description,it.tag,it.serial,it.assetId,it.dateAssigned,it.dept,it.deptId,it.notes,it.platforms.GD.url,it.platforms.PS.url].join(" ").toLowerCase();
  return term.toLowerCase().split(/\s+/).every(t=>hay.includes(t));
}
function passesColumnFilters(it){
  for(const c of COLUMNS){
    // text filter
    const t=colText[c.key];
    if(t){ if(!cellValue(it,c.key).toLowerCase().includes(t.toLowerCase())) return false; }
    // set filter
    const set=colSet[c.key];
    if(set){
      if(c.key==="platforms"){
        const vals=[]; if(it.platforms.GD.listed) vals.push("GovDeals"); if(it.platforms.PS.listed) vals.push("Public Surplus"); if(!vals.length) vals.push("Not listed");
        if(!vals.some(v=>set.has(v))) return false;
      } else {
        const v=cellValue(it,c.key)||"(blank)";
        if(!set.has(v)) return false;
      }
    }
  }
  return true;
}
function visibleItems(){
  let v=items.filter(it=>(showArchived||!it.archived)&&matchesGlobal(it,globalTerm)&&passesColumnFilters(it));
  const col=COLUMNS.find(c=>c.key===sortKey);
  v.sort((a,b)=>{
    if(col&&col.sort==="num"){
      const an=Number(cellValue(a,sortKey))||0, bn=Number(cellValue(b,sortKey))||0; return (an-bn)*sortDir;
    }
    if(col&&col.sort==="date"){
      const ad=Date.parse(a[sortKey])||0, bd=Date.parse(b[sortKey])||0; return (ad-bd)*sortDir;
    }
    const av=cellValue(a,sortKey).toLowerCase(), bv=cellValue(b,sortKey).toLowerCase();
    return av<bv?-1*sortDir:av>bv?1*sortDir:0;
  });
  return v;
}

/* ===================================================================
   Rendering
   =================================================================== */
function money(v){ const n=Number(v); if(!v||isNaN(n)) return '<span class="none">—</span>'; return "$"+n.toLocaleString(undefined,{maximumFractionDigits:0}); }
function esc(s){ return String(s==null?"":s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }

function buildHead(){
  const tr=document.getElementById("headRow");
  // keep checkbox th (first child), remove rest
  while(tr.children.length>1) tr.removeChild(tr.lastChild);
  for(const c of COLUMNS){
    const th=document.createElement("th");
    if(c.align==="right") th.style.textAlign="right";
    const sorted = sortKey===c.key;
    let filterHtml="";
    if(c.filter==="text"){
      filterHtml=`<div class="th-filter"><input type="text" data-coltext="${c.key}" placeholder="filter…" value="${esc(colText[c.key]||"")}"></div>`;
    } else if(c.filter==="select"){
      const active = colSet[c.key] && colSet[c.key].size>0;
      const lbl = active ? `${colSet[c.key].size} selected` : "All";
      filterHtml=`<div class="th-filter"><button class="th-filterbtn ${active?'active':''}" data-colfilter="${c.key}">${lbl}<span class="caret">▼</span></button></div>`;
    }
    th.innerHTML=`<div class="th-pad">
      <div class="th-label ${sorted?'sorted':''}" data-sort="${c.key}">${esc(c.label)}<span class="sortarr">${sorted?(sortDir===1?'▲':'▼'):'▲'}</span></div>
      ${filterHtml}</div>`;
    tr.appendChild(th);
  }
  // actions column
  const tha=document.createElement("th"); tha.className="th-actions"; tha.innerHTML=`<div class="th-pad"></div>`;
  tr.appendChild(tha);
}

function statusSelect(it){
  const opts=STATUSES.map(s=>`<option value="${s}" ${s===it.status?"selected":""}>${STATUS_LABEL[s]}</option>`).join("");
  return `<span class="statwrap"><select class="statsel s-${it.status}" data-statusfor="${it.id}">${opts}</select></span>`;
}
function platCell(it){
  let h="";
  for(const p of ["GD","PS"]){ const pl=it.platforms[p]; if(pl.listed){ h+=`<span class="plat ${p}">${p}${pl.releases>0?`<span class="x">×${pl.releases}</span>`:""}</span>`; } }
  return h?`<span class="platcell">${h}</span>`:'<span class="td-amt"><span class="none">—</span></span>';
}
function rowHtml(it){
  const sel=selected.has(it.id);
  const cells = COLUMNS.map(c=>{
    switch(c.key){
      case "survey": return `<td class="td-survey">${it.survey?esc(it.survey):'<span class="nosurvey">—</span>'}</td>`;
      case "description": return `<td class="td-desc"><span class="d">${it.archived?'<span class="archbadge">Archived</span> ':''}${esc(it.description)}</span>${it.serial?`<span class="sub">S/N ${esc(it.serial)}</span>`:""}</td>`;
      case "tag": return `<td class="td-mono">${esc(it.tag)||"—"}</td>`;
      case "dateAssigned": return `<td class="td-mono">${esc(it.dateAssigned)||"—"}</td>`;
      case "dept": return `<td class="td-mono" title="${esc(it.dept)}">${esc(it.deptId)?esc(it.deptId)+" · ":""}${esc(it.dept)||"—"}</td>`;
      case "condition": return `<td class="td-mono">${esc(it.condition)||"—"}</td>`;
      case "status": return `<td>${statusSelect(it)}</td>`;
      case "platforms": return `<td>${platCell(it)}</td>`;
      case "currentBid": return `<td class="td-amt">${money(rowCurrentBid(it))}</td>`;
      default: return `<td>${esc(cellValue(it,c.key))}</td>`;
    }
  }).join("");
  return `<tr data-id="${it.id}" class="${sel?'sel':''} ${it.archived?'arch':''}">
    <td class="td-chk"><input type="checkbox" class="rowchk" data-id="${it.id}" ${sel?'checked':''}></td>
    ${cells}
    <td class="td-actions"><button class="editlink" data-edit="${it.id}">Edit</button></td>
  </tr>`;
}

function render(){
  buildHead();
  const v=visibleItems();
  const tb=document.getElementById("tbody");
  const empty=document.getElementById("emptyState");
  tb.innerHTML = v.map(rowHtml).join("");
  empty.classList.toggle("hidden", items.length>0);

  const activeItems=items.filter(i=>!i.archived);
  document.getElementById("statTotal").textContent=activeItems.length;
  document.getElementById("statLive").textContent=activeItems.filter(i=>i.status==="LIVE").length;
  document.getElementById("statRelist").textContent=activeItems.filter(i=>i.status==="RELISTED").length;
  document.getElementById("statSold").textContent=activeItems.filter(i=>["SOLD","PAID","PICKED_UP"].includes(i.status)).length;
  document.getElementById("statReview").textContent=activeItems.filter(i=>i.status==="NEEDS_REVIEW").length;
  const archivedCount=items.length-activeItems.length;
  const at=document.getElementById("archiveToggle");
  if(at){ at.textContent=archivedCount?(showArchived?`Hide archived (${archivedCount})`:`Show archived (${archivedCount})`):""; at.style.display=archivedCount?"inline-flex":"none"; }
  const au=document.getElementById("bulkUnarchive");
  if(au) au.style.display=(showArchived&&selected.size)?"inline-flex":"none";

  renderActiveFilters();
  updateBulkbar();
  const chkAll=document.getElementById("chkAll");
  const allSel=v.length>0&&v.every(it=>selected.has(it.id));
  chkAll.checked=allSel; chkAll.indeterminate=!allSel&&v.some(it=>selected.has(it.id));
}

function renderActiveFilters(){
  const wrap=document.getElementById("activeFilters");
  const pills=[];
  for(const c of COLUMNS){
    if(colText[c.key]) pills.push({key:c.key,label:c.label,val:`"${colText[c.key]}"`,kind:"text"});
    if(colSet[c.key]&&colSet[c.key].size>0){ const n=colSet[c.key].size; pills.push({key:c.key,label:c.label,val:n===1?[...colSet[c.key]][0]:`${n} values`,kind:"set"}); }
  }
  if(!pills.length){ wrap.innerHTML=""; return; }
  wrap.innerHTML = pills.map(p=>`<span class="fpill"><span class="k">${esc(p.label)}</span>${esc(p.val)}<button data-clearcol="${p.key}" data-kind="${p.kind}">×</button></span>`).join("")
    + `<button class="clearall" id="clearAllFilters">Clear all filters</button>`;
}

/* ===================================================================
   Column filter popup (Excel-style autofilter)
   =================================================================== */
let popValues=[], popChecked=new Set();
function openFilterPop(colKey, btnEl){
  popCol=colKey;
  popValues=distinctValues(colKey);
  const current=colSet[colKey];
  popChecked = current ? new Set(current) : new Set(popValues.map(v=>v[0]));
  document.getElementById("fpSearch").value="";
  renderPopList("");
  const pop=document.getElementById("filterPop");
  pop.classList.add("show");
  const r=btnEl.getBoundingClientRect();
  let left=r.left, top=r.bottom+4;
  const pw=250;
  if(left+pw>window.innerWidth-10) left=window.innerWidth-pw-10;
  pop.style.left=left+"px"; pop.style.top=top+"px";
  setTimeout(()=>document.getElementById("fpSearch").focus(),30);
}
function renderPopList(search){
  const list=document.getElementById("fpList");
  const s=search.toLowerCase();
  list.innerHTML=popValues.filter(([v])=>v.toLowerCase().includes(s)).map(([v,c])=>
    `<label><input type="checkbox" class="fpchk" data-v="${esc(v)}" ${popChecked.has(v)?"checked":""}><span class="v">${esc(v)}</span><span class="c">${c}</span></label>`
  ).join("") || `<div style="padding:14px;text-align:center;color:var(--ink-soft);font-size:12px">No matches</div>`;
}
function closeFilterPop(){ document.getElementById("filterPop").classList.remove("show"); popCol=null; }
function applyFilterPop(){
  if(!popCol) return;
  const all=popValues.length;
  if(popChecked.size===0){ colSet[popCol]=new Set(); }          // nothing → show nothing
  else if(popChecked.size===all){ delete colSet[popCol]; }      // all → no filter
  else colSet[popCol]=new Set(popChecked);
  closeFilterPop(); render();
}

/* ===================================================================
   Bulk operations
   =================================================================== */
function updateBulkbar(){
  const bar=document.getElementById("bulkbar");
  if(selected.size>0){ bar.classList.add("show"); document.getElementById("bulkCount").textContent=selected.size; }
  else bar.classList.remove("show");
}
function applyBulk(){
  const st=document.getElementById("bulkStatus").value, plat=document.getElementById("bulkPlatform").value;
  if(!st&&!plat){ toast("Choose a status or platform to apply."); return; }
  let n=0;
  items.forEach(it=>{ if(!selected.has(it.id)) return;
    if(st) it.status=st;
    if(plat){ if(plat==="both"){ it.platforms.GD.listed=true; it.platforms.PS.listed=true; } else it.platforms[plat].listed=true; }
    n++; });
  save(); render();
  document.getElementById("bulkStatus").value=""; document.getElementById("bulkPlatform").value="";
  toast(`Updated ${n} item${n>1?"s":""}.`);
}
function bulkRelease(){
  let n=0;
  items.forEach(it=>{ if(!selected.has(it.id)) return;
    ["GD","PS"].forEach(p=>{ if(it.platforms[p].listed){ it.platforms[p].releases++; n++; } });
    if(it.status==="UNSOLD") it.status="RELISTED"; });
  save(); render();
  toast(n?`Logged a release on ${n} listing${n>1?"s":""}.`:"Selected items aren't listed on a platform yet.", n?"":"warn");
}
function bulkDelete(){
  const del=items.filter(it=>selected.has(it.id)); if(!del.length) return;
  if(!confirm(`Delete ${del.length} selected item${del.length>1?"s":""}?`)) return;
  lastDeleted=del.map(d=>({...d})); items=items.filter(it=>!selected.has(it.id)); selected.clear();
  save(); render(); toast(`Deleted ${del.length} item${del.length>1?"s":""}.`,"warn",true);
}
function bulkArchive(){
  if(!selected.size) return;
  let n=0;
  items.forEach(it=>{ if(selected.has(it.id)){ it.archived=true; n++; } });
  selected.clear(); save(); render(); toast(`Archived ${n} item${n>1?"s":""}.`);
}
function bulkUnarchive(){
  if(!selected.size) return;
  let n=0;
  items.forEach(it=>{ if(selected.has(it.id)){ it.archived=false; n++; } });
  save(); render(); toast(`Restored ${n} item${n>1?"s":""} from archive.`);
}

/* ===================================================================
   Detail drawer
   =================================================================== */
function openDrawer(id){
  const it=items.find(i=>i.id===id); if(!it) return;
  editingId=id; drawerDraft=JSON.parse(JSON.stringify(it));
  document.getElementById("dSurvey").textContent=it.survey?"Survey #"+it.survey:"No survey number";
  document.getElementById("dTitle").textContent=it.description||"Untitled item";
  document.getElementById("drawerBody").innerHTML=drawerForm(drawerDraft);
  bindDrawer(); setDirty(false);
  const ab=document.getElementById("dArchive"); if(ab) ab.textContent=it.archived?"Unarchive":"Archive";
  document.getElementById("scrim").classList.add("show");
  document.getElementById("drawer").classList.add("show");
}
function closeDrawer(force){ if(!force&&drawerDirty()&&!confirm("Discard unsaved changes?")) return;
  editingId=null; drawerDraft=null;
  document.getElementById("scrim").classList.remove("show"); document.getElementById("drawer").classList.remove("show"); }
function drawerDirty(){ if(!editingId||!drawerDraft) return false; return JSON.stringify(items.find(i=>i.id===editingId))!==JSON.stringify(drawerDraft); }
function setDirty(){ document.getElementById("dDirty").classList.toggle("show", drawerDirty()); }
function statusOpts(sel){ return STATUSES.map(s=>`<option value="${s}" ${s===sel?"selected":""}>${STATUS_LABEL[s]}</option>`).join(""); }
function platBlock(p,pl){
  const name=p==="GD"?"GovDeals":"Public Surplus";
  return `<div class="platblock ${pl.listed?'on':''}" data-platblock="${p}">
    <div class="ph"><label class="sw"><input type="checkbox" data-f="platforms.${p}.listed" ${pl.listed?"checked":""}> <span class="plat ${p}">${p}</span> ${name}</label>
      ${pl.url?`<span class="url">${esc(pl.url)}</span>`:""}</div>
    <div class="frow">
      <div class="fgroup"><label>Releases</label><input type="number" min="0" data-f="platforms.${p}.releases" value="${esc(pl.releases)}"></div>
      <div class="fgroup"><label>List price</label><input data-f="platforms.${p}.price" value="${esc(pl.price)}" placeholder="$"></div>
      <div class="fgroup"><label>Current bid</label><input data-f="platforms.${p}.currentBid" value="${esc(pl.currentBid)}" placeholder="$"></div>
      <div class="fgroup"><label>Sold price</label><input data-f="platforms.${p}.soldPrice" value="${esc(pl.soldPrice)}" placeholder="$"></div>
    </div>
    <div class="fgroup" style="margin-bottom:0"><label>Listing URL</label><input data-f="platforms.${p}.url" value="${esc(pl.url)}" placeholder="https://…"></div>
  </div>`;
}
function drawerForm(it){
  return `<div class="dsection">Item <span class="bar"></span></div>
  <div class="frow"><div class="fgroup" style="flex:.6"><label>Survey #</label><input data-f="survey" value="${esc(it.survey)}"></div>
    <div class="fgroup"><label>Status</label><select data-f="status">${statusOpts(it.status)}</select></div></div>
  <div class="fgroup"><label>Description</label><textarea data-f="description">${esc(it.description)}</textarea></div>
  <div class="frow"><div class="fgroup"><label>Tag #</label><input data-f="tag" value="${esc(it.tag)}"></div>
    <div class="fgroup"><label>Serial #</label><input data-f="serial" value="${esc(it.serial)}"></div>
    <div class="fgroup"><label>Asset ID</label><input data-f="assetId" value="${esc(it.assetId)}"></div></div>
  <div class="frow"><div class="fgroup"><label>Date Assigned</label><input type="date" data-f="dateAssigned" value="${esc(it.dateAssigned)}"></div>
    <div class="fgroup"><label>Current bid ($)</label><input data-f="currentBid" value="${esc(it.currentBid)}" placeholder="live price now"></div>
    <div class="fgroup"><label>Est. value at purchase ($)</label><input data-f="amount" value="${esc(it.amount)}"></div></div>
  <div class="frow"><div class="fgroup"><label>Condition</label><input data-f="condition" value="${esc(it.condition)}"></div>
    <div class="fgroup"><label>Service year</label><input data-f="year" value="${esc(it.year)}"></div></div>
  <div class="frow"><div class="fgroup"><label>Dept ID</label><input data-f="deptId" value="${esc(it.deptId)}"></div>
    <div class="fgroup" style="flex:1.6"><label>Department</label><input data-f="dept" value="${esc(it.dept)}"></div></div>
  <div class="dsection">Platform Activity <span class="bar"></span></div>
  ${platBlock("GD",it.platforms.GD)}${platBlock("PS",it.platforms.PS)}
  <div class="dsection">Notes <span class="bar"></span></div>
  <div class="fgroup" style="margin-bottom:0"><label>Internal notes</label><textarea data-f="notes" style="min-height:92px">${esc(it.notes)}</textarea></div>`;
}
function setPath(o,path,v){ const p=path.split("."); let t=o; for(let i=0;i<p.length-1;i++) t=t[p[i]]; t[p[p.length-1]]=v; }
function bindDrawer(){
  document.querySelectorAll("#drawerBody [data-f]").forEach(el=>{
    const h=()=>{ let v; if(el.type==="checkbox") v=el.checked; else if(el.type==="number") v=el.value===""?0:Number(el.value); else v=el.value;
      setPath(drawerDraft,el.dataset.f,v);
      if(el.dataset.f.endsWith(".listed")){ const blk=el.closest("[data-platblock]"); if(blk) blk.classList.toggle("on",el.checked); }
      setDirty(); };
    el.addEventListener(el.type==="checkbox"?"change":"input",h);
  });
}
function saveDrawer(){ if(!editingId) return; const i=items.findIndex(x=>x.id===editingId); if(i>=0) items[i]=JSON.parse(JSON.stringify(drawerDraft));
  save(); render(); toast("Saved."); closeDrawer(true); }
function deleteFromDrawer(){ if(!editingId) return; const it=items.find(i=>i.id===editingId);
  if(!confirm(`Delete "${it.description||it.survey}"?`)) return;
  lastDeleted=[{...it}]; items=items.filter(i=>i.id!==editingId); save(); closeDrawer(true); render(); toast("Item deleted.","warn",true); }
function toggleArchiveInDrawer(){ if(!drawerDraft) return;
  drawerDraft.archived=!drawerDraft.archived; setDirty();
  const ab=document.getElementById("dArchive"); if(ab) ab.textContent=drawerDraft.archived?"Unarchive":"Archive"; }

/* ===================================================================
   Export / restore
   =================================================================== */
function exportJson(){ download(new Blob([JSON.stringify({version:2,exportedAt:new Date().toISOString(),items},null,2)],{type:"application/json"}),`auction-console-backup-${new Date().toISOString().slice(0,10)}.json`); toast("Backup downloaded."); }
function exportCsv(){
  const cols=["survey","description","tag","serial","assetId","dateAssigned","year","amount","currentBid","deptId","dept","condition","status","archived","GD_listed","GD_releases","GD_price","GD_currentBid","GD_soldPrice","GD_url","PS_listed","PS_releases","PS_price","PS_currentBid","PS_soldPrice","PS_url","notes"];
  const q=s=>`"${String(s==null?"":s).replace(/"/g,'""')}"`;
  const lines=[cols.join(",")];
  items.forEach(it=>lines.push(cols.map(c=>c.startsWith("GD_")?q(it.platforms.GD[c.slice(3)]):c.startsWith("PS_")?q(it.platforms.PS[c.slice(3)]):q(it[c])).join(",")));
  download(new Blob([lines.join("\n")],{type:"text/csv"}),`auction-console-${new Date().toISOString().slice(0,10)}.csv`); toast("CSV downloaded.");
}
function download(blob,name){ const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=name; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1000); }
function importJsonFile(file){ const fr=new FileReader(); fr.onload=()=>{ try{ const d=JSON.parse(fr.result); const arr=Array.isArray(d)?d:d.items; if(!Array.isArray(arr)) throw 0;
  items=arr.map(migrate); selected.clear(); save(); render(); toast(`Restored ${items.length} items.`); }catch(e){ toast("That file isn't a valid backup.","warn"); } }; fr.readAsText(file); }

/* ===================================================================
   Import modal
   =================================================================== */
function openImport(){ document.getElementById("importModal").classList.add("show"); document.getElementById("previewBar").classList.remove("show"); pendingRows=null; }
function closeImport(){ document.getElementById("importModal").classList.remove("show"); }
function handleFile(file){
  if(typeof XLSX==="undefined"){ toast("Spreadsheet engine didn't load — check that xlsx.full.min.js is present.","warn"); return; }
  const fr=new FileReader();
  fr.onload=()=>{ try{ const p=parseWorkbook(fr.result); if(!p.rows.length){ toast("Couldn't find survey rows in that file.","warn"); return; }
    pendingRows=p.rows; pendingSheet=p.sheet; refreshPreview(); }catch(e){ console.error(e); toast("Couldn't read that file.","warn"); } };
  fr.readAsArrayBuffer(file);
}
function refreshPreview(){
  if(!pendingRows) return;
  const auctionOnly=document.getElementById("optAuctionOnly").checked, group=document.getElementById("optGroup").checked;
  const prepared=rowsToItems(pendingRows,{auctionOnly,group});
  const auctionCount=pendingRows.filter(r=>r.action.includes("AUCTION")).length;
  const bar=document.getElementById("previewBar"); bar.classList.add("show");
  bar.innerHTML=`Found <b>${pendingRows.length}</b> rows in “${esc(pendingSheet)}”, <b>${auctionCount}</b> marked AUCTION → importing <b>${prepared.length}</b> ${group?"surveys":"line items"}.<button class="btn btn-sm btn-gold" id="confirmImport" style="margin-left:auto">Import ${prepared.length}</button>`;
  document.getElementById("confirmImport").onclick=()=>{ const merge=document.getElementById("optMerge").checked; const added=importItems(prepared,{merge}); selected.clear(); save(); render(); closeImport();
    toast(merge?`Added ${added} new item${added!==1?"s":""} (kept your edits).`:`Imported ${prepared.length} items.`); };
}

/* ===================================================================
   Toast
   =================================================================== */
let toastTimer=null;
function toast(msg,kind="",undo=false){ const t=document.getElementById("toast"); t.className="toast show "+(kind||"");
  t.innerHTML=esc(msg)+(undo&&lastDeleted?' <span class="act" id="undoBtn">Undo</span>':"");
  if(undo){ const u=document.getElementById("undoBtn"); if(u) u.onclick=()=>{ if(lastDeleted){ lastDeleted.forEach(d=>items.push(migrate(d))); lastDeleted=null; save(); render(); hideToast(); toast("Restored."); } }; }
  clearTimeout(toastTimer); toastTimer=setTimeout(hideToast,undo?6000:2600); }
function hideToast(){ document.getElementById("toast").classList.remove("show"); }

/* ===================================================================
   Wiring
   =================================================================== */
function toggleSel(id,on){ if(on) selected.add(id); else selected.delete(id);
  const tr=document.querySelector(`tr[data-id="${id}"]`); if(tr) tr.classList.toggle("sel",on);
  updateBulkbar();
  const v=visibleItems(), chkAll=document.getElementById("chkAll");
  chkAll.checked=v.length>0&&v.every(it=>selected.has(it.id)); chkAll.indeterminate=!chkAll.checked&&v.some(it=>selected.has(it.id)); }

function showExportMenu(){ const t=document.getElementById("toast"); t.className="toast show";
  t.innerHTML=`Export as <span class="act" id="exJson">JSON backup</span> · <span class="act" id="exCsv">CSV</span>`;
  document.getElementById("exJson").onclick=()=>{ exportJson(); hideToast(); };
  document.getElementById("exCsv").onclick=()=>{ exportCsv(); hideToast(); };
  clearTimeout(toastTimer); toastTimer=setTimeout(hideToast,5000); }

function wire(){
  document.getElementById("btnImport").onclick=openImport;
  document.getElementById("btnImportEmpty").onclick=openImport;
  document.getElementById("btnAdd").onclick=()=>{ const it=blankItem(); items.unshift(it); save(); render(); openDrawer(it.id); };
  document.getElementById("btnExport").onclick=showExportMenu;
  document.getElementById("btnRestore").onclick=()=>document.getElementById("fileJson").click();
  document.getElementById("fileJson").onchange=e=>{ if(e.target.files[0]) importJsonFile(e.target.files[0]); e.target.value=""; };
  document.getElementById("btnPull").onclick=openPull;
  document.getElementById("btnConvert").onclick=openConvert;
  document.getElementById("btnRecap").onclick=openRecap;
  wireConvert(); wirePull(); wireRecap();

  document.getElementById("globalSearch").addEventListener("input",e=>{ globalTerm=e.target.value; render(); });

  // header: sort + column text filter + open autofilter (delegated)
  const head=document.getElementById("headRow");
  head.addEventListener("click",e=>{
    const sortEl=e.target.closest("[data-sort]");
    if(sortEl){ const k=sortEl.dataset.sort; if(sortKey===k) sortDir*=-1; else { sortKey=k; sortDir=1; } render(); return; }
    const fbtn=e.target.closest("[data-colfilter]");
    if(fbtn){ e.stopPropagation(); openFilterPop(fbtn.dataset.colfilter, fbtn); return; }
  });
  head.addEventListener("input",e=>{
    const ti=e.target.closest("[data-coltext]");
    if(ti){ const k=ti.dataset.coltext; if(ti.value) colText[k]=ti.value; else delete colText[k];
      // re-render body only (keep focus): rebuild but restore focus
      const pos=ti.selectionStart; render();
      const again=document.querySelector(`[data-coltext="${k}"]`); if(again){ again.focus(); again.setSelectionRange(pos,pos); } }
  });
  head.addEventListener("click",e=>{ if(e.target.closest("[data-coltext]")) e.stopPropagation(); });

  // active filter pills
  document.getElementById("activeFilters").addEventListener("click",e=>{
    const c=e.target.closest("[data-clearcol]");
    if(c){ const k=c.dataset.clearcol; if(c.dataset.kind==="text") delete colText[k]; else delete colSet[k]; render(); return; }
    if(e.target.id==="clearAllFilters"){ colText={}; colSet={}; render(); }
  });

  // filter popup controls
  document.getElementById("fpSearch").addEventListener("input",e=>renderPopList(e.target.value));
  document.getElementById("fpList").addEventListener("change",e=>{ const chk=e.target.closest(".fpchk"); if(chk){ const v=chk.dataset.v; if(chk.checked) popChecked.add(v); else popChecked.delete(v); } });
  document.getElementById("fpAll").onclick=()=>{ popValues.forEach(([v])=>popChecked.add(v)); renderPopList(document.getElementById("fpSearch").value); };
  document.getElementById("fpNone").onclick=()=>{ popChecked.clear(); renderPopList(document.getElementById("fpSearch").value); };
  document.getElementById("fpApply").onclick=applyFilterPop;
  document.getElementById("fpCancel").onclick=closeFilterPop;
  document.addEventListener("click",e=>{ if(popCol&&!e.target.closest("#filterPop")&&!e.target.closest("[data-colfilter]")) closeFilterPop(); });
  window.addEventListener("resize",()=>{ if(popCol) closeFilterPop(); });

  // table body (delegated): checkbox, status select, edit, row click
  const tb=document.getElementById("tbody");
  tb.addEventListener("change",e=>{
    const chk=e.target.closest(".rowchk"); if(chk){ toggleSel(chk.dataset.id,chk.checked); return; }
    const ss=e.target.closest("[data-statusfor]");
    if(ss){ const it=items.find(i=>i.id===ss.dataset.statusfor); if(it){ it.status=ss.value; ss.className="statsel s-"+it.status; save();
      // refresh stats + any status filter without full re-render losing the select
      document.getElementById("statTotal").textContent=items.length;
      document.getElementById("statLive").textContent=items.filter(i=>i.status==="LIVE").length;
      document.getElementById("statRelist").textContent=items.filter(i=>i.status==="RELISTED").length;
      document.getElementById("statSold").textContent=items.filter(i=>["SOLD","PAID","PICKED_UP"].includes(i.status)).length;
      document.getElementById("statReview").textContent=items.filter(i=>i.status==="NEEDS_REVIEW").length;
      if(colSet.status||sortKey==="status") render();
    } }
  });
  tb.addEventListener("click",e=>{
    if(e.target.closest(".rowchk")||e.target.closest("[data-statusfor]")) return;
    const ed=e.target.closest("[data-edit]"); if(ed){ openDrawer(ed.dataset.edit); return; }
    const tr=e.target.closest("tr[data-id]"); if(tr) openDrawer(tr.dataset.id);
  });
  document.getElementById("chkAll").addEventListener("change",e=>{ const v=visibleItems();
    if(e.target.checked) v.forEach(it=>selected.add(it.id)); else v.forEach(it=>selected.delete(it.id)); render(); });

  // bulk
  document.getElementById("bulkApply").onclick=applyBulk;
  document.getElementById("bulkRelease").onclick=bulkRelease;
  document.getElementById("bulkDelete").onclick=bulkDelete;
  document.getElementById("bulkArchive").onclick=bulkArchive;
  document.getElementById("bulkUnarchive").onclick=bulkUnarchive;
  document.getElementById("bulkClear").onclick=()=>{ selected.clear(); render(); };
  if(document.getElementById("archiveToggle"))
    document.getElementById("archiveToggle").onclick=()=>{ showArchived=!showArchived; selected.clear(); render(); };

  // drawer
  document.getElementById("dClose").onclick=()=>closeDrawer(false);
  document.getElementById("dCancel").onclick=()=>closeDrawer(false);
  document.getElementById("scrim").onclick=()=>closeDrawer(false);
  document.getElementById("dSave").onclick=saveDrawer;
  document.getElementById("dDelete").onclick=deleteFromDrawer;
  const dab=document.getElementById("dArchive"); if(dab) dab.onclick=toggleArchiveInDrawer;

  // import modal
  document.getElementById("impClose").onclick=closeImport;
  document.getElementById("importModal").addEventListener("click",e=>{ if(e.target.id==="importModal") closeImport(); });
  const dz=document.getElementById("dropzone");
  dz.onclick=()=>document.getElementById("fileXlsx").click();
  document.getElementById("fileXlsx").onchange=e=>{ if(e.target.files[0]) handleFile(e.target.files[0]); e.target.value=""; };
  ["dragover","dragenter"].forEach(ev=>dz.addEventListener(ev,e=>{ e.preventDefault(); dz.classList.add("drag"); }));
  ["dragleave","drop"].forEach(ev=>dz.addEventListener(ev,e=>{ e.preventDefault(); dz.classList.remove("drag"); }));
  dz.addEventListener("drop",e=>{ const f=e.dataTransfer.files[0]; if(f) handleFile(f); });
  document.getElementById("optAuctionOnly").onchange=refreshPreview;
  document.getElementById("optGroup").onchange=refreshPreview;

  document.addEventListener("keydown",e=>{
    if(e.key==="Escape"){ closeFilterPop(); closeDrawer(false); closeImport();
      ["convertModal","pullModal","recapModal"].forEach(id=>document.getElementById(id).classList.remove("show")); }
    if((e.metaKey||e.ctrlKey)&&e.key==="s"&&editingId){ e.preventDefault(); saveDrawer(); }
  });
}

function fillBulkStatus(){ const sel=document.getElementById("bulkStatus"); STATUSES.forEach(s=>{ const o=document.createElement("option"); o.value=s; o.textContent=STATUS_LABEL[s]; sel.appendChild(o); }); }

/* ===================================================================
   v6: Convert / Pull / Recap modal handlers
   =================================================================== */
const F = window.AuctionFeatures;
function showModal(id){ document.getElementById(id).classList.add("show"); }
function hideModal(id){ document.getElementById(id).classList.remove("show"); }

/* ---- Convert HTML ---- */
let cvLastParsed=null;
function openConvert(){ showModal("convertModal"); }
function runConvert(target){
  const src=document.getElementById("cvInput").value.trim();
  if(!src){ toast("Paste a listing first.","warn"); return; }
  try{
    cvLastParsed=F.parseListingHtml(src);
    document.getElementById("cvOutput").value=F.convertListing(src,target);
    document.getElementById("cvOutDir").textContent="→ "+(target==="GD"?"GovDeals":"Public Surplus");
    const p=cvLastParsed;
    document.getElementById("cvParsed").innerHTML=
      `Read: <b>${esc(p.title||"—")}</b> · Condition <b>${esc(p.condition||"—")}</b> · Survey# <b>${esc(p.survey||"—")}</b> · Tag# <b>${esc(p.tag||"—")}</b>`;
  }catch(e){ console.error(e); toast("Couldn't parse that HTML.","warn"); }
}
function wireConvert(){
  document.getElementById("cvClose").onclick=()=>hideModal("convertModal");
  document.getElementById("convertModal").addEventListener("click",e=>{ if(e.target.id==="convertModal") hideModal("convertModal"); });
  document.getElementById("cvToGD").onclick=()=>runConvert("GD");
  document.getElementById("cvToPS").onclick=()=>runConvert("PS");
  document.getElementById("cvCopy").onclick=()=>{ const o=document.getElementById("cvOutput"); if(!o.value){ toast("Nothing to copy yet."); return; } copyText(o.value); };
  document.getElementById("cvToItem").onclick=()=>{
    const src=document.getElementById("cvInput").value.trim(); if(!src){ toast("Paste a listing first.","warn"); return; }
    const p=F.parseListingHtml(src); const it=itemFromParsed(p);
    items.unshift(it); save(); render(); hideModal("convertModal"); openDrawer(it.id); toast("Item created from listing.");
  };
}

/* Build an item record from a parsed listing block */
function itemFromParsed(p){
  const it=blankItem();
  it.survey=p.survey||""; it.tag=p.tag||""; it.description=p.title||"Untitled item";
  it.condition=p.condition||"";
  if(p.bullets&&p.bullets.length) it.notes=p.bullets.map(b=>"• "+b).join("\n");
  if(p.platform==="GD"){ it.platforms.GD.listed=true; }
  else if(p.platform==="PS"){ it.platforms.PS.listed=true; }
  return it;
}
/* Find the existing item that best matches a parsed listing (survey# > tag# > title) */
function matchItem(p){
  if(p.survey){ const m=items.find(i=>(i.survey||"").replace(/\s/g,"")===p.survey.replace(/\s/g,"")); if(m) return m; }
  if(p.tag){ const m=items.find(i=>(i.tag||"")===p.tag); if(m) return m; }
  if(p.title){ const t=p.title.toLowerCase(); const m=items.find(i=>i.description.toLowerCase().includes(t)||t.includes(i.description.toLowerCase())); if(m) return m; }
  return null;
}

/* ---- Pull / Paste listings ---- */
function openPull(){ document.getElementById("pullResult").classList.remove("show"); showModal("pullModal"); }
async function doFetch(){
  const box=document.getElementById("pullResult"); box.classList.add("show");
  box.innerHTML="Fetching… this can take a moment.";
  try{
    const r=await F.pullListings();
    const total=r.PS.length+r.GD.length;
    const via=r.usingScraper?"scraper Worker":"legacy CORS proxy";
    let html=`Public Surplus: <b>${r.PS.length}</b> · GovDeals: <b>${r.GD.length}</b> <span style="color:var(--ink-mute);font-size:11px">(via ${via})</span>.`;
    if(r.errors.length) html+=`<br><span style="color:#e06b58">Errors: ${esc(r.errors.join(" / "))}</span>`;
    if(total){
      const applied=applyScraped([...r.PS,...r.GD]);
      html+=`<br>Matched & updated <b>${applied}</b> existing item${applied!==1?"s":""} with live bids.`;
      save(); render();
    } else if(!r.errors.length){ html+=`<br>No listings returned — try the paste path below.`; }
    box.innerHTML=html;
  }catch(e){ box.innerHTML=`<span style="color:#e06b58">Fetch failed: ${esc(e.message)}. Use the paste path below.</span>`; }
}
/* apply scraped {platform,title,currentBid,url} rows onto matching items */
function applyScraped(rows){
  let n=0;
  for(const row of rows){
    const m=matchItem({title:row.title,survey:"",tag:""});
    if(!m) continue;
    const pk=row.platform==="GD"?"GD":"PS";
    m.platforms[pk].listed=true;
    if(row.currentBid) m.platforms[pk].currentBid=row.currentBid;
    if(row.url) m.platforms[pk].url=row.url;
    if(m.status==="PREP") m.status="LIVE";
    n++;
  }
  return n;
}
function applyPaste(asNew){
  const src=document.getElementById("plPaste").value.trim();
  if(!src){ toast("Paste a listing first.","warn"); return; }
  const p=F.parseListingHtml(src);
  if(asNew){ const it=itemFromParsed(p); items.unshift(it); save(); render(); hideModal("pullModal"); openDrawer(it.id); toast("Added as new item."); return; }
  const m=matchItem(p);
  if(!m){ toast("No matching item found — try 'Add as new item'.","warn"); return; }
  if(p.condition) m.condition=p.condition;
  if(p.tag&&!m.tag) m.tag=p.tag;
  const pk=p.platform==="GD"?"GD":(p.platform==="PS"?"PS":null);
  if(pk) m.platforms[pk].listed=true;
  if(p.bullets&&p.bullets.length){ const note="Listing details:\n"+p.bullets.map(b=>"• "+b).join("\n"); m.notes=m.notes?m.notes+"\n\n"+note:note; }
  save(); render(); hideModal("pullModal"); openDrawer(m.id);
  toast(`Updated "${m.survey||m.description}".`);
}
function wirePull(){
  document.getElementById("plClose").onclick=()=>hideModal("pullModal");
  document.getElementById("pullModal").addEventListener("click",e=>{ if(e.target.id==="pullModal") hideModal("pullModal"); });
  document.getElementById("plFetch").onclick=doFetch;
  document.getElementById("plApplyPaste").onclick=()=>applyPaste(false);
  document.getElementById("plPasteAsNew").onclick=()=>applyPaste(true);
  const sv=document.getElementById("plScraperUrl");
  if(sv){
    try{ sv.value=localStorage.getItem("csulb_scraper_url")||""; }catch(e){}
    sv.addEventListener("change",()=>{
      try{ localStorage.setItem("csulb_scraper_url",sv.value.trim().replace(/\/+$/,"")); }catch(e){}
      const hint=document.getElementById("plScraperSaved"); if(hint) hint.style.display=sv.value.trim()?"block":"none";
    });
  }
}

/* ---- Recap ---- */
function openRecap(){ showModal("recapModal"); }
function wireRecap(){
  document.getElementById("rcClose").onclick=()=>hideModal("recapModal");
  document.getElementById("recapModal").addEventListener("click",e=>{ if(e.target.id==="recapModal") hideModal("recapModal"); });
  document.getElementById("rcGen").onclick=()=>{
    const out=F.buildRecap(items,{
      weekOf:document.getElementById("rcWeek").value.trim()||undefined,
      to:document.getElementById("rcTo").value.trim()||undefined,
      from:document.getElementById("rcFrom").value.trim()||undefined
    });
    document.getElementById("rcOut").value=out;
  };
  document.getElementById("rcCopy").onclick=()=>{ const v=document.getElementById("rcOut").value; if(!v){ toast("Generate first."); return; } copyText(v); };
  document.getElementById("rcDownload").onclick=()=>{ const v=document.getElementById("rcOut").value; if(!v){ toast("Generate first."); return; }
    download(new Blob([v],{type:"text/plain"}),`recap-${new Date().toISOString().slice(0,10)}.txt`); toast("Recap downloaded."); };
}

/* clipboard helper with fallback */
function copyText(t){
  if(navigator.clipboard&&navigator.clipboard.writeText){ navigator.clipboard.writeText(t).then(()=>toast("Copied."),()=>fallbackCopy(t)); }
  else fallbackCopy(t);
}
function fallbackCopy(t){ const ta=document.createElement("textarea"); ta.value=t; document.body.appendChild(ta); ta.select();
  try{ document.execCommand("copy"); toast("Copied."); }catch(e){ toast("Copy failed — select manually.","warn"); } document.body.removeChild(ta); }

/* ---------- boot ---------- */
load(); fillBulkStatus(); wire(); render();
