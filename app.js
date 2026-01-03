/* YB Rankboard v5 (clean) – 쿼리/헤더 변경에도 견고 + UI 정리 + 티어/바 의미 일치
   - 티어는 '값(누적)'의 구간을 의미하며, 막대 색도 동일 기준으로 변합니다.
   - 시트/헤더 자동 인식(동의어 매핑) 유지
*/
const CONFIG = {
  defaultXlsxUrl: "./data/YB.xlsx",

  headerCandidates: {
    rank: ["순위", "랭킹", "등수", "Rank", "rank"],
    name: ["비제이명", "BJ명", "BJ", "이름", "닉네임", "스트리머", "방송인", "name", "Name"],
    value: ["월별 누적별풍선", "월별누적별풍선", "누적별풍선", "누적 별풍선", "별풍선", "풍력", "기여도", "value", "Value"],
    refresh: ["새로고침시간", "새로고침 시간", "갱신시간", "업데이트시간", "업데이트 시간", "refresh", "Refresh"],
  },

  // ✅ IMPORTANT: 숫자(min)와 라벨이 일치해야 함.
  // 지금 데이터가 127,144처럼 '십만 단위'라면: 10만+/5만+/2만+가 자연스럽습니다.
  tiers: [
    { key: "T1", min: 100000, label: "10만+" },
    { key: "T2", min: 50000,  label: "5만+" },
    { key: "T3", min: 20000,  label: "2만+" },
  ],

  // 점유율 표시(지저분하면 false)
  showShare: false,
};

let rawRows = [];
let viewRows = [];
let currentSort = "rank";

const el = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat("ko-KR");

function setHint(t){ const h = document.getElementById("dataHint"); if (h) h.textContent = t; }

function crownSvg(){
  return `
  <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M3.5 8.5l4.8 4.2L12 6.8l3.7 5.9 4.8-4.2V18a2 2 0 0 1-2 2H5.5a2 2 0 0 1-2-2V8.5Z" stroke="rgba(255,255,255,.85)" stroke-width="1.6" />
  </svg>`;
}
function medalSvg(){
  return `
  <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M8 3h8l-2.8 5.2a6.5 6.5 0 1 1-2.4 0L8 3Z" stroke="rgba(255,255,255,.85)" stroke-width="1.6"/>
    <path d="M12 10.2a3.8 3.8 0 1 0 0 7.6 3.8 3.8 0 0 0 0-7.6Z" stroke="rgba(255,255,255,.85)" stroke-width="1.6"/>
  </svg>`;
}

function rankClass(r){
  if (r === 1) return "top1";
  if (r === 2) return "top2";
  if (r === 3) return "top3";
  return "";
}
function rowTopClass(r){
  if (r === 1) return "row top1";
  if (r === 2) return "row top2";
  if (r === 3) return "row top3";
  return "row";
}
function topIcon(r){
  if (r === 1) return `<span class="medal" title="1등">${crownSvg()}</span>`;
  if (r === 2) return `<span class="medal" title="2등">${medalSvg()}</span>`;
  if (r === 3) return `<span class="medal" title="3등">${medalSvg()}</span>`;
  return "";
}

function escapeHtml(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}
function normalizeHeader(h){
  return String(h ?? "").trim().toLowerCase().replace(/\s+/g, "");
}
function buildHeaderMap(headers){
  const norm = headers.map(normalizeHeader);
  const pick = (key) => {
    const candidates = CONFIG.headerCandidates[key].map(normalizeHeader);
    for (let i=0;i<norm.length;i++){
      if (!norm[i]) continue;
      if (candidates.includes(norm[i])) return headers[i];
    }
    for (let i=0;i<norm.length;i++){
      if (!norm[i]) continue;
      for (const c of candidates){
        if (c && norm[i].includes(c)) return headers[i];
        if (c && c.includes(norm[i])) return headers[i];
      }
    }
    return null;
  };
  return { rank: pick("rank"), name: pick("name"), value: pick("value"), refresh: pick("refresh") };
}
function findBestSheet(workbook){
  for (const sheetName of workbook.SheetNames){
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!rows || !rows.length) continue;
    const map = buildHeaderMap(rows[0] || []);
    if (map.name && map.value) return { sheetName, map };
  }
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const map = buildHeaderMap((rows && rows[0]) || []);
  return { sheetName, map };
}
function parseExcelToRows(workbook){
  const { sheetName, map } = findBestSheet(workbook);
  const ws = workbook.Sheets[sheetName];
  if (!ws) throw new Error("워크시트가 없습니다.");

  const json = XLSX.utils.sheet_to_json(ws, { defval: null });

  if (!map.name || !map.value){
    const sampleHeaders = Object.keys(json?.[0] || {}).slice(0, 12).join(", ");
    throw new Error(`엑셀 헤더를 인식하지 못했습니다. (시트: ${sheetName}) 헤더 예시: ${sampleHeaders}`);
  }

  const out = json.map(r => {
    const rankRaw = map.rank ? r[map.rank] : null;
    const name = (r[map.name] ?? "").toString().trim();
    const value = Number(r[map.value] ?? 0);
    const refresh = map.refresh ? r[map.refresh] : null;
    const rank = rankRaw === null || rankRaw === undefined || rankRaw === "" ? NaN : Number(rankRaw);
    return { rank, name, value, refresh };
  }).filter(r => r.name);

  const hasValidRank = out.some(r => Number.isFinite(r.rank));
  if (!hasValidRank){
    out.sort((a,b)=> (b.value||0) - (a.value||0));
    out.forEach((r, i) => r.rank = i + 1);
  }

  out.sort((a,b)=>{
    const ar = Number.isFinite(a.rank) ? a.rank : 1e9;
    const br = Number.isFinite(b.rank) ? b.rank : 1e9;
    if (ar !== br) return ar - br;
    return (b.value||0) - (a.value||0);
  });
  let next = 1;
  for (const r of out){
    if (!Number.isFinite(r.rank)) r.rank = next;
    next = Math.max(next, r.rank + 1);
  }

  return { rows: out, sheetName };
}

async function loadDefaultXlsx(){
  setHint("data/YB.xlsx 불러오는 중…");
  const res = await fetch(CONFIG.defaultXlsxUrl, { cache: "no-store" });
  if (!res.ok) throw new Error("기본 엑셀을 찾지 못했습니다. (data/YB.xlsx 업로드 필요)");
  const ab = await res.arrayBuffer();
  return XLSX.read(ab, { type:"array" });
}
function readFileAsArrayBuffer(file){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function tierOf(value){
  const [a,b,c] = CONFIG.tiers;
  if (value >= a.min) return { key:"T1", label: a.label, swatch:"s1", colorVar:"var(--t1)" };
  if (value >= b.min) return { key:"T2", label: b.label, swatch:"s2", colorVar:"var(--t2)" };
  if (value >= c.min) return { key:"T3", label: c.label, swatch:"s3", colorVar:"var(--t3)" };
  return { key:"-", label: "-", swatch:"", colorVar:"var(--t0)" };
}

function renderLegend(){
  const host = document.getElementById("legendInline");
  if (!host) return;
  host.innerHTML = CONFIG.tiers.map((t, idx) => {
    const cls = idx === 0 ? "s1" : (idx === 1 ? "s2" : "s3");
    // 라벨은 실제 값 기준 구간이라는 의미가 명확해야 함
    return `<span class="legendChip" title="누적 값 구간">
      <span class="swatch ${cls}"></span>${t.label}
    </span>`;
  }).join("");
}

function computeMeta(rows){
  const total = rows.reduce((a,b)=>a+(b.value||0),0);
  const count = rows.length || 1;
  const avg = Math.round(total / count);
  let refresh = null;
  for (const r of rows){
    if (r.refresh != null){ refresh = r.refresh; break; }
  }
  return { total, avg, count, refresh };
}
function formatRefresh(v){
  if (!v) return "-";
  if (typeof v === "number"){
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")} ${String(d.H).padStart(2,"0")}:${String(d.M).padStart(2,"0")}`;
  }
  if (v instanceof Date) return v.toLocaleString("ko-KR");
  return String(v);
}

function applyFilters(){
  const q = el("searchInput").value.trim().toLowerCase();
  let rows = [...rawRows];

  if (q) rows = rows.filter(r => r.name.toLowerCase().includes(q));

  if (currentSort === "rank"){
    rows.sort((a,b)=> a.rank - b.rank);
  } else if (currentSort === "value"){
    rows.sort((a,b)=> (b.value||0) - (a.value||0));
  } else {
    rows.sort((a,b)=> a.name.localeCompare(b.name, "ko"));
  }

  viewRows = rows;
  render();
}

function render(){
  const rowsEl = el("rows");
  if (!viewRows.length){
    rowsEl.innerHTML = `<div class="empty">표시할 데이터가 없습니다.</div>`;
    const tm0 = document.getElementById("tableMeta"); if (tm0) tm0.textContent = "-";
    el("metaCount").textContent = "-";
    el("metaAvg").textContent = "-";
    el("metaRefresh").textContent = "-";
    return;
  }

  const totalAll = rawRows.reduce((a,b)=>a+(b.value||0),0) || 1;
  const maxVal = Math.max(...rawRows.map(r=>r.value||0), 1);

  rowsEl.innerHTML = viewRows.map(r => {
    const t = tierOf(r.value||0);
    const width = Math.max(2, Math.min(100, Math.round((r.value||0) / maxVal * 100)));
    const share = (r.value||0)/totalAll;

    const shareHtml = CONFIG.showShare
      ? `<div class="valueSub">점유율 ${(Math.round(share*1000)/10)}%</div>`
      : ``;

    return `
      <div class="${rowTopClass(r.rank)}">
        <div class="rankNum ${rankClass(r.rank)}">${r.rank}${topIcon(r.rank)}</div>

        <div class="nameCol">
          <div class="nameLine">
            <div class="name">${escapeHtml(r.name)}</div>
          </div>
          <div class="bar">
            <div class="barFill" style="width:${width}%; background:${t.colorVar};"></div>
          </div>
        </div>

        <div class="value">
          ${fmt.format(r.value||0)}
          ${shareHtml}
        </div>
      </div>
    `;
  }).join("");

  const { avg, count, refresh } = computeMeta(rawRows);
  el("metaCount").textContent = fmt.format(count) + "명";
  el("metaAvg").textContent = fmt.format(avg);
  el("metaRefresh").textContent = formatRefresh(refresh);
  const tm = document.getElementById("tableMeta"); if (tm) tm.textContent = `표시 ${viewRows.length}명 · 갱신 ${formatRefresh(refresh)}`;
}

function exportCsv(){
  if (!viewRows.length) return;
  const totalAll = rawRows.reduce((a,b)=>a+(b.value||0),0) || 1;

  const header = ["순위","이름","값","점유율"];
  const lines = viewRows.map(r => {
    const share = (r.value||0)/totalAll;
    return [r.rank, r.name, r.value||0, (share*100).toFixed(1)+"%"].join(",");
  });
  const csv = [header.join(","), ...lines].join("\n");
  const blob = new Blob([csv], { type:"text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "YB_ranking.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function loadAndRender(workbook, label){
  const parsed = parseExcelToRows(workbook);
  rawRows = parsed.rows;
  setHint(`${label} 로드 완료 (시트: ${parsed.sheetName})`);
  applyFilters();
}

async function boot(){
  renderLegend();

  el("searchInput").addEventListener("input", applyFilters);
  el("sortSelect").addEventListener("change", (e) => {
    currentSort = e.target.value;
    applyFilters();
  });
  const be = document.getElementById("btnExport"); if (be) be.addEventListener("click", exportCsv);

el("btnReload").addEventListener("click", async () => {
    try{
      const wb = await loadDefaultXlsx();
      await loadAndRender(wb, "data/YB.xlsx");
    }catch(e){
      setHint(e.message);
      console.error(e);
    }
  });

  el("fileInput").addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try{
      setHint("업로드 파일 읽는 중…");
      const ab = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(ab, { type:"array" });
      await loadAndRender(wb, file.name);
    }catch(err){
      setHint("엑셀을 읽지 못했습니다. 시트/헤더명을 확인해주세요.");
      console.error(err);
    }
  });

  try{
    const wb = await loadDefaultXlsx();
    await loadAndRender(wb, "data/YB.xlsx");
  }catch(e){
    setHint("data/YB.xlsx 자동 로드 실패 → 엑셀 업로드 버튼을 사용하세요");
    rawRows = [];
    viewRows = [];
    render();
    console.warn(e);
  }
}

boot();


// ===== Effects Pack (v6) =====
function applyRankClasses(){
  document.querySelectorAll('.row').forEach((row, i)=>{
    row.classList.add(`rank-${i+1}`);
  });
}

function applyAvgLine(avgRatio){
  const board = document.querySelector('.rows');
  if(!board) return;
  let line = document.querySelector('.avgLine');
  if(!line){
    line = document.createElement('div');
    line.className='avgLine';
    board.appendChild(line);
  }
  line.style.left = `${Math.min(100, Math.max(0, avgRatio*100))}%`;
}

if (typeof render === 'function') {
  const _render = render;
  render = function(...args){
    _render.apply(this, args);
    applyRankClasses();
    try{
      if(window.__meta && window.__meta.avg && window.__meta.max){
        applyAvgLine(window.__meta.avg / window.__meta.max);
      }
    }catch(e){}
  }
}


// ===== v6a metaTotal fill (best-effort) =====
function setMetaTotalFromRows(){
  const el = document.getElementById('metaTotal');
  if(!el) return;
  // Try to compute from rendered DOM values (data-value attr) first
  const vals = Array.from(document.querySelectorAll('.row [data-raw]')).map(n=>Number(n.getAttribute('data-raw'))).filter(Number.isFinite);
  if(vals.length){
    const sum = vals.reduce((a,b)=>a+b,0);
    el.textContent = new Intl.NumberFormat('ko-KR').format(sum);
  }
}
document.addEventListener('DOMContentLoaded', ()=>setTimeout(setMetaTotalFromRows, 0));
