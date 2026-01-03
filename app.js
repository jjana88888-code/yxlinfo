/* YB Rankboard v3 (robust) – Excel 쿼리/헤더가 바뀌어도 최대한 자동으로 붙게 만든 버전.
   - 시트명이 바뀌어도: '헤더(순위/이름/값/갱신)'를 포함한 시트를 자동 탐색
   - 컬럼명이 살짝 바뀌어도: 후보 목록(동의어)로 자동 매핑
   - 순위 컬럼이 없으면: 값 내림차순 기준으로 자동 순위 생성
*/
const CONFIG = {
  defaultXlsxUrl: "./data/YB.xlsx",

  // ✅ 여기서 필요하면 동의어만 추가하면 됨
  headerCandidates: {
    rank: ["순위", "랭킹", "등수", "Rank", "rank"],
    name: ["비제이명", "BJ명", "BJ", "이름", "닉네임", "스트리머", "방송인", "name", "Name"],
    value: ["월별 누적별풍선", "월별누적별풍선", "누적별풍선", "누적 별풍선", "별풍선", "풍력", "기여도", "value", "Value"],
    refresh: ["새로고침시간", "새로고침 시간", "갱신시간", "업데이트시간", "업데이트 시간", "refresh", "Refresh"],
  },

  // 티어 기준(원하면 숫자만 바꾸면 됨)
  tiers: [
    { key: "T1", min: 100000, label: "100만+" },
    { key: "T2", min: 50000,  label: "50만+" },
    { key: "T3", min: 20000,  label: "20만+" },
  ],
};

let rawRows = [];
let viewRows = [];
let currentSort = "rank";

const el = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat("ko-KR");

function setHint(t){ el("dataHint").textContent = t; }

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

function tierOf(value){
  const [a,b,c] = CONFIG.tiers;
  if (value >= a.min) return { key:"T1", dot:"t1", cap:"cap1", text: a.label, min: a.min };
  if (value >= b.min) return { key:"T2", dot:"t2", cap:"cap2", text: b.label, min: b.min };
  if (value >= c.min) return { key:"T3", dot:"t3", cap:"cap3", text: c.label, min: c.min };
  return { key:"-", dot:"", cap:"cap0", text:"-", min: 0 };
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
  // headers: array of original header strings
  const norm = headers.map(normalizeHeader);
  const pick = (key) => {
    const candidates = CONFIG.headerCandidates[key].map(normalizeHeader);
    for (let i=0;i<norm.length;i++){
      if (!norm[i]) continue;
      if (candidates.includes(norm[i])) return headers[i]; // return original header
    }
    // fuzzy contains match
    for (let i=0;i<norm.length;i++){
      if (!norm[i]) continue;
      for (const c of candidates){
        if (c && norm[i].includes(c)) return headers[i];
        if (c && c.includes(norm[i])) return headers[i];
      }
    }
    return null;
  };

  return {
    rank: pick("rank"),
    name: pick("name"),
    value: pick("value"),
    refresh: pick("refresh"),
  };
}

function findBestSheet(workbook){
  // try each sheet: convert first row as header + map
  for (const sheetName of workbook.SheetNames){
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (!rows || !rows.length) continue;
    const headers = (rows[0] || []).filter(h => h !== null && h !== undefined);
    if (!headers.length) continue;

    const map = buildHeaderMap(rows[0] || []);
    // if we can find at least name & value, accept (rank optional)
    if (map.name && map.value){
      return { sheetName, map };
    }
  }
  // fallback: first sheet
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

  // If header mapping failed, give a clear hint
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

  // If rank missing for many rows, auto-generate rank by value desc
  const hasValidRank = out.some(r => Number.isFinite(r.rank));
  if (!hasValidRank){
    out.sort((a,b)=> (b.value||0) - (a.value||0));
    out.forEach((r, i) => r.rank = i + 1);
  }

  // If rank exists but gaps/NaN exist, fill missing ranks after sorting by rank then value
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

function updateLegend(){
  el("tier1Text").textContent = fmt.format(CONFIG.tiers[0].min) + "+";
  el("tier2Text").textContent = fmt.format(CONFIG.tiers[1].min) + "+";
  el("tier3Text").textContent = fmt.format(CONFIG.tiers[2].min) + "+";
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

  if (q){
    rows = rows.filter(r => r.name.toLowerCase().includes(q));
  }

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
    el("tableMeta").textContent = "-";
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
    return `
      <div class="${rowTopClass(r.rank)}">
        <div class="topFx"></div>
        <div class="rankNum ${rankClass(r.rank)}">${r.rank}${topIcon(r.rank)}</div>

        <div class="nameCol">
          <div class="nameLine">
            <div class="name">${escapeHtml(r.name)}</div>
            <div class="tierPill">
              ${t.text === "-" ? `<strong>-</strong>` : `<span class="dot ${t.dot}"></span><strong>${t.text}</strong>`}
            </div>
          </div>
          <div class="bar">
            <div class="barFill" style="width:${width}%"></div>
            <div class="barCap ${t.cap}" title="티어"></div>
          </div>
        </div>

        <div class="value">
          ${fmt.format(r.value||0)}
          <div class="valueSub">점유율 ${Math.round(share*1000)/10}%</div>
        </div>
      </div>
    `;
  }).join("");

  const { avg, count, refresh } = computeMeta(rawRows);
  el("metaCount").textContent = fmt.format(count) + "명";
  el("metaAvg").textContent = fmt.format(avg);
  el("metaRefresh").textContent = formatRefresh(refresh);
  el("tableMeta").textContent = `표시 ${viewRows.length}명 · 갱신 ${formatRefresh(refresh)}`;
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
  updateLegend();

  el("searchInput").addEventListener("input", applyFilters);
  el("sortSelect").addEventListener("change", (e) => {
    currentSort = e.target.value;
    applyFilters();
  });
  el("btnExport").addEventListener("click", exportCsv);

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
