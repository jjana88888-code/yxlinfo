/* YB Rankboard – renders a reference-like ranking list.
   Excel 기반(업로드한 YB.xlsx 기준):
   - 시트: 쿼리1
   - 컬럼: 순위, 비제이명, 월별 누적별풍선, 새로고침시간
*/
const CONFIG = {
  defaultXlsxUrl: "./data/YB.xlsx",
  sheetName: "쿼리1",
  columns: {
    rank: "순위",
    name: "비제이명",
    value: "월별 누적별풍선",
    refresh: "새로고침시간",
  },
  // 티어 기준(원하면 숫자만 바꾸면 됨)
  tiers: [
    { key: "T1", min: 100000, label: "100만+" },   // 라벨은 원하는대로 바꿔도 됨
    { key: "T2", min: 50000,  label: "50만+" },
    { key: "T3", min: 20000,  label: "20만+" },
  ],
  // 헤더 기간 텍스트(원하면 여기 수정)
  periodMain: "25.12.01 — 12.31",
  periodSub: "23:59 기준 집계",
};

let rawRows = [];
let viewRows = [];
let currentSort = "rank";

const el = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat("ko-KR");

function setHint(t){ el("dataHint").textContent = t; }

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

function escapeHtml(s){
  return String(s)
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function parseExcelToRows(workbook){
  const ws = workbook.Sheets[CONFIG.sheetName] || workbook.Sheets[workbook.SheetNames[0]];
  if (!ws) throw new Error("워크시트가 없습니다.");

  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

  const out = rows.map(r => {
    const rank = Number(r[CONFIG.columns.rank]);
    const name = (r[CONFIG.columns.name] ?? "").toString().trim();
    const value = Number(r[CONFIG.columns.value] ?? 0);
    const refresh = r[CONFIG.columns.refresh];
    return { rank, name, value, refresh };
  }).filter(r => r.name && !Number.isNaN(r.rank));

  return out;
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

function updateHeader(){
  el("periodMain").textContent = CONFIG.periodMain;
  el("periodSub").textContent = "⏱ " + CONFIG.periodSub;

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
    return;
  }

  const totalAll = rawRows.reduce((a,b)=>a+(b.value||0),0) || 1;
  const maxVal = Math.max(...rawRows.map(r=>r.value||0), 1);

  rowsEl.innerHTML = viewRows.map(r => {
    const t = tierOf(r.value||0);
    const width = Math.max(2, Math.min(100, Math.round((r.value||0) / maxVal * 100)));
    const share = (r.value||0)/totalAll;
    return `
      <div class="row">
        <div class="rankBox">
          <div class="rankNum ${rankClass(r.rank)}">${r.rank}</div>
        </div>

        <div class="nameCol">
          <div class="nameLine">
            <div class="name">${escapeHtml(r.name)}</div>
            <div class="tierMark">
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

  const { total, avg, count, refresh } = computeMeta(rawRows);
  el("metaTotal").textContent = fmt.format(total);
  el("metaAvg").textContent = fmt.format(avg);
  el("metaCount").textContent = fmt.format(count) + "명";
  el("avgText").textContent = "평균: " + fmt.format(avg);
  el("tableMeta").textContent = `표시 ${viewRows.length}명 · 갱신 ${formatRefresh(refresh)}`;
}

function exportCsv(){
  if (!viewRows.length) return;
  const totalAll = rawRows.reduce((a,b)=>a+(b.value||0),0) || 1;

  const header = ["순위","비제이명","월별 누적별풍선","점유율"];
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

async function boot(){
  updateHeader();

  el("searchInput").addEventListener("input", applyFilters);
  el("sortSelect").addEventListener("change", (e) => {
    currentSort = e.target.value;
    applyFilters();
  });
  el("btnExport").addEventListener("click", exportCsv);

  el("btnReload").addEventListener("click", async () => {
    try{
      const wb = await loadDefaultXlsx();
      rawRows = parseExcelToRows(wb);
      setHint("기본 엑셀 로드 완료");
      applyFilters();
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
      rawRows = parseExcelToRows(wb);
      setHint(`업로드 로드 완료: ${file.name}`);
      applyFilters();
    }catch(err){
      setHint("엑셀을 읽지 못했습니다. 시트/컬럼명을 확인해주세요.");
      console.error(err);
    }
  });

  try{
    const wb = await loadDefaultXlsx();
    rawRows = parseExcelToRows(wb);
    setHint("data/YB.xlsx 로드 완료");
    applyFilters();
  }catch(e){
    setHint("data/YB.xlsx 자동 로드 실패 → 엑셀 업로드 버튼을 사용하세요");
    rawRows = [];
    viewRows = [];
    render();
    console.warn(e);
  }
}

boot();
