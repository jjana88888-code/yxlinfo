/* YB Dashboard – reads /data/YB.xlsx (default) or user uploaded file.
   Excel 구조(현재 파일 기준):
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
  // 티어 기준(원하면 숫자만 바꿔)
  tiers: [
    { key: "T1", min: 100000 },
    { key: "T2", min: 50000 },
    { key: "T3", min: 20000 },
  ],
};

let rawRows = [];
let viewRows = [];
let currentSort = "rank";
let chart = null;

const el = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat("ko-KR");
const pct = new Intl.NumberFormat("ko-KR", { style: "percent", maximumFractionDigits: 1 });

function setHint(text){ el("dataHint").textContent = text; }

function tierOf(value){
  const t1 = CONFIG.tiers[0]?.min ?? Infinity;
  const t2 = CONFIG.tiers[1]?.min ?? Infinity;
  const t3 = CONFIG.tiers[2]?.min ?? Infinity;

  if (value >= t1) return { label: `Tier 1`, cls: "pill1", dot:"t1" };
  if (value >= t2) return { label: `Tier 2`, cls: "pill2", dot:"t2" };
  if (value >= t3) return { label: `Tier 3`, cls: "pill3", dot:"t3" };
  return { label: "-", cls: "", dot:"" };
}

function rankBadgeClass(r){
  if (r === 1) return "rankBadge rankTop1";
  if (r === 2) return "rankBadge rankTop2";
  if (r === 3) return "rankBadge rankTop3";
  return "rankBadge";
}

function parseExcelToRows(workbook){
  const ws = workbook.Sheets[CONFIG.sheetName] || workbook.Sheets[workbook.SheetNames[0]];
  if (!ws) throw new Error("워크시트가 없습니다.");

  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  // normalize
  const out = rows.map((r) => {
    const rank = Number(r[CONFIG.columns.rank]);
    const name = (r[CONFIG.columns.name] ?? "").toString().trim();
    const value = Number(r[CONFIG.columns.value] ?? 0);
    const refresh = r[CONFIG.columns.refresh]; // can be string or Date-like
    return { rank, name, value, refresh };
  }).filter(r => r.name && !Number.isNaN(r.rank));

  return out;
}

async function loadDefaultXlsx(){
  setHint("data/YB.xlsx 불러오는 중…");
  const res = await fetch(CONFIG.defaultXlsxUrl, { cache: "no-store" });
  if (!res.ok) throw new Error("기본 엑셀을 찾지 못했습니다. (data/YB.xlsx 업로드 필요)");
  const ab = await res.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });
  return wb;
}

function readFileAsArrayBuffer(file){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function updateTierText(){
  el("tier1Text").textContent = fmt.format(CONFIG.tiers[0].min) + "+";
  el("tier2Text").textContent = fmt.format(CONFIG.tiers[1].min) + "+";
  el("tier3Text").textContent = fmt.format(CONFIG.tiers[2].min) + "+";
}

function computeKpis(rows){
  const count = rows.length;
  const total = rows.reduce((a,b)=>a + (b.value||0), 0);
  const avg = count ? Math.round(total / count) : 0;

  // refresh time: choose max or first non-null
  let refresh = null;
  for (const r of rows){
    if (r.refresh != null){ refresh = r.refresh; break; }
  }

  return { count, total, avg, refresh };
}

function formatRefresh(v){
  if (!v) return "-";
  // SheetJS sometimes returns Excel date serial; sometimes string.
  if (typeof v === "number") {
    // Excel serial -> JS date
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

  // sort
  if (currentSort === "rank"){
    rows.sort((a,b)=> a.rank - b.rank);
  } else if (currentSort === "value"){
    rows.sort((a,b)=> (b.value||0) - (a.value||0));
  } else if (currentSort === "name"){
    rows.sort((a,b)=> a.name.localeCompare(b.name, "ko"));
  }

  viewRows = rows;
  renderAll();
}

function renderKpis(){
  const {count,total,avg,refresh} = computeKpis(rawRows);
  el("kpiCount").textContent = fmt.format(count) + "명";
  el("kpiTotal").textContent = fmt.format(total);
  el("kpiAvg").textContent = fmt.format(avg);
  el("kpiRefresh").textContent = formatRefresh(refresh);

  // extra sub text
  const top = [...rawRows].sort((a,b)=> (b.value||0)-(a.value||0))[0];
  if (top){
    el("kpiTotalSub").textContent = `1위: ${top.name} (${fmt.format(top.value)})`;
  }
  el("kpiAvgSub").textContent = `총합 ÷ 인원`;
  el("kpiCountSub").textContent = `랭킹 기준`;
  el("kpiRefreshSub").textContent = `쿼리 기준`;
}

function renderTable(){
  const tbody = el("tableBody");
  const rows = viewRows;

  if (!rows.length){
    tbody.innerHTML = `<tr><td colspan="5" class="empty">표시할 데이터가 없습니다.</td></tr>`;
    el("tableMeta").textContent = "-";
    return;
  }

  const total = rawRows.reduce((a,b)=>a+(b.value||0),0) || 1;

  tbody.innerHTML = rows.map(r=>{
    const t = tierOf(r.value||0);
    const share = (r.value||0)/total;
    return `
      <tr>
        <td class="col-rank"><span class="${rankBadgeClass(r.rank)}">${r.rank}</span></td>
        <td class="col-name"><strong>${escapeHtml(r.name)}</strong></td>
        <td class="col-tier">
          ${t.label === "-" ? `<span class="tierPill">-</span>` :
            `<span class="tierPill ${t.cls}"><span class="dot ${t.dot}"></span>${t.label}</span>`}
        </td>
        <td class="col-value"><strong>${fmt.format(r.value||0)}</strong></td>
        <td class="col-share">${pct.format(share)}</td>
      </tr>
    `;
  }).join("");

  el("tableMeta").textContent = `표시 ${rows.length}명 / 전체 ${rawRows.length}명`;
}

function renderChart(){
  const topN = Number(el("topN").value);
  const sorted = [...rawRows].sort((a,b)=> (b.value||0)-(a.value||0));
  const slice = topN === 999 ? sorted : sorted.slice(0, topN);

  const labels = slice.map(r=> r.name);
  const data = slice.map(r=> r.value||0);

  const ctx = el("barChart");
  if (!chart){
    chart = new Chart(ctx, {
      type: "bar",
      data: { labels, datasets: [{ label: "월별 누적별풍선", data }] },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (c) => ` ${fmt.format(c.raw)}`
            }
          }
        },
        scales: {
          x: {
            ticks: { color: getComputedStyle(document.documentElement).getPropertyValue("--muted") },
            grid: { display: false }
          },
          y: {
            ticks: {
              color: getComputedStyle(document.documentElement).getPropertyValue("--muted"),
              callback: (v) => fmt.format(v)
            },
            grid: { color: "rgba(255,255,255,.08)" }
          }
        }
      }
    });
  } else {
    chart.data.labels = labels;
    chart.data.datasets[0].data = data;
    chart.update();
  }
}

function renderAll(){
  renderKpis();
  renderTable();
  renderChart();
}

function setSort(sort){
  currentSort = sort;
  for (const btn of document.querySelectorAll(".segBtn")){
    btn.classList.toggle("active", btn.dataset.sort === sort);
  }
  applyFilters();
}

function exportCsv(){
  if (!viewRows.length) return;

  const total = rawRows.reduce((a,b)=>a+(b.value||0),0) || 1;
  const header = ["순위","비제이명","월별 누적별풍선","점유율"];
  const lines = viewRows.map(r => {
    const share = (r.value||0)/total;
    return [r.rank, r.name, r.value||0, (share*100).toFixed(1)+"%"].join(",");
  });
  const csv = [header.join(","), ...lines].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "YB_ranking.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function toggleTheme(){
  const root = document.documentElement;
  const cur = root.getAttribute("data-theme") || "dark";
  const next = cur === "dark" ? "light" : "dark";
  root.setAttribute("data-theme", next);
  localStorage.setItem("yb_theme", next);
  // update chart grid/tick colors by re-rendering
  if (chart){ chart.destroy(); chart = null; }
  renderChart();
}

function escapeHtml(s){
  return String(s)
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

async function boot(){
  updateTierText();

  // theme restore
  const saved = localStorage.getItem("yb_theme");
  document.documentElement.setAttribute("data-theme", saved || "dark");

  // events
  el("searchInput").addEventListener("input", applyFilters);
  el("topN").addEventListener("change", () => renderChart());
  el("btnExport").addEventListener("click", exportCsv);
  el("btnTheme").addEventListener("click", toggleTheme);

  for (const btn of document.querySelectorAll(".segBtn")){
    btn.addEventListener("click", () => setSort(btn.dataset.sort));
  }

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
      const wb = XLSX.read(ab, { type: "array" });
      rawRows = parseExcelToRows(wb);
      setHint(`업로드 로드 완료: ${file.name}`);
      applyFilters();
    }catch(err){
      setHint("엑셀을 읽지 못했습니다. 파일/시트/컬럼명을 확인해주세요.");
      console.error(err);
    }
  });

  // initial load: try default excel
  try{
    const wb = await loadDefaultXlsx();
    rawRows = parseExcelToRows(wb);
    setHint("data/YB.xlsx 로드 완료");
    setSort("rank");
  }catch(e){
    setHint("data/YB.xlsx 자동 로드 실패 → 엑셀 업로드 버튼을 사용하세요");
    rawRows = [];
    viewRows = [];
    renderAll();
    console.warn(e);
  }
}

boot();
