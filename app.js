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
    { key: "T1", min: 500000, label: "50만+" },
    { key: "T2", min: 300000, label: "30만+" },
    { key: "T3", min: 100000, label: "10만+" },
  ],


  // 점유율 표시(지저분하면 false)
  showShare: false,
};

let rawRows = []; // combined (meta용)
let rawMaleRows = [];
let rawFemaleRows = [];

let viewRows = [];      // 단일 보기(남자/여자)
let viewMaleRows = [];  // 전체 보기(왼쪽)
let viewFemaleRows = []; // 전체 보기(오른쪽)

let currentTab = "all"; // all | male | female

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

function buildSplitHeaderMap(headers){
  // 헤더 예: "남자 순위","남자 비제이명","남자 월별 누적별풍선","여자 순위","여자 비제이명","여자 월별 누적별풍선"
  const norm = headers.map(normalizeHeader);
  const cand = (key) => CONFIG.headerCandidates[key].map(normalizeHeader);

  const pickWithPrefix = (prefix, key) => {
    const p = normalizeHeader(prefix);
    const cs = cand(key);
    // 1) prefix + 후보 포함
    for (let i=0;i<norm.length;i++){
      if (!norm[i]) continue;
      if (!norm[i].includes(p)) continue;
      for (const c of cs){
        if (c && norm[i].includes(c)) return headers[i];
      }
    }
    // 2) 느슨 매칭(부분 포함)
    for (let i=0;i<norm.length;i++){
      if (!norm[i]) continue;
      if (!norm[i].includes(p)) continue;
      for (const c of cs){
        if (c && (c.includes(norm[i]) || norm[i].includes(c))) return headers[i];
      }
    }
    return null;
  };

  const male = {
    rank: pickWithPrefix("남자", "rank"),
    name: pickWithPrefix("남자", "name"),
    value: pickWithPrefix("남자", "value"),
  };
  const female = {
    rank: pickWithPrefix("여자", "rank"),
    name: pickWithPrefix("여자", "name"),
    value: pickWithPrefix("여자", "value"),
  };

  // 새로고침 시간은 공용 1개 컬럼일 가능성이 큼
  const common = buildHeaderMap(headers);
  const refresh = common.refresh || pickWithPrefix("남자","refresh") || pickWithPrefix("여자","refresh");

  const ok = !!(male.name && male.value && female.name && female.value);
  return { ok, male, female, refresh };
}

function normalizeName(v){
  const s = String(v ?? "").trim();
  return s ? s : null;
}
function coerceNumber(v){
  if (v == null) return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const s = String(v).replaceAll(",","").trim();
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}
function coerceRank(v){
  if (v == null) return null;
  const n = Number(String(v).replaceAll(",","").trim());
  return Number.isFinite(n) ? n : null;
}

function finalizeRanks(list){
  // rank가 비어 있으면 value 내림차순으로 새로 계산
  const hasAnyRank = list.some(r => Number.isFinite(r.rank));
  if (!hasAnyRank){
    list.sort((a,b)=> (b.value||0) - (a.value||0));
    list.forEach((r,i)=> r.rank = i+1);
    return list;
  }
  // rank 우선, 보조로 value 내림차순
  list.sort((a,b)=>{
    const ar = Number.isFinite(a.rank) ? a.rank : 1e9;
    const br = Number.isFinite(b.rank) ? b.rank : 1e9;
    if (ar !== br) return ar - br;
    return (b.value||0) - (a.value||0);
  });
  // 결번 채우기
  let next = 1;
  for (const r of list){
    if (!Number.isFinite(r.rank)) r.rank = next;
    next = Math.max(next, r.rank + 1);
  }
  return list;
}

function parseExcelToRows(workbook){
  // 1) 가장 적절한 시트를 찾고(기본 로직 유지) 헤더를 확인
  const { sheetName } = findBestSheet(workbook);
  const ws = workbook.Sheets[sheetName];
  if (!ws) throw new Error("워크시트가 없습니다.");

  const headerRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const headers = (headerRows && headerRows[0]) ? headerRows[0] : [];
  const split = buildSplitHeaderMap(headers);

  // 2) ✅ '남자/여자'가 좌우로 분리된 2개 표 형태면: male/female 각각 파싱
  if (split.ok){
    const json = XLSX.utils.sheet_to_json(ws, { defval: null });
    const male = [];
    const female = [];

    for (const o of json){
      const refreshVal = split.refresh ? o[split.refresh] : null;

      const mn = normalizeName(o[split.male.name]);
      if (mn){
        male.push({
          rank: coerceRank(split.male.rank ? o[split.male.rank] : null),
          name: mn,
          value: coerceNumber(o[split.male.value]),
          refresh: refreshVal ?? null
        });
      }

      const fn = normalizeName(o[split.female.name]);
      if (fn){
        female.push({
          rank: coerceRank(split.female.rank ? o[split.female.rank] : null),
          name: fn,
          value: coerceNumber(o[split.female.value]),
          refresh: refreshVal ?? null
        });
      }
    }

    finalizeRanks(male);
    finalizeRanks(female);

    // combined는 메타(총/평균/인원/갱신) 계산용
    const combined = [...male, ...female];
    return { mode: "split", rows: combined, maleRows: male, femaleRows: female, sheetName };
  }

  // 3) ✅ 기존(단일 표) 파싱: 헤더 자동 인식
  const { map } = findBestSheet(workbook);
  const json = XLSX.utils.sheet_to_json(ws, { defval: null });

  const out = [];
  for (const o of json){
    const name = normalizeName(o[map.name]);
    if (!name) continue;
    out.push({
      rank: coerceRank(map.rank ? o[map.rank] : null),
      name,
      value: coerceNumber(o[map.value]),
      refresh: map.refresh ? o[map.refresh] : null
    });
  }

  finalizeRanks(out);

  return { mode: "single", rows: out, sheetName };
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
function initGenderTabs(){
  const host = document.getElementById("genderTabs");
  if (!host) return;

  host.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-tab]");
    if (!btn) return;
    const tab = btn.getAttribute("data-tab");
    if (!tab) return;
    currentTab = tab;
    // active 표시
    [...host.querySelectorAll("button[data-tab]")].forEach(b => {
      const isOn = b.getAttribute("data-tab") === currentTab;
      b.classList.toggle("isActive", isOn);
      b.setAttribute("aria-selected", isOn ? "true" : "false");
    });
    applyFilters();
  });
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

  const filterAndSort = (src) => {
    let rows = [...(src || [])];
    if (q) rows = rows.filter(r => (r.name||"").toLowerCase().includes(q));

    if (currentSort === "rank"){
      rows.sort((a,b)=> (a.rank||1e9) - (b.rank||1e9));
    } else if (currentSort === "value"){
      rows.sort((a,b)=> (b.value||0) - (a.value||0));
    } else {
      rows.sort((a,b)=> (a.name||"").localeCompare((b.name||""), "ko"));
    }
    return rows;
  };

  if (currentTab === "male"){
    viewRows = filterAndSort(rawMaleRows);
  } else if (currentTab === "female"){
    viewRows = filterAndSort(rawFemaleRows);
  } else {
    viewMaleRows = filterAndSort(rawMaleRows);
    viewFemaleRows = filterAndSort(rawFemaleRows);
    viewRows = []; // 전체는 분할 렌더
  }

  render();
}

function renderListHTML(list, maxVal, totalAll){
  return (list || []).map(r => {
    const t = tierOf(r.value||0);
    const width = Math.max(2, Math.min(100, Math.round((r.value||0) / (maxVal||1) * 100)));
    const share = (r.value||0) / (totalAll||1);

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
}

function render(){
  const rowsEl = el("rows");

  const setMetaEmpty = () => {
    const tm0 = document.getElementById("tableMeta"); if (tm0) tm0.textContent = "-";
    el("metaCount").textContent = "-";
    el("metaAvg").textContent = "-";
    el("metaRefresh").textContent = "-";
    const mt0 = document.getElementById("metaTotal"); if (mt0) mt0.textContent = "-";
  };

  // ✅ 전체(남/여 2분할) 보기
  if (currentTab === "all"){
    const male = viewMaleRows || [];
    const female = viewFemaleRows || [];

    const hasAny = male.length || female.length;
    if (!hasAny){
      rowsEl.innerHTML = `<div class="empty">표시할 데이터가 없습니다.</div>`;
      setMetaEmpty();
      return;
    }

    const totalMale = male.reduce((a,b)=>a+(b.value||0),0) || 1;
    const totalFemale = female.reduce((a,b)=>a+(b.value||0),0) || 1;
    const maxMale = Math.max(...male.map(r=>r.value||0), 1);
    const maxFemale = Math.max(...female.map(r=>r.value||0), 1);

    rowsEl.innerHTML = `
      <div class="splitGrid">
        <div class="splitCol">
          <div class="splitHead">남자</div>
          <div class="splitRows">
            ${male.length ? renderListHTML(male, maxMale, totalMale) : `<div class="empty small">남자 데이터 없음</div>`}
          </div>
        </div>
        <div class="splitCol">
          <div class="splitHead">여자</div>
          <div class="splitRows">
            ${female.length ? renderListHTML(female, maxFemale, totalFemale) : `<div class="empty small">여자 데이터 없음</div>`}
          </div>
        </div>
      </div>
    `;

    // 메타는 combined 기준
    const combined = [...(rawMaleRows||[]), ...(rawFemaleRows||[])];
    const { avg, count, refresh } = computeMeta(combined);
    const total = combined.reduce((a,b)=>a+(b.value||0),0);
    const mt = document.getElementById("metaTotal"); if (mt) mt.textContent = fmt.format(total);
    el("metaCount").textContent = fmt.format(count) + "명";
    el("metaAvg").textContent = fmt.format(avg);
    el("metaRefresh").textContent = formatRefresh(refresh);
    const tm = document.getElementById("tableMeta"); if (tm) tm.textContent = `남자 ${male.length}명 · 여자 ${female.length}명 · 갱신 ${formatRefresh(refresh)}`;
    return;
  }

  // ✅ 단일(남자/여자) 보기
  if (!viewRows.length){
    rowsEl.innerHTML = `<div class="empty">표시할 데이터가 없습니다.</div>`;
    setMetaEmpty();
    return;
  }

  const activeRaw = (currentTab === "male") ? (rawMaleRows||[]) : (rawFemaleRows||[]);
  const totalAll = activeRaw.reduce((a,b)=>a+(b.value||0),0) || 1;
  const maxVal = Math.max(...activeRaw.map(r=>r.value||0), 1);

  rowsEl.innerHTML = renderListHTML(viewRows, maxVal, totalAll);

  const { avg, count, refresh } = computeMeta(activeRaw);
  const total = activeRaw.reduce((a,b)=>a+(b.value||0),0);
  const mt = document.getElementById("metaTotal"); if (mt) mt.textContent = fmt.format(total);
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

  rawRows = parsed.rows || [];
  rawMaleRows = parsed.maleRows || [];
  rawFemaleRows = parsed.femaleRows || [];

  setHint(`${label} 로드 완료 (시트: ${parsed.sheetName})`);
  applyFilters();
}

async function boot(){
  initGenderTabs();

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
