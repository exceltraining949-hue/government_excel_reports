/* ============ GovData Analytics Portal — Frontend Logic ============ */
"use strict";

/* ---------------- i18n ---------------- */
const I18N = {
  en: {
    tagline: "Government Data Analytics Portal — upload Excel/CSV, the AI Agent does the complete analysis",
    step1: "Upload File", step2: "Auto-Inspect", step3: "Clean & Analyze", step4: "Report & Dashboard",
    uploadTitle: "Upload Your Data File",
    uploadHint: "Excel (.xlsx, .xls, .xlsm) or CSV — max 25 MB. Your raw data is never modified.",
    dzMain: "Drag & drop file here", dzSub: "or browse using the button below",
    browse: "Browse Files",
    opt1: "Standardize text case & spacing", opt2: 'Mark blank categories as "MISSING"',
    opt3: "Remove exact duplicate rows (from Cleaned sheet)",
    sheetLabel: "Worksheet:", rerun: "Re-analyze",
    analyzeBtn: "⚡ Run AI Agent Analysis",
    ps1: "Inspecting workbook…", ps2: "Data quality checks (duplicates, missing, errors)…",
    ps3: "Cleaning & standardization…", ps4: "Summaries, rankings & anomaly detection…",
    ps5: "Generating Excel report & dashboard…",
    reportReady: "✅ Government Analytics Report is ready",
    download: "Download Excel Report",
    agentTitle: "AI Agent Command Box",
    agentHint: 'Type in Roman Urdu or English — e.g. "top 10 dikhao", "department wise summary", "million me dikhao", "duplicates dikhao"',
    cmdGo: "Run",
    numberFormat: "Numbers:",
    fmtFull: "Full",
    tabQuality: "Data Quality", tabSummary: "Summary & Pivots", tabRank: "Top / Bottom 50",
    tabExc: "Exceptions", tabChanges: "Change Log", tabAssume: "Assumptions", tabCols: "Columns",
    tabPivot: "Pivot Builder", tabModel: "Data Model", tabValid: "Validation", tabPQ: "Power Query",
    tabDash: "Dashboard", tabLookup: "Lookup (V/X)", tabPivotCharts: "Pivot Charts",
    footer1: "🔐 Your file stays confidential — analysis happens only on this server, nothing is sent elsewhere.",
    footer2: "GovData Analytics Portal · Automated Excel analysis for government offices · No data is fabricated — every figure is computed from the uploaded file."
  },
  ur: {
    tagline: "حکومتی ڈیٹا تجزیہ پورٹل — ایکسل/سی ایس وی اپلوڈ کریں، اے آئی ایجنٹ مکمل تجزیہ خود کرے گا",
    step1: "فائل اپلوڈ", step2: "معائنہ", step3: "صفائی و تجزیہ", step4: "رپورٹ و ڈیش بورڈ",
    uploadTitle: "اپنی ڈیٹا فائل اپلوڈ کریں",
    uploadHint: "ایکسل (.xlsx, .xls, .xlsm) یا سی ایس وی — زیادہ سے زیادہ 25 ایم بی۔ اصل ڈیٹا میں کوئی تبدیلی نہیں ہوگی۔",
    dzMain: "فائل یہاں چھوڑیں", dzSub: "یا نیچے بٹن سے منتخب کریں",
    browse: "فائل منتخب کریں",
    opt1: "متن کی ہجے اور خلا کی اصلاح", opt2: 'خالی اقسام کو "MISSING" لکھیں',
    opt3: "یکساں ڈپلیکیٹ قطاریں صاف کریں (صرف کلیئنڈ شیٹ سے)",
    sheetLabel: "شیٹ:", rerun: "دوبارہ تجزیہ",
    analyzeBtn: "⚡ اے آئی ایجنٹ تجزیہ شروع کریں",
    ps1: "ورک بک کا معائنہ ہو رہا ہے…", ps2: "ڈیٹا کوالٹی جانچ (ڈپلیکیٹ، غائب، غلطیاں)…",
    ps3: "صفائی اور معیار بندی…", ps4: "خلاصے، درجہ بندی اور غیر معمولی چیزیں…",
    ps5: "ایکسل رپورٹ اور ڈیش بورڈ بن رہا ہے…",
    reportReady: "✅ حکومتی تجزیاتی رپورٹ تیار ہے",
    download: "ایکسل رپورٹ ڈاؤن لوڈ کریں",
    agentTitle: "اے آئی ایجنٹ کمانڈ باکس",
    agentHint: "رومن اردو یا انگریزی میں لکھیں — مثلاً ”ٹاپ 10 دکھاؤ“، ”محکمہ وار خلاصہ“، ”ملین میں دکھاؤ“",
    cmdGo: "چلائیں",
    numberFormat: "اعداد:",
    fmtFull: "مکمل",
    tabQuality: "ڈیٹا کوالٹی", tabSummary: "خلاصہ و تجزیہ", tabRank: "ٹاپ / بٹم 50",
    tabExc: "استثنائات", tabChanges: "تبدیلی رپورٹ", tabAssume: "مفروضات", tabCols: "کالم",
    tabPivot: "پیوٹ بلڈر", tabModel: "ڈیٹا ماڈل", tabValid: "ویلیڈیشن", tabPQ: "پاور کوری",
    tabDash: "ڈیش بورڈ", tabLookup: "لک اپ", tabPivotCharts: "پیوٹ چارٹس",
    footer1: "🔐 آپ کی فائل رازدارانہ ہے — تجزیہ صرف اسی سرور پر ہوتا ہے۔",
    footer2: "گوو ڈیٹا اینالیٹکس پورٹل · حکومتی دفاتر کے لیے خودکار ایکسل تجزیہ · کوئی ڈیٹا گھڑا نہیں جاتا۔"
  }
};
let LANG = "en";
function setLang(l) {
  LANG = l;
  document.getElementById("langEN").classList.toggle("active", l === "en");
  document.getElementById("langUR").classList.toggle("active", l === "ur");
  document.body.classList.toggle("urdu", l === "ur");
  document.querySelectorAll("[data-i18n]").forEach(el => {
    const k = el.getAttribute("data-i18n");
    if (I18N[l][k]) el.textContent = I18N[l][k];
  });
}
window.setLang = setLang;

/* ---------------- state ---------------- */
let PAYLOAD = null;
let NUMFMT = "full";
let CURRENT_TAB = "dashboard";

/* ---------------- helpers ---------------- */
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

function fmtNum(v, opts) {
  if (v == null || v === "" || isNaN(v)) return v == null || v === "" ? "—" : esc(v);
  const n = Number(v);
  if (NUMFMT === "full") return n.toLocaleString("en-US", { maximumFractionDigits: 2 });
  const abs = Math.abs(n);
  if (NUMFMT === "k") return (n / 1e3).toLocaleString("en-US", { maximumFractionDigits: 1 }) + "K";
  if (NUMFMT === "m") return (n / 1e6).toLocaleString("en-US", { maximumFractionDigits: 2 }) + "M";
  if (NUMFMT === "b") return (n / 1e9).toLocaleString("en-US", { maximumFractionDigits: 2 }) + "B";
  return String(n);
}
function fmtPct(v) { return v == null ? "—" : Number(v).toLocaleString("en-US", { maximumFractionDigits: 1 }) + "%"; }

/* ---------------- upload UI ---------------- */
const dropzone = $("dropzone"), fileInput = $("fileInput"), analyzeBtn = $("analyzeBtn");
let selectedFile = null;

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") fileInput.click(); });
["dragenter", "dragover"].forEach(ev => dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.classList.add("drag"); }));
["dragleave", "drop"].forEach(ev => dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.classList.remove("drag"); }));
dropzone.addEventListener("drop", e => { if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]); });
fileInput.addEventListener("change", () => { if (fileInput.files.length) setFile(fileInput.files[0]); });
$("browseBtn").addEventListener("click", e => { e.stopPropagation(); fileInput.click(); });
$("chipRemove").addEventListener("click", () => {
  selectedFile = null; fileInput.value = ""; $("fileChip").hidden = true; analyzeBtn.disabled = true;
});

function setFile(f) {
  const ext = f.name.toLowerCase().split(".").pop();
  if (!["xlsx", "xls", "xlsm", "csv"].includes(ext)) { showError("Unsupported file type. Please upload .xlsx, .xls, .xlsm or .csv."); return; }
  if (f.size > 25 * 1024 * 1024) { showError("File is larger than 25 MB."); return; }
  selectedFile = f;
  $("chipName").textContent = f.name;
  $("chipSize").textContent = (f.size / 1024).toFixed(1) + " KB · " + ext.toUpperCase();
  $("fileChip").hidden = false;
  analyzeBtn.disabled = false;
  hideError();
}
function showError(msg) { const b = $("errorBox"); b.textContent = "⚠ " + msg; b.hidden = false; }
function hideError() { $("errorBox").hidden = true; }

/* ---------------- progress animation ---------------- */
const STEP_MSGS = ["ps1", "ps2", "ps3", "ps4", "ps5"];
let progressTimer = null;
function startProgress() {
  const steps = $("progressSteps").querySelectorAll("li");
  steps.forEach(li => li.className = "");
  $("progressWrap").hidden = false;
  let p = 8;
  $("progressFill").style.width = p + "%";
  progressTimer = setInterval(() => {
    p = Math.min(p + Math.random() * 9, 92);
    $("progressFill").style.width = p + "%";
    const activeIdx = Math.min(Math.floor(p / 20), 4);
    steps.forEach((li, i) => li.className = i < activeIdx ? "done" : (i === activeIdx ? "active" : ""));
    $("progressMsg").textContent = I18N[LANG][STEP_MSGS[activeIdx]] || STEP_MSGS[activeIdx];
  }, 700);
  setStep(1); setStep(2);
}
function finishProgress(ok) {
  clearInterval(progressTimer);
  if (ok) {
    $("progressFill").style.width = "100%";
    $("progressSteps").querySelectorAll("li").forEach(li => li.className = "done");
    setTimeout(() => { $("progressWrap").hidden = true; }, 900);
    [1, 2, 3, 4].forEach(i => setStep(i, true));
  } else {
    $("progressWrap").hidden = true;
    [1, 2, 3, 4].forEach(i => setStep(i, false, true));
  }
}
function setStep(n, done, reset) {
  const el = document.querySelector(`.step[data-step="${n}"]`);
  if (!el) return;
  el.classList.remove("active", "done");
  if (reset) return;
  el.classList.add(done ? "done" : "active");
  if (n > 1 && !done) { document.querySelectorAll(".step").forEach(s => { const sn = +s.dataset.step; if (sn < n) s.classList.add("done"); }); }
}

/* ---------------- analyze ---------------- */
analyzeBtn.addEventListener("click", runAnalysis);
async function runAnalysis() {
  if (!selectedFile) return;
  hideError();
  analyzeBtn.disabled = true;
  $("results").hidden = true;
  startProgress();
  const fd = new FormData();
  fd.append("file", selectedFile);
  fd.append("standardize_case", $("optStandardize").checked);
  fd.append("fill_missing", $("optFill").checked);
  fd.append("remove_duplicates", $("optDedup").checked);
  try {
    const res = await fetch("/api/analyze", { method: "POST", body: fd });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || "Analysis failed.");
    PAYLOAD = data;
    finishProgress(true);
    renderResults();
    $("results").hidden = false;
    $("results").scrollIntoView({ behavior: "smooth" });
  } catch (err) {
    finishProgress(false);
    showError(err.message || "Something went wrong while processing the file.");
  } finally {
    analyzeBtn.disabled = false;
  }
}

$("rerunBtn").addEventListener("click", async () => {
  if (!PAYLOAD) return;
  hideError();
  startProgress();
  try {
    const res = await fetch("/api/reanalyze", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        job_id: PAYLOAD.job_id, sheet: $("sheetSelect").value,
        standardize_case: $("optStandardize").checked,
        fill_missing: $("optFill").checked,
        remove_duplicates: $("optDedup").checked,
      })
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || "Re-analysis failed.");
    PAYLOAD = data;
    finishProgress(true);
    renderResults();
  } catch (err) {
    finishProgress(false);
    showError(err.message);
  }
});

/* ---------------- render results ---------------- */
function renderResults() {
  const p = PAYLOAD;
  $("downloadBtn").href = p.download_url;
  $("reportMeta").textContent = `${p.file_name} → sheet "${p.sheet_name}" · ${p.kpis.total_records.toLocaleString()} records × ${p.kpis.columns} columns · analyzed ${new Date().toLocaleString()}`;
  renderSheetPicker();
  renderTab(CURRENT_TAB);
  $("fmtBar").style.display = "flex";
}

function renderSheetPicker() {
  const p = PAYLOAD;
  const sel = $("sheetSelect");
  sel.innerHTML = "";
  (p.sheets_info || []).forEach(s => {
    const o = document.createElement("option");
    o.value = s.name;
    o.textContent = `${s.name} (${s.rows.toLocaleString()}×${s.cols})${s.hidden ? " — hidden" : ""}`;
    if (s.name === p.sheet_name) o.selected = true;
    sel.appendChild(o);
  });
  $("sheetPick").hidden = (p.sheets_info || []).length < 2;
  const others = (p.sheets_info || []).filter(s => s.name !== p.sheet_name);
  $("sheetsNote").textContent = others.length
    ? `Workbook me ${others.length + 1} sheets hain. "${p.sheet_name}" analyze hui hai — upar dropdown se koi aur sheet chunein aur "Re-analyze" dabaein.`
    : "";
}

function renderKpis() {
  const k = PAYLOAD.kpis;
  const grid = $("kpiGrid");
  const cards = [];
  cards.push({ label: "Total Records", value: k.total_records.toLocaleString(), sub: `${k.columns} columns`, cls: "" });
  cards.push({ label: "Duplicate Rows", value: k.duplicate_rows.toLocaleString(), sub: k.duplicate_rows ? "exact copies found — see Exceptions" : "none found", cls: k.duplicate_rows ? "bad" : "good" });
  cards.push({ label: "Missing Cells", value: k.missing_cells.toLocaleString(), sub: `${k.missing_pct}% of all cells`, cls: k.missing_cells ? "warn" : "good" });
  if (k.utilization != null) {
    cards.push({ label: "Budget Utilization", value: fmtPct(k.utilization), sub: `${fmtNum(k.total_expenditure)} of ${fmtNum(k.total_budget)}`, cls: k.utilization > 100 ? "bad" : (k.utilization > 90 ? "warn" : "good") });
  }
  if (k.main_metric && k.main_metric_total != null) {
    cards.push({ label: `Total ${k.main_metric}`, value: fmtNum(k.main_metric_total), sub: "sum of all records", cls: "" });
  }
  if (k.date_range) cards.push({ label: "Date Range", value: k.date_range[0], sub: `to ${k.date_range[1]}`, cls: "" });
  if (k.top_category) cards.push({ label: `Top ${k.top_category.column}`, value: esc(k.top_category.label), sub: `${k.top_category.count.toLocaleString()} records`, cls: "" });
  cards.push({ label: "Anomalies (IQR)", value: k.outliers.toLocaleString(), sub: "flagged for review only", cls: k.outliers ? "warn" : "good" });
  grid.innerHTML = cards.map(c => `
    <div class="kpi ${c.cls}">
      <div class="kpi-label">${esc(c.label)}</div>
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-sub">${esc(c.sub)}</div>
    </div>`).join("");
}

/* ---------------- charts (pure SVG/CSS) ---------------- */
const PALETTE = ["#0F4C3A", "#C9A227", "#2E7D32", "#7B1FA2", "#00838F", "#B7791F", "#C0392B", "#455A64", "#6D4C41", "#0288D1"];

function renderCharts() {
  const p = PAYLOAD;
  const grid = $("chartsGrid");
  grid.innerHTML = "";
  if (p.pivots && p.pivots.length) {
    const piv = p.pivots[0];
    const rows = piv.rows.slice(0, 10);
    const max = Math.max(...rows.map(r => Number(r.sum ?? r.count) || 0), 1e-9);
    const bars = rows.map((r, i) => {
      const v = Number(r.sum ?? r.count) || 0;
      return `<div class="bar-row">
        <div class="bar-label" title="${esc(r.label)}">${esc(r.label)}</div>
        <div class="bar-track"><div class="bar-fill${i === 0 ? " gold" : ""}" style="width:${Math.max(v / max * 100, 2)}%"></div>
        <span class="bar-val">${piv.has_measure ? fmtNum(v) : v.toLocaleString()}</span></div>
      </div>`;
    }).join("");
    grid.innerHTML += `<div class="chart-card">
      <h4>${piv.has_measure ? `${esc(piv.measure)} by ${esc(piv.dimension)}` : `Records by ${esc(piv.dimension)}`}</h4>
      <div class="chart-sub">${piv.has_measure ? "Top 10 by sum" : "Top categories"} · source: ${esc(piv.dimension)}</div>
      ${bars}</div>`;
  }
  if (p.status_dist) {
    const rows = p.status_dist.rows;
    const total = rows.reduce((a, r) => a + r.count, 0) || 1;
    grid.innerHTML += donutCard(`${p.status_dist.column} distribution`, rows, total);
  }
  if (p.trend && p.trend.labels && p.trend.labels.length >= 2) {
    grid.innerHTML += lineCard(`Monthly trend — ${p.trend.measure || "records"}`, p.trend.labels, p.trend.values);
  }
  if (p.budget && p.budget.per_dept && p.budget.per_dept.length) {
    const rows = p.budget.per_dept.slice(0, 10).filter(d => d.budget != null || d.expenditure != null);
    const max = Math.max(...rows.flatMap(d => [d.budget || 0, d.expenditure || 0]), 1e-9);
    const bars = rows.map(d => `
      <div class="bar-row">
        <div class="bar-label" title="${esc(d.label)}">${esc(d.label)}</div>
        <div class="bar-track">
          <div class="bar-fill" style="width:${Math.max((d.budget || 0) / max * 100, 2)}%"></div>
          <span class="bar-val">B ${fmtNum(d.budget)}</span>
        </div>
      </div>
      <div class="bar-row">
        <div class="bar-label"></div>
        <div class="bar-track">
          <div class="bar-fill gold" style="width:${Math.max((d.expenditure || 0) / max * 100, 2)}%"></div>
          <span class="bar-val">E ${fmtNum(d.expenditure)} · ${fmtPct(d.utilization)}</span>
        </div>
      </div>`).join("");
    grid.innerHTML += `<div class="chart-card">
      <h4>Budget vs Expenditure by department</h4>
      <div class="chart-sub">Green bar = Budget (B) · Gold bar = Expenditure (E) with utilization %</div>
      ${bars}</div>`;
  }
  if (!grid.innerHTML) grid.innerHTML = '<div class="chart-card"><div class="empty">Charts not possible with this data — numeric or category columns required.</div></div>';
}

function donutCard(title, rows, total) {
  let angle = -90;
  const segs = rows.map((r, i) => {
    const frac = r.count / total;
    let a0 = angle, a1 = angle + frac * 360;
    // single full-circle segment (100%) breaks SVG arc rendering - cap it
    if (a1 - a0 >= 359.99) a1 = a0 + 359.99;
    angle = a1;
    return { seg: arcPath(80, 80, 55, 60, a0, a1), color: PALETTE[i % PALETTE.length], label: r.label, count: r.count, pct: (frac * 100).toFixed(1) };
  });
  return `<div class="chart-card">
    <h4>${esc(title)}</h4>
    <div class="chart-sub">${total.toLocaleString()} records</div>
    <div class="donut-wrap">
      <svg class="chart-svg" viewBox="0 0 160 160" style="max-width:170px">
        ${segs.map(s => `<path d="${s.seg}" fill="${s.color}" stroke="#fff" stroke-width="1.5"><title>${esc(s.label)}: ${s.count.toLocaleString()} (${s.pct}%)</title></path>`).join("")}
        <circle cx="80" cy="80" r="33" fill="#fff"/>
        <text x="80" y="76" text-anchor="middle" font-size="13" font-weight="700" fill="#0B3D2E">${total.toLocaleString()}</text>
        <text x="80" y="90" text-anchor="middle" font-size="9" fill="#5E6E66">records</text>
      </svg>
      <div class="legend">
        ${segs.map(s => `<div class="legend-item"><span class="legend-dot" style="background:${s.color}"></span>${esc(s.label)} — <b>${s.count.toLocaleString()}</b> (${s.pct}%)</div>`).join("")}
      </div>
    </div></div>`;
}
function arcPath(cx, cy, r0, r1, a0, a1) {
  const rad = a => (a * Math.PI) / 180;
  const x0 = cx + r1 * Math.cos(rad(a0)), y0 = cy + r1 * Math.sin(rad(a0));
  const x1 = cx + r1 * Math.cos(rad(a1)), y1 = cy + r1 * Math.sin(rad(a1));
  const x2 = cx + r0 * Math.cos(rad(a1)), y2 = cy + r0 * Math.sin(rad(a1));
  const x3 = cx + r0 * Math.cos(rad(a0)), y3 = cy + r0 * Math.sin(rad(a0));
  const large = a1 - a0 > 180 ? 1 : 0;
  return `M${x0} ${y0} A${r1} ${r1} 0 ${large} 1 ${x1} ${y1} L${x2} ${y2} A${r0} ${r0} 0 ${large} 0 ${x3} ${y3} Z`;
}
function lineCard(title, labels, values) {
  const w = 460, h = 200, padL = 46, padR = 10, padT = 14, padB = 42;
  const vals = values.map(v => Number(v) || 0);
  const min = Math.min(...vals), max = Math.max(...vals);
  const span = max - min || 1;
  const x = i => padL + (i * (w - padL - padR)) / Math.max(labels.length - 1, 1);
  const y = v => padT + (1 - (v - min) / span) * (h - padT - padB);
  const pts = vals.map((v, i) => `${x(i).toFixed(1)},${y(v).toFixed(1)}`).join(" ");
  const area = `${padL},${y(min)} ` + pts + ` ${x(vals.length - 1)},${y(min)}`;
  const gridLines = [0, 0.25, 0.5, 0.75, 1].map(f => {
    const gv = min + f * span;
    return `<line x1="${padL}" y1="${y(gv)}" x2="${w - padR}" y2="${y(gv)}" stroke="#E3EAE6" stroke-width="1"/>
      <text x="${padL - 6}" y="${y(gv) + 3}" text-anchor="end" font-size="9" fill="#5E6E66">${fmtNum(Math.round(gv))}</text>`;
  }).join("");
  const labStep = Math.ceil(labels.length / 10);
  const xLabels = labels.map((l, i) => i % labStep === 0
    ? `<text x="${x(i)}" y="${h - 22}" text-anchor="end" font-size="8.5" fill="#5E6E66" transform="rotate(-40 ${x(i)} ${h - 22})">${esc(l)}</text>` : "").join("");
  const dots = vals.map((v, i) => `<circle cx="${x(i)}" cy="${y(v)}" r="3" fill="#0F4C3A"><title>${esc(labels[i])}: ${fmtNum(v)}</title></circle>`).join("");
  return `<div class="chart-card" style="grid-column:1/-1">
    <h4>${esc(title)}</h4>
    <div class="chart-sub">last ${labels.length} months · by "${esc(PAYLOAD.trend.date_col)}"</div>
    <svg class="chart-svg" viewBox="0 0 ${w} ${h}">
      ${gridLines}
      <polygon points="${area}" fill="rgba(15,76,58,.08)"/>
      <polyline points="${pts}" fill="none" stroke="#0F4C3A" stroke-width="2.2" stroke-linejoin="round"/>
      ${dots}${xLabels}
    </svg></div>`;
}

/* ---------------- tabs ---------------- */
document.querySelectorAll(".tab").forEach(t => t.addEventListener("click", () => {
  document.querySelectorAll(".tab").forEach(x => x.classList.remove("active"));
  t.classList.add("active");
  CURRENT_TAB = t.dataset.tab;
  renderTab(CURRENT_TAB);
}));

function tableHTML(headers, rows, opts) {
  opts = opts || {};
  if (!rows.length) return '<div class="empty">No records to show.</div>';
  const ths = headers.map(h => `<th class="${h.num ? "num" : ""}">${esc(h.label)}</th>`).join("");
  const trs = rows.map(r => `<tr>${r.map((cell, i) => `<td class="${headers[i] && headers[i].num ? "num" : ""}">${cell == null ? "—" : cell}</td>`).join("")}</tr>`).join("");
  return `<div class="table-wrap"><table class="data"><thead><tr>${ths}</tr></thead><tbody>${trs}</tbody></table></div>`;
}

function renderTab(tab) {
  const p = PAYLOAD;
  const body = $("tabBody");
  if (!p) return;
  if (tab === "quality") body.innerHTML = tabQuality(p);
  else if (tab === "summary") body.innerHTML = tabSummary(p);
  else if (tab === "rank") body.innerHTML = tabRank(p);
  else if (tab === "exceptions") body.innerHTML = tabExceptions(p);
  else if (tab === "changes") body.innerHTML = tabChanges(p);
  else if (tab === "assume") body.innerHTML = tabAssume(p);
  else if (tab === "columns") body.innerHTML = tabColumns(p);
  else if (tab === "pivot") { body.innerHTML = tabPivotBuilder(p); wirePivotBuilder(); }
  else if (tab === "pivotcharts") { body.innerHTML = tabPivotCharts(p); wirePivotCharts(); }
  else if (tab === "dashboard") { body.innerHTML = tabDashboard(p); renderKpis(); renderCharts(); }
  else if (tab === "lookup") { body.innerHTML = tabLookup(p); wireLookup(); }
  else if (tab === "model") { body.innerHTML = tabModel(p); wireModel(); }
  else if (tab === "valid") body.innerHTML = tabValidation(p);
  else if (tab === "pq") { body.innerHTML = tabPowerQuery(p); wirePowerQuery(); }
}

function sevBadge(s) {
  const cls = s === "CRITICAL" ? "crit" : (s === "WARNING" ? "warn" : "info");
  return `<span class="badge ${cls}">${esc(s)}</span>`;
}

function tabQuality(p) {
  const rows = (p.issues || []).map(i => [
    `<span class="sev-cell ${i.severity === "CRITICAL" ? "crit" : i.severity === "WARNING" ? "warn" : "info"}">${esc(i.severity)}</span>`,
    esc(i.area), esc(i.description), i.count == null ? "—" : Number(i.count).toLocaleString()
  ]);
  const counts = { CRITICAL: 0, WARNING: 0, INFO: 0 };
  (p.issues || []).forEach(i => counts[i.severity] = (counts[i.severity] || 0) + 1);
  return `
    <div class="kpi-grid" style="margin-bottom:16px">
      <div class="kpi bad"><div class="kpi-label">Critical</div><div class="kpi-value">${counts.CRITICAL}</div><div class="kpi-sub">must be reviewed</div></div>
      <div class="kpi warn"><div class="kpi-label">Warnings</div><div class="kpi-value">${counts.WARNING}</div><div class="kpi-sub">attention required</div></div>
      <div class="kpi good"><div class="kpi-label">Info / Notes</div><div class="kpi-value">${counts.INFO}</div><div class="kpi-sub">for information</div></div>
    </div>
    ${tableHTML([{ label: "Severity" }, { label: "Area" }, { label: "Issue" }, { label: "Count", num: true }], rows)}`;
}

function tabSummary(p) {
  let html = "";
  if (p.stats && p.stats.length) {
    html += `<div class="section-h">📊 Descriptive Statistics</div>`;
    html += tableHTML(
      [{ label: "Field" }, { label: "Count", num: true }, { label: "Sum", num: true }, { label: "Mean", num: true },
       { label: "Median", num: true }, { label: "Min", num: true }, { label: "Max", num: true }, { label: "Std Dev", num: true }],
      p.stats.map(s => [esc(s.column), s.count.toLocaleString(), fmtNum(s.sum), fmtNum(s.mean), fmtNum(s.median), fmtNum(s.min), fmtNum(s.max), fmtNum(s.std)])
    );
  }
  if (p.budget) {
    html += `<div class="section-h">💰 Budget vs Expenditure</div>
      <p class="muted small">Utilization % = Expenditure ÷ Budget × 100 · Variance = Budget − Expenditure. Budget and expenditure are treated as distinct concepts.</p>`;
    html += tableHTML(
      [{ label: "Department" }, { label: "Records", num: true }, { label: "Budget", num: true },
       { label: "Expenditure", num: true }, { label: "Remaining", num: true }, { label: "Utilization %", num: true }],
      p.budget.per_dept.map(d => [esc(d.label), d.records.toLocaleString(), fmtNum(d.budget), fmtNum(d.expenditure), fmtNum(d.remaining), fmtPct(d.utilization)])
    );
  }
  (p.pivots || []).forEach(pv => {
    html += `<div class="section-h">📌 ${esc(pv.dimension)}-wise summary${pv.has_measure ? ` — ${esc(pv.measure)}` : ""}</div>`;
    if (pv.has_measure) {
      html += tableHTML(
        [{ label: pv.dimension }, { label: "Records", num: true }, { label: "Sum", num: true }, { label: "Average", num: true },
         { label: "Min", num: true }, { label: "Max", num: true }, { label: "Median", num: true }, { label: "Share %", num: true }],
        pv.rows.map(r => [esc(r.label), r.count.toLocaleString(), fmtNum(r.sum), fmtNum(r.avg), fmtNum(r.min), fmtNum(r.max), fmtNum(r.median), fmtPct(r.pct)])
      );
    } else {
      html += tableHTML(
        [{ label: pv.dimension }, { label: "Records", num: true }, { label: "Share %", num: true }],
        pv.rows.map(r => [esc(r.label), r.count.toLocaleString(), fmtPct(r.pct)])
      );
    }
  });
  if (!html) html = '<div class="empty">No summary tables available for this data.</div>';
  return html;
}

function tabRank(p) {
  let html = "";
  const mk = (title, obj) => {
    if (!obj || !obj.rows.length) return "";
    const headers = [{ label: "Rank", num: true }].concat(obj.columns.map(c => ({ label: c, num: c === PAYLOAD.kpis.main_metric || c === PAYLOAD.trend?.measure })));
    const rows = obj.rows.map(r => [r.rank].concat(obj.columns.map(c => {
      const v = r[c];
      if (typeof v === "number") return fmtNum(v);
      return esc(v);
    })));
    return `<div class="section-h">🏆 ${title}</div>` + tableHTML(headers, rows);
  };
  html += mk(`Top 50 by ${p.kpis.main_metric} (highest first)`, p.top50);
  html += mk(`Bottom 50 by ${p.kpis.main_metric} (lowest first)`, p.bottom50);
  if (!html) html = '<div class="empty">Ranking not possible — no numeric amount column detected.</div>';
  return html;
}

function tabExceptions(p) {
  let html = "";
  html += `<div class="section-h">🔁 Exact duplicate records (${p.dup_rows})</div>`;
  if (p.dup_records && p.dup_records.length) {
    const cols = Object.keys(p.dup_records[0]);
    html += tableHTML(cols.map(c => ({ label: c === "row" ? "Row (Raw_Data)" : c, num: c === "row" })),
      p.dup_records.map(d => cols.map(c => d[c] === "" || d[c] == null ? "—" : esc(d[c]))));
  } else html += '<div class="empty">No exact duplicate rows found.</div>';
  if (p.id_dup) {
    html += `<div class="section-h">🔑 Duplicate key values — "${esc(p.id_dup.column)}"</div>`;
    html += tableHTML([{ label: "Value" }, { label: "Times repeated", num: true }],
      p.id_dup.top.map(t => [esc(t.value), t.times.toLocaleString()]));
  }
  html += `<div class="section-h">⚠ Statistical anomalies — IQR method (${p.outlier_total} values)</div>
    <p class="muted small">Har flag ka matlab hai "review required" — ye kisi ghalti ya misconduct ka saboot nahi.</p>`;
  if (p.outliers && p.outliers.length) {
    html += tableHTML([{ label: "Row", num: true }, { label: "Field" }, { label: "Value", num: true }, { label: "Review note" }],
      p.outliers.map(o => [o.row, esc(o.column), fmtNum(o.value), esc(o.detail)]));
  } else html += '<div class="empty">No anomalies detected by the IQR method.</div>';
  if (p.budget && p.budget.above_budget_rows && p.budget.above_budget_rows.length) {
    html += `<div class="section-h">🚩 Expenditure exceeds budget (${p.budget.above_budget_count} records)</div>`;
    html += tableHTML([{ label: "Row", num: true }, { label: "Budget", num: true }, { label: "Expenditure", num: true }, { label: "Review note" }],
      p.budget.above_budget_rows.map(a => [a.row, fmtNum(a.budget), fmtNum(a.expenditure), esc(a.detail)]));
  }
  return html;
}

function tabChanges(p) {
  if (!p.change_log || !p.change_log.length) return '<div class="empty">Koi tabdeeli zaroori nahi thi — data pehle se saaf tha. Raw_Data bilkul wahi hai jo upload ki gayi thi.</div>';
  return `<p class="muted small">Sirf Cleaned_Data me tabdeeli hui hai. Raw_Data (original) me kuch nahi badla. Excel report me poori detail hai.</p>` +
    tableHTML([{ label: "Field" }, { label: "Original" }, { label: "Updated" }, { label: "Reason" }, { label: "Cells", num: true }],
      p.change_log.map(c => [esc(c.field), esc(c.original), esc(c.updated), esc(c.reason), c.cells.toLocaleString()]));
}

function tabAssume(p) {
  if (!p.assumptions || !p.assumptions.length) return '<div class="empty">No assumptions were required.</div>';
  return `<ol class="assume-list">${p.assumptions.map(a => `<li>${esc(a)}</li>`).join("")}</ol>`;
}

function tabColumns(p) {
  return tableHTML(
    [{ label: "Column" }, { label: "Detected role" }, { label: "Type" }, { label: "Unique", num: true },
     { label: "Missing", num: true }, { label: "Missing %", num: true }, { label: "Min / Max or samples" }],
    (p.columns || []).map(c => [esc(c.name), esc(c.role), esc(c.dtype), c.unique.toLocaleString(),
      c.missing.toLocaleString(), fmtPct(c.missing_pct),
      c.sample && c.sample.length ? c.sample.map(esc).join(", ") : (c.min != null ? `${fmtNum(c.min)} … ${fmtNum(c.max)}` : "—")])
  );
}

/* ================= PIVOT BUILDER ================= */
let PIVOT_STATE = { row: null, col: "", measure: "__count__", agg: "sum", filter_col: "", filter_val: "" };

function pivotGridHtml(p, btnLabel) {
  const cats = p.categorical_columns || [];
  const nums = p.numeric_columns || [];
  PIVOT_STATE.row = PIVOT_STATE.row && cats.some(c => c.name === PIVOT_STATE.row) ? PIVOT_STATE.row : cats[0].name;
  const opt = (v, l, sel) => `<option value="${esc(v)}"${sel === v ? " selected" : ""}>${esc(l)}</option>`;
  return `
      <div class="builder-field"><label>Rows</label>
        <select id="pvRow">${cats.map(c => opt(c.name, c.name, PIVOT_STATE.row)).join("")}</select></div>
      <div class="builder-field"><label>Columns (optional)</label>
        <select id="pvCol"><option value="">— none —</option>${cats.map(c => opt(c.name, c.name, PIVOT_STATE.col)).join("")}</select></div>
      <div class="builder-field"><label>Values</label>
        <select id="pvVal"><option value="__count__">Record count</option>${nums.map(c => opt(c, c, PIVOT_STATE.measure)).join("")}</select></div>
      <div class="builder-field"><label>Aggregation</label>
        <select id="pvAgg">${["sum", "avg", "min", "max", "median"].map(a => opt(a, { sum: "Sum", avg: "Average", min: "Minimum", max: "Maximum", median: "Median" }[a], PIVOT_STATE.agg)).join("")}</select></div>
      <div class="builder-field"><label>Filter (optional)</label>
        <select id="pvFCol"><option value="">— none —</option>${cats.map(c => opt(c.name, c.name, PIVOT_STATE.filter_col)).join("")}</select></div>
      <div class="builder-field"><label>Filter value</label>
        <select id="pvFVal"><option value="">— all —</option></select></div>
      <button class="btn btn-primary" id="pvBuild">${btnLabel}</button>`;
}

function tabPivotBuilder(p) {
  const cats = p.categorical_columns || [];
  if (!cats.length) return '<div class="empty">Pivot not possible — no categorical column detected in this data.</div>';
  return `
  <p class="muted small">PivotTable jaisa builder — Rows, Columns, Values, Aggregation aur Filter sab aap ke data ke asli columns se. Excel ke PivotTable options ke mutabiq. Isi pivot ka <b>chart</b> chahiye to result me "chart dekhein" button ya <b>Pivot Charts</b> tab use karein.</p>
  <div class="builder">
    <div class="builder-grid">${pivotGridHtml(p, "📊 Build Pivot")}</div>
  </div>
  <div id="pvResult">${PIVOT_LAST ? renderPivotResult(PIVOT_LAST) : '<div class="empty">Fields chunein aur "Build Pivot" dabaein.</div>'}</div>`;
}

let PIVOT_LAST = null;

function wirePivotBuilder() {
  wirePivotCore(() => { $("pvResult").innerHTML = renderPivotResult(PIVOT_LAST); wirePvToChart(); }, "pvResult");
  wirePvToChart();
}

function wirePvToChart() {
  const b = $("pvToChart");
  if (b) b.addEventListener("click", () => activateTab("pivotcharts"));
}

/* shared wiring — Pivot Builder aur Pivot Charts tabs same field ids use karte hain (ek waqt me sirf ek tab render hota hai) */
function wirePivotCore(renderResult, errId) {
  const cats = PAYLOAD.categorical_columns || [];
  const btnLabel = $("pvBuild").textContent;
  const fillVals = () => {
    const fc = $("pvFCol").value;
    const fv = $("pvFVal");
    const col = cats.find(c => c.name === fc);
    fv.innerHTML = '<option value="">— all —</option>' + (col ? col.values.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join("") : "");
  };
  fillVals();
  $("pvFCol").addEventListener("change", fillVals);
  $("pvVal").addEventListener("change", () => { $("pvAgg").disabled = $("pvVal").value === "__count__"; });
  $("pvAgg").disabled = $("pvVal").value === "__count__";
  $("pvBuild").addEventListener("click", async () => {
    const btn = $("pvBuild");
    btn.disabled = true; btn.textContent = "Building…";
    try {
      const res = await fetch("/api/pivot", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          job_id: PAYLOAD.job_id, sheet: PAYLOAD.sheet_name,
          row: $("pvRow").value, col: $("pvCol").value,
          measure: $("pvVal").value === "__count__" ? null : $("pvVal").value,
          agg: $("pvAgg").value,
          filter_col: $("pvFCol").value || null,
          filter_val: ($("pvFCol").value && $("pvFVal").value) ? $("pvFVal").value : null,
        })
      });
      const data = await res.json();
      if (!data.ok) throw new Error(data.error || "Pivot failed.");
      PIVOT_LAST = data.pivot;
      PIVOT_STATE = { row: $("pvRow").value, col: $("pvCol").value, measure: $("pvVal").value, agg: $("pvAgg").value, filter_col: $("pvFCol").value, filter_val: $("pvFVal").value };
      renderResult();
    } catch (err) {
      const tgt = $(errId);
      if (tgt) tgt.innerHTML = `<div class="alert alert-error">⚠ ${esc(err.message)}</div>`;
    } finally {
      btn.disabled = false; btn.textContent = btnLabel;
    }
  });
}

function renderPivotResult(pv, opts) {
  const mLabel = pv.measure ? `${pv.agg === "count" ? "Count" : pv.agg} of ${esc(pv.measure)}` : "Record count";
  const fNote = pv.filter ? ` · filter: ${esc(pv.filter.column)} = ${esc(pv.filter.value)}` : "";
  let html = `<div class="section-h">📊 ${esc(pv.row_dim)}${pv.col_dim ? " × " + esc(pv.col_dim) : ""} — ${mLabel}${fNote} <span class="muted small">(${pv.records.toLocaleString()} records)</span></div>`;
  if (pv.mode === "single") {
    const hasPct = pv.rows.some(r => r.pct != null);
    const headers = [{ label: pv.row_dim }].concat([{ label: mLabel, num: true }])
      .concat(hasPct ? [{ label: "% of total", num: true }] : []);
    const rows = pv.rows.map(r => [esc(r.label), fmtNum(r.value)].concat(hasPct ? [fmtPct(r.pct)] : []));
    html += tableHTML(headers, rows);
    if (pv.grand_total != null) html += `<p class="muted small">Grand total: <b>${fmtNum(pv.grand_total)}</b></p>`;
  } else {
    const headers = [{ label: pv.row_dim + " \\ " + pv.col_dim }].concat(pv.col_labels.map(l => ({ label: l, num: true }))).concat([{ label: "Total", num: true }]);
    const rows = pv.row_labels.map((rl, i) => {
      const cells = pv.matrix[i].map(v => v == null ? "—" : fmtNum(v));
      const tot = pv.row_totals[i];
      return [esc(rl)].concat(cells).concat([tot == null ? "—" : fmtNum(tot)]);
    });
    const trow = ["<b>Column total</b>"].concat(pv.col_totals.map(v => v == null ? "—" : `<b>${fmtNum(v)}</b>`))
      .concat(pv.grand_total != null ? [`<b>${fmtNum(pv.grand_total)}</b>`] : ["—"]);
    html += tableHTML(headers, rows.concat([trow]));
  }
  if (!opts || opts.withChartBtn !== false) html += `<div style="margin-top:12px"><button class="btn" id="pvToChart">📈 Isi pivot ka chart dekhein — Pivot Charts tab</button></div>`;
  html += `<p class="muted small">Note: ye calculation hai — values uploaded file se compute hui hain, kuch invent nahi hua. Top 30 rows / 15 columns tak display.</p>`;
  return html;
}

/* ================= PIVOT CHARTS TAB (PivotChart jaisa) ================= */
let PIVOT_CHART_TYPE = "bar";

function pcMeasureLabel(pv) {
  return pv.measure ? `${pv.agg === "count" ? "Count" : pv.agg} of ${pv.measure}` : "Record count";
}
function pcAxisFmt(v) {
  const r = Math.abs(v) >= 100 ? Math.round(v) : Math.round(v * 100) / 100;
  return fmtNum(r);
}
function pcAvailableTypes(pv) {
  return pv.mode === "cross"
    ? [["grouped", "📊 Grouped Bars"], ["stacked", "📚 Stacked"], ["stacked100", "🧱 100% Stacked"], ["heatmap", "🔥 Heatmap"]]
    : [["bar", "📈 Bar"], ["column", "📊 Column"], ["donut", "🍩 Donut"], ["line", "📉 Line"]];
}
function pcTypesHtml(pv) {
  const avail = pcAvailableTypes(pv);
  if (!avail.some(t => t[0] === PIVOT_CHART_TYPE)) PIVOT_CHART_TYPE = avail[0][0];
  return `<div class="fmt-bar" style="margin:14px 0">
    <span class="muted">Chart type:</span>
    ${avail.map(t => `<button class="fmt-btn${t[0] === PIVOT_CHART_TYPE ? " active" : ""}" data-ctype="${t[0]}">${t[1]}</button>`).join("")}
  </div>`;
}

function tabPivotCharts(p) {
  const cats = p.categorical_columns || [];
  if (!cats.length) return '<div class="empty">Pivot charts not possible — no categorical column detected in this data.</div>';
  return `
  <p class="muted small">PivotChart — Excel ke PivotChart jaisa: jo pivot aap banate hain wo chart ki shakal me dekhein. Fields chunein, "Build Chart" dabaein, phir upar se chart type badlein. Har chart ke neeche uska exact data table bhi hai — values uploaded file se compute hoti hain, kuch invent nahi hota.</p>
  <div class="builder">
    <div class="builder-grid">${pivotGridHtml(p, "📈 Build Chart")}</div>
  </div>
  ${PIVOT_LAST ? pcTypesHtml(PIVOT_LAST) : ""}
  <div id="pcCharts">${PIVOT_LAST ? pcChartsHtml(PIVOT_LAST) : '<div class="empty">Fields chunein aur "Build Chart" dabaein — chart yahan banega. Pivot Builder me pivot bana hua hai to wo yahan turant chart ban jayega.</div>'}</div>`;
}

function wirePivotCharts() {
  wirePivotCore(() => {
    const t = $("pcTypes");
    if (t) t.innerHTML = pcTypesHtml(PIVOT_LAST);
    wirePcTypes();
    $("pcCharts").innerHTML = pcChartsHtml(PIVOT_LAST);
  }, "pcCharts");
  wirePcTypes();
}

function wirePcTypes() {
  document.querySelectorAll("[data-ctype]").forEach(b => b.addEventListener("click", () => {
    PIVOT_CHART_TYPE = b.dataset.ctype;
    document.querySelectorAll("[data-ctype]").forEach(x => x.classList.toggle("active", x === b));
    $("pcCharts").innerHTML = pcChartsHtml(PIVOT_LAST);
  }));
}

function pcChartsHtml(pv) {
  const mLabel = pcMeasureLabel(pv);
  const fNote = pv.filter ? ` · filter: ${esc(pv.filter.column)} = ${esc(pv.filter.value)}` : "";
  const nCats = pv.mode === "cross" ? pv.row_labels.length : pv.rows.length;
  let cards = `<div class="kpi"><div class="kpi-label">Records</div><div class="kpi-value">${pv.records.toLocaleString()}</div><div class="kpi-sub">${nCats} categories</div></div>`;
  if (pv.grand_total != null) cards += `<div class="kpi"><div class="kpi-label">Grand total</div><div class="kpi-value">${fmtNum(pv.grand_total)}</div><div class="kpi-sub">${esc(mLabel)}</div></div>`;
  const topLabel = pv.mode === "single" ? (pv.rows[0] && pv.rows[0].label) : pv.row_labels[0];
  const topValue = pv.mode === "single" ? (pv.rows[0] ? pv.rows[0].value : null) : (pv.row_totals ? pv.row_totals[0] : null);
  if (topLabel != null) cards += `<div class="kpi good"><div class="kpi-label">Top: ${esc(String(topLabel))}</div><div class="kpi-value">${topValue == null ? "—" : fmtNum(topValue)}</div><div class="kpi-sub">sab se bara</div></div>`;
  let chart;
  switch (PIVOT_CHART_TYPE) {
    case "column": chart = pcColumnChart(pv); break;
    case "donut": chart = pcDonutChart(pv); break;
    case "line": chart = pcLineChart(pv); break;
    case "grouped": chart = pcGroupedChart(pv); break;
    case "stacked": chart = pcStackedChart(pv, false); break;
    case "stacked100": chart = pcStackedChart(pv, true); break;
    case "heatmap": chart = pcHeatmap(pv); break;
    default: chart = pcBarChart(pv);
  }
  return `<div class="section-h">📈 ${esc(pv.row_dim)}${pv.col_dim ? " × " + esc(pv.col_dim) : ""} — ${esc(mLabel)}${fNote} <span class="muted small">(${pv.records.toLocaleString()} records)</span></div>
    <div class="kpi-grid" style="margin-bottom:14px">${cards}</div>
    <div class="charts-grid">${chart}</div>
    <div class="section-h">📋 Chart ke peeche ka data (exact values)</div>
    ${renderPivotResult(pv, { withChartBtn: false })}`;
}

/* ---- 1. horizontal bar (CSS) ---- */
function pcBarChart(pv) {
  const rows = pv.rows.slice(0, 15);
  if (!rows.length) return '<div class="chart-card"><div class="empty">Koi category nahi mili.</div></div>';
  const max = Math.max(...rows.map(r => Number(r.value) || 0), 1e-9);
  const bars = rows.map((r, i) => {
    const v = Number(r.value) || 0;
    return `<div class="bar-row">
      <div class="bar-label" title="${esc(r.label)}">${esc(r.label)}</div>
      <div class="bar-track"><div class="bar-fill${i === 0 ? " gold" : ""}" style="width:${Math.max(v / max * 100, 2)}%"></div>
      <span class="bar-val">${fmtNum(r.value)}${r.pct != null ? " · " + fmtPct(r.pct) : ""}</span></div>
    </div>`;
  }).join("");
  return `<div class="chart-card" style="grid-column:1/-1">
    <h4>${esc(pv.row_dim)} — ${esc(pcMeasureLabel(pv))}</h4>
    <div class="chart-sub">Top ${rows.length} of ${pv.rows.length} categories · gold = sab se bara · number format (K/M/B) upar switch se badla ja sakta hai</div>
    ${bars}</div>`;
}

/* ---- 2. vertical column (SVG) ---- */
function pcColumnChart(pv) {
  const rows = pv.rows.slice(0, 12);
  if (!rows.length) return '<div class="chart-card"><div class="empty">Koi category nahi mili.</div></div>';
  const w = 740, h = 330, padL = 60, padR = 14, padT = 30, padB = 88;
  const vals = rows.map(r => Number(r.value) || 0);
  const max = Math.max(...vals, 0), min = Math.min(...vals, 0);
  const span = (max - min) || 1;
  const cw = (w - padL - padR) / rows.length;
  const y = v => padT + (1 - (v - min) / span) * (h - padT - padB);
  const base = y(Math.max(min, 0));
  const grid = [0, 0.25, 0.5, 0.75, 1].map(f => {
    const gv = min + f * span;
    return `<line x1="${padL}" y1="${y(gv).toFixed(1)}" x2="${w - padR}" y2="${y(gv).toFixed(1)}" stroke="#E3EAE6" stroke-width="1"/>
      <text x="${padL - 6}" y="${(y(gv) + 3).toFixed(1)}" text-anchor="end" font-size="9" fill="#5E6E66">${pcAxisFmt(gv)}</text>`;
  }).join("");
  const cols = rows.map((r, i) => {
    const v = vals[i];
    const x0 = padL + i * cw + cw * 0.16;
    const bw = cw * 0.68;
    const yy = v >= 0 ? y(v) : base;
    const hh = Math.max(Math.abs(base - y(v)), 1);
    const cx = x0 + bw / 2;
    const lab = `<text x="${cx.toFixed(1)}" y="${(h - padB + 16).toFixed(1)}" text-anchor="end" font-size="9" fill="#33413B" transform="rotate(-34 ${cx.toFixed(1)} ${h - padB + 16})">${esc(r.label)}</text>`;
    const val = cw >= 40 ? `<text x="${cx.toFixed(1)}" y="${(v >= 0 ? y(v) - 5 : y(v) + 12).toFixed(1)}" text-anchor="middle" font-size="8.5" font-weight="700" fill="#0B3D2E">${fmtNum(v)}</text>` : "";
    return `<rect x="${x0.toFixed(1)}" y="${yy.toFixed(1)}" width="${bw.toFixed(1)}" height="${hh.toFixed(1)}" fill="${i === 0 ? "#C9A227" : "#0F4C3A"}" rx="3"><title>${esc(r.label)}: ${fmtNum(r.value)}</title></rect>${val}${lab}`;
  }).join("");
  return `<div class="chart-card" style="grid-column:1/-1">
    <h4>${esc(pv.row_dim)} — ${esc(pcMeasureLabel(pv))}</h4>
    <div class="chart-sub">Top ${rows.length} of ${pv.rows.length} categories · gold = sab se bara · hover kar ke exact value dekhein</div>
    <svg class="chart-svg" viewBox="0 0 ${w} ${h}">${grid}${cols}</svg></div>`;
}

/* ---- 3. donut (SVG) ---- */
function pcDonutChart(pv) {
  const totalAll = pv.rows.reduce((a, r) => a + (Number(r.value) || 0), 0);
  if (!(totalAll > 0)) return `<div class="chart-card"><div class="empty">Donut is pivot par possible nahi — total ${fmtNum(totalAll)} hai (negative/zero totals donut me share nahi dikha sakte). Bar ya Column try karein.</div></div>`;
  const top = pv.rows.slice(0, 9);
  const rest = pv.rows.slice(9);
  let segs = top.map(r => ({ label: r.label, value: Number(r.value) || 0 }));
  if (rest.length) segs.push({ label: `Others (${rest.length} categories)`, value: rest.reduce((a, r) => a + (Number(r.value) || 0), 0) });
  const tot = totalAll || 1;
  let angle = -90;
  const paths = segs.map((s, i) => {
    const frac = s.value / tot;
    let a0 = angle, a1 = angle + frac * 360;
    if (a1 - a0 >= 359.99) a1 = a0 + 359.99;
    angle = a1;
    return `<path d="${arcPath(80, 80, 55, 60, a0, a1)}" fill="${PALETTE[i % PALETTE.length]}" stroke="#fff" stroke-width="1.5"><title>${esc(s.label)}: ${fmtNum(s.value)} (${(frac * 100).toFixed(1)}%)</title></path>`;
  }).join("");
  const legend = segs.map((s, i) => `<div class="legend-item"><span class="legend-dot" style="background:${PALETTE[i % PALETTE.length]}"></span><span style="max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(s.label)}">${esc(s.label)}</span> — <b>${fmtNum(s.value)}</b> (${(s.value / tot * 100).toFixed(1)}%)</div>`).join("");
  return `<div class="chart-card" style="grid-column:1/-1">
    <h4>${esc(pv.row_dim)} — total ka share</h4>
    <div class="chart-sub">${pv.rows.length} categories me se top ${top.length}${rest.length ? ` · baqi ${rest.length} "Others" me grouped` : ""} · total ${fmtNum(tot)}</div>
    <div class="donut-wrap">
      <svg class="chart-svg" viewBox="0 0 160 160" style="max-width:180px">${paths}
        <circle cx="80" cy="80" r="33" fill="#fff"/>
        <text x="80" y="76" text-anchor="middle" font-size="12" font-weight="700" fill="#0B3D2E">${fmtNum(tot)}</text>
        <text x="80" y="90" text-anchor="middle" font-size="9" fill="#5E6E66">total</text>
      </svg>
      <div class="legend">${legend}</div>
    </div></div>`;
}

/* ---- 4. line / area (SVG) ---- */
function pcLineChart(pv) {
  const rows = pv.rows.slice(0, 30);
  if (rows.length < 2) return '<div class="chart-card"><div class="empty">Line chart ke liye kam se kam 2 categories chahiye.</div></div>';
  const labels = rows.map(r => r.label), vals = rows.map(r => Number(r.value) || 0);
  const w = 740, h = 300, padL = 60, padR = 14, padT = 18, padB = 70;
  const min = Math.min(...vals), max = Math.max(...vals);
  const span = (max - min) || 1;
  const x = i => padL + (i * (w - padL - padR)) / Math.max(labels.length - 1, 1);
  const y = v => padT + (1 - (v - min) / span) * (h - padT - padB);
  const pts = vals.map((v, i) => `${x(i).toFixed(1)},${y(v).toFixed(1)}`).join(" ");
  const area = `${padL},${y(min).toFixed(1)} ${pts} ${x(vals.length - 1).toFixed(1)},${y(min).toFixed(1)}`;
  const grid = [0, 0.25, 0.5, 0.75, 1].map(f => {
    const gv = min + f * span;
    return `<line x1="${padL}" y1="${y(gv).toFixed(1)}" x2="${w - padR}" y2="${y(gv).toFixed(1)}" stroke="#E3EAE6" stroke-width="1"/>
      <text x="${padL - 6}" y="${(y(gv) + 3).toFixed(1)}" text-anchor="end" font-size="9" fill="#5E6E66">${pcAxisFmt(gv)}</text>`;
  }).join("");
  const labStep = Math.ceil(labels.length / 12);
  const xLabels = labels.map((l, i) => i % labStep === 0
    ? `<text x="${x(i).toFixed(1)}" y="${h - padB + 16}" text-anchor="end" font-size="8.5" fill="#5E6E66" transform="rotate(-34 ${x(i).toFixed(1)} ${h - padB + 16})">${esc(l)}</text>` : "").join("");
  const dots = vals.map((v, i) => `<circle cx="${x(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="3" fill="#0F4C3A"><title>${esc(labels[i])}: ${fmtNum(v)}</title></circle>`).join("");
  return `<div class="chart-card" style="grid-column:1/-1">
    <h4>${esc(pv.row_dim)} — ${esc(pcMeasureLabel(pv))}</h4>
    <div class="chart-sub">${labels.length} categories, value ke descending order me · hover kar ke exact value dekhein</div>
    <svg class="chart-svg" viewBox="0 0 ${w} ${h}">${grid}
      <polygon points="${area}" fill="rgba(15,76,58,.08)"/>
      <polyline points="${pts}" fill="none" stroke="#0F4C3A" stroke-width="2.2" stroke-linejoin="round"/>
      ${dots}${xLabels}
    </svg></div>`;
}

/* ---- 5. grouped bars — cross tab (SVG) ---- */
function pcGroupedChart(pv) {
  const rowLabels = pv.row_labels.slice(0, 8), colLabels = pv.col_labels.slice(0, 6);
  if (!rowLabels.length || !colLabels.length) return '<div class="chart-card"><div class="empty">Koi categories nahi mili.</div></div>';
  const w = 760, h = 340, padL = 60, padR = 14, padT = 26, padB = 96;
  const vals = [];
  rowLabels.forEach((_, i) => (pv.matrix[i] || []).forEach((v, j) => { if (j < colLabels.length && v != null) vals.push(Number(v) || 0); }));
  const max = Math.max(...vals, 0), min = Math.min(...vals, 0);
  const span = (max - min) || 1;
  const gw = (w - padL - padR) / rowLabels.length;
  const bw = Math.min(gw / colLabels.length * 0.82, 46);
  const y = v => padT + (1 - (v - min) / span) * (h - padT - padB);
  const base = y(Math.max(min, 0));
  const grid = [0, 0.25, 0.5, 0.75, 1].map(f => {
    const gv = min + f * span;
    return `<line x1="${padL}" y1="${y(gv).toFixed(1)}" x2="${w - padR}" y2="${y(gv).toFixed(1)}" stroke="#E3EAE6" stroke-width="1"/>
      <text x="${padL - 6}" y="${(y(gv) + 3).toFixed(1)}" text-anchor="end" font-size="9" fill="#5E6E66">${pcAxisFmt(gv)}</text>`;
  }).join("");
  const groups = rowLabels.map((rl, i) => {
    const mrow = pv.matrix[i] || [];
    const bars = colLabels.map((cl, j) => {
      const v = mrow[j];
      if (v == null) return "";
      const vv = Number(v) || 0;
      const x0 = padL + i * gw + (gw - bw * colLabels.length) / 2 + j * bw;
      const yy = vv >= 0 ? y(vv) : base;
      const hh = Math.max(Math.abs(base - y(vv)), 1);
      return `<rect x="${x0.toFixed(1)}" y="${yy.toFixed(1)}" width="${Math.max(bw - 2, 1).toFixed(1)}" height="${hh.toFixed(1)}" fill="${PALETTE[j % PALETTE.length]}" rx="2"><title>${esc(rl)} · ${esc(cl)}: ${fmtNum(v)}</title></rect>`;
    }).join("");
    const cx = padL + i * gw + gw / 2;
    return bars + `<text x="${cx.toFixed(1)}" y="${h - padB + 16}" text-anchor="end" font-size="9" fill="#33413B" transform="rotate(-34 ${cx.toFixed(1)} ${h - padB + 16})">${esc(rl)}</text>`;
  }).join("");
  const legend = colLabels.map((cl, j) => `<span style="display:inline-flex;align-items:center;gap:6px;margin-right:14px;font-size:.8rem"><span class="legend-dot" style="background:${PALETTE[j % PALETTE.length]}"></span>${esc(cl)}</span>`).join("");
  return `<div class="chart-card" style="grid-column:1/-1">
    <h4>${esc(pv.row_dim)} × ${esc(pv.col_dim)} — ${esc(pcMeasureLabel(pv))}</h4>
    <div class="chart-sub">Top ${rowLabels.length} ${esc(pv.row_dim)} × ${colLabels.length} ${esc(pv.col_dim)} (value ke hisab se sorted) · hover kar ke exact value dekhein</div>
    <div style="margin-bottom:8px">${legend}</div>
    <svg class="chart-svg" viewBox="0 0 ${w} ${h}">${grid}${groups}</svg></div>`;
}

/* ---- 6. stacked bars — cross tab (SVG), absolute ya 100% ---- */
function pcStackedChart(pv, hundred) {
  const rowLabels = pv.row_labels.slice(0, 12), colLabels = pv.col_labels.slice(0, 8);
  if (!rowLabels.length || !colLabels.length) return '<div class="chart-card"><div class="empty">Koi categories nahi mili.</div></div>';
  const w = 760, h = 340, padL = 60, padR = 14, padT = 26, padB = 96;
  const gw = (w - padL - padR) / rowLabels.length;
  const bw = Math.min(gw * 0.6, 64);
  const rowSum = rowLabels.map((_, i) => colLabels.reduce((a, _, j) => {
    const v = (pv.matrix[i] || [])[j];
    return a + (v != null && Number(v) > 0 ? Number(v) : 0);
  }, 0));
  const max = hundred ? 1 : Math.max(...rowSum, 1e-9);
  const y = v => padT + (1 - v / max) * (h - padT - padB);
  const grid = hundred
    ? [0, 0.25, 0.5, 0.75, 1].map(f => `<line x1="${padL}" y1="${y(f).toFixed(1)}" x2="${w - padR}" y2="${y(f).toFixed(1)}" stroke="#E3EAE6" stroke-width="1"/>
      <text x="${padL - 6}" y="${(y(f) + 3).toFixed(1)}" text-anchor="end" font-size="9" fill="#5E6E66">${Math.round(f * 100)}%</text>`).join("")
    : [0, 0.25, 0.5, 0.75, 1].map(f => `<line x1="${padL}" y1="${y(max * f).toFixed(1)}" x2="${w - padR}" y2="${y(max * f).toFixed(1)}" stroke="#E3EAE6" stroke-width="1"/>
      <text x="${padL - 6}" y="${(y(max * f) + 3).toFixed(1)}" text-anchor="end" font-size="9" fill="#5E6E66">${pcAxisFmt(max * f)}</text>`).join("");
  const cols = rowLabels.map((rl, i) => {
    let acc = 0;
    const mrow = pv.matrix[i] || [];
    const segs = colLabels.map((cl, j) => {
      const v0 = mrow[j];
      if (v0 == null || Number(v0) <= 0) return "";
      const v = Number(v0);
      const frac = hundred ? v / (rowSum[i] || 1) : v;
      const y0 = y(acc), y1 = y(acc + frac);
      acc += frac;
      const x0 = padL + i * gw + (gw - bw) / 2;
      return `<rect x="${x0.toFixed(1)}" y="${y1.toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.max(y0 - y1, 1).toFixed(1)}" fill="${PALETTE[j % PALETTE.length]}" stroke="#fff" stroke-width="0.6"><title>${esc(rl)} · ${esc(cl)}: ${fmtNum(v0)}${hundred ? ` (${(frac * 100).toFixed(1)}%)` : ""}</title></rect>`;
    }).join("");
    const cx = padL + i * gw + gw / 2;
    const totLab = hundred ? "100%" : (rowSum[i] > 0 ? fmtNum(rowSum[i]) : "—");
    return segs +
      `<text x="${cx.toFixed(1)}" y="${h - padB + 16}" text-anchor="end" font-size="9" fill="#33413B" transform="rotate(-34 ${cx.toFixed(1)} ${h - padB + 16})">${esc(rl)}</text>` +
      (hundred || rowSum[i] <= 0 ? "" : `<text x="${cx.toFixed(1)}" y="${(y(rowSum[i]) - 5).toFixed(1)}" text-anchor="middle" font-size="8.5" font-weight="700" fill="#0B3D2E">${totLab}</text>`);
  }).join("");
  const legend = colLabels.map((cl, j) => `<span style="display:inline-flex;align-items:center;gap:6px;margin-right:14px;font-size:.8rem"><span class="legend-dot" style="background:${PALETTE[j % PALETTE.length]}"></span>${esc(cl)}</span>`).join("");
  return `<div class="chart-card" style="grid-column:1/-1">
    <h4>${esc(pv.row_dim)} × ${esc(pv.col_dim)} — ${hundred ? "100% stacked" : "stacked"} ${esc(pcMeasureLabel(pv))}</h4>
    <div class="chart-sub">${hundred ? "har category ka apne total me se hissa (percentage)" : `stacked total = top ${colLabels.length} columns ka sum`} · Top ${rowLabels.length} ${esc(pv.row_dim)} · negative values (agar hain) chart me shamil nahi — table me dekh sakte hain</div>
    <div style="margin-bottom:8px">${legend}</div>
    <svg class="chart-svg" viewBox="0 0 ${w} ${h}">${grid}${cols}</svg></div>`;
}

/* ---- 7. heatmap — cross tab (HTML table) ---- */
function pcHeatmap(pv) {
  const rowLabels = pv.row_labels, colLabels = pv.col_labels;
  if (!rowLabels.length || !colLabels.length) return '<div class="chart-card"><div class="empty">Koi categories nahi mili.</div></div>';
  let max = 0;
  pv.matrix.forEach(rw => (rw || []).forEach(v => { if (v != null && Number(v) > max) max = Number(v); }));
  const cell = v => {
    if (v == null) return '<td class="num" style="background:#FAFBFA;color:#9AA8A1">—</td>';
    const t = max > 0 ? Number(v) / max : 0;
    const alpha = 0.07 + 0.88 * t;
    const dark = t > 0.55;
    return `<td class="num" style="background:rgba(15,76,58,${alpha.toFixed(3)});${dark ? "color:#fff;font-weight:600" : ""}" title="${esc(String(v))}">${fmtNum(v)}</td>`;
  };
  const head = `<tr><th style="text-align:left">${esc(pv.row_dim)} \\ ${esc(pv.col_dim)}</th>${colLabels.map(c => `<th class="num" title="${esc(c)}">${esc(c)}</th>`).join("")}<th class="num">Total</th></tr>`;
  const bodyRows = rowLabels.map((rl, i) => `<tr><td style="font-weight:600">${esc(rl)}</td>${(pv.matrix[i] || []).map(cell).join("")}<td class="num" style="font-weight:700">${pv.row_totals[i] == null ? "—" : fmtNum(pv.row_totals[i])}</td></tr>`).join("");
  const totRow = `<tr><td style="font-weight:700">Column total</td>${pv.col_totals.map(v => `<td class="num" style="font-weight:700">${v == null ? "—" : fmtNum(v)}</td>`).join("")}<td class="num" style="font-weight:700">${pv.grand_total == null ? "—" : fmtNum(pv.grand_total)}</td></tr>`;
  return `<div class="chart-card" style="grid-column:1/-1">
    <h4>Heatmap — ${esc(pv.row_dim)} × ${esc(pv.col_dim)}</h4>
    <div class="chart-sub">gehra (dark) rang = bara number · ${rowLabels.length} × ${colLabels.length} cells · hover kar ke exact value dekhein · ${pv.records.toLocaleString()} records se</div>
    <div class="table-wrap" style="max-height:520px;overflow:auto"><table class="data"><thead>${head}</thead><tbody>${bodyRows}${totRow}</tbody></table></div></div>`;
}

/* ================= DATA MODEL ================= */
let MODEL_RELS = null;

function tabModel(p) {
  let html = "";
  if (p.join_report) {
    const j = p.join_report;
    html += `<div class="join-banner">
      <b>🔗 Merged dataset active</b> — "${esc(j.left_sheet)}" [${esc(j.left_key)}] + "${esc(j.right_sheet)}" [${esc(j.right_key)}] · ${esc(j.join_type).toUpperCase()} join ·
      ${j.matched_left_rows.toLocaleString()}/${j.left_rows.toLocaleString()} left rows matched · ${j.unmatched_left_rows.toLocaleString()} unmatched (blank rakhi gayi) ·
      result: ${j.result_rows.toLocaleString()} rows × ${j.result_columns} columns${j.note ? `<br>⚠ ${esc(j.note)}` : ""}
    </div>`;
    html += `<p class="muted small">Excel report me "Merged_Data" sheet hai. Original workbook par wapas jane ke liye dobara file upload karein.</p>`;
  }
  const multi = (p.sheets_info || []).filter(s => s.rows > 0).length > 1;
  if (multi && !p.join_report) {
    html += `<div class="section-h">🔗 Sheet Relationships (Data Model)</div>
      <p class="muted small">Excel ke Data Model ki tarah — sheets ke darmiyan common keys detect karta hai, phir VALIDATED merge (join) karta hai. Merge se pehle matching, duplicates aur unmatched counts report hote hain.</p>
      <button class="btn btn-primary" id="dmDetect">🔍 Detect Relationships</button>
      <div id="dmResult"></div>`;
  } else if (!p.join_report) {
    html += `<div class="section-h">🔗 Data Model</div>
      <p class="muted small">Is workbook me ek hi data sheet hai, is liye sheet-to-sheet relationships possible nahi. Neeche data dictionary aur structure hai. Multi-sheet workbook upload karein to relationship detection + merge (join) khul jayega.</p>`;
  }
  // structure summary
  const idCol = (p.columns || []).find(c => c.role === "id");
  html += `<div class="section-h">🧩 Structure</div>`;
  const struct = [];
  if (idCol) {
    const uniqPct = idCol.unique / Math.max(p.kpis.total_records, 1) * 100;
    struct.push(["Primary key candidate", `"${esc(idCol.name)}" — ${idCol.unique.toLocaleString()} unique values (${uniqPct.toFixed(1)}% of rows)${uniqPct < 95 ? " — duplicates found, review required" : ""}`]);
  } else struct.push(["Primary key candidate", "No ID-like column detected"]);
  if (p.kpis.main_metric) struct.push(["Main measure", esc(p.kpis.main_metric)]);
  if (p.date_range) struct.push(["Reference date column", `${p.trend ? esc(p.trend.date_col) : "—"} (${p.date_range[0]} to ${p.date_range[1]})`]);
  if (p.kpis.utilization != null) struct.push(["Budget relationship", `Budget ↔ Expenditure columns linked (utilisation ${fmtPct(p.kpis.utilization)})`]);
  struct.push(["Categorical (dimension) fields", (p.categorical_columns || []).map(c => esc(c.name)).join(", ") || "—"]);
  struct.push(["Numeric (measure) fields", (p.numeric_columns || []).map(esc).join(", ") || "—"]);
  html += tableHTML([{ label: "Aspect" }, { label: "Detail" }], struct);
  html += `<div class="section-h">📖 Data Dictionary</div>` + tableHTML(
    [{ label: "Column" }, { label: "Role" }, { label: "Type" }, { label: "Unique", num: true }, { label: "Missing", num: true }, { label: "Sample / range" }],
    (p.columns || []).map(c => [esc(c.name), esc(c.role), esc(c.dtype), c.unique.toLocaleString(), c.missing.toLocaleString(),
      c.sample && c.sample.length ? c.sample.slice(0, 2).map(esc).join(", ") : (c.min != null ? `${fmtNum(c.min)} … ${fmtNum(c.max)}` : "—")]));
  return html;
}

function wireModel() {
  const btn = $("dmDetect");
  if (!btn) return;
  btn.addEventListener("click", async () => {
    btn.disabled = true; btn.textContent = "Detecting…";
    try {
      const res = await fetch("/api/model", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ job_id: PAYLOAD.job_id })
      });
      const data = await res.json();
      if (!data.ok) throw new Error(data.error || "Detection failed.");
      MODEL_RELS = data.relationships;
      $("dmResult").innerHTML = renderRels(MODEL_RELS);
      wireJoinButtons();
    } catch (err) {
      $("dmResult").innerHTML = `<div class="alert alert-error">⚠ ${esc(err.message)}</div>`;
    } finally {
      btn.disabled = false; btn.textContent = "🔍 Detect Relationships";
    }
  });
}

function renderRels(rels) {
  if (!rels || !rels.length) return `<div class="empty">Koi common key relationship detect nahi hui. Sheets me shared values wale columns nahi mile (ya bohat kam overlap hai).</div>`;
  const rows = rels.map((r, i) => [
    `${esc(r.left_sheet)}<br><span class="muted small">${r.left_rows.toLocaleString()} rows</span>`,
    `<b>${esc(r.left_key)}</b>`,
    `${esc(r.right_sheet)}<br><span class="muted small">${r.right_rows.toLocaleString()} rows</span>`,
    `<b>${esc(r.right_key)}</b>`,
    r.matched_left_pct != null ? fmtPct(r.matched_left_pct) : "—",
    r.duplicate_right_keys ? `<span class="badge warn">${r.duplicate_right_keys} dup</span>` : '<span class="badge info">unique</span>',
    `<div style="display:flex;gap:6px;align-items:center">
       <select class="join-how" data-i="${i}"><option value="left">Left join</option><option value="inner">Inner join</option></select>
       <button class="btn btn-primary join-btn" data-i="${i}" style="padding:6px 12px;font-size:.8rem">Merge & Analyze</button>
     </div>`
  ]);
  return `<p class="muted small">"Match %" = kitne left values right sheet me mile. "Unique" = right key duplicate-free hai ya nahi (one-to-many hoga to rows multiply hoti hain — report me note hoga).</p>` +
    tableHTML([{ label: "Fact sheet" }, { label: "Left key" }, { label: "Lookup sheet" }, { label: "Right key" }, { label: "Match %", num: true }, { label: "Right key" }, { label: "Action" }], rows);
}

function wireJoinButtons() {
  document.querySelectorAll(".join-btn").forEach(b => b.addEventListener("click", async () => {
    const i = +b.dataset.i;
    const rel = MODEL_RELS[i];
    const how = document.querySelector(`.join-how[data-i="${i}"]`).value;
    b.disabled = true; b.textContent = "Merging…";
    try {
      const res = await fetch("/api/join", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          job_id: PAYLOAD.job_id,
          left_sheet: rel.left_sheet, left_key: rel.left_key,
          right_sheet: rel.right_sheet, right_key: rel.right_key,
          how,
          standardize_case: $("optStandardize").checked,
          fill_missing: $("optFill").checked,
          remove_duplicates: $("optDedup").checked,
        })
      });
      const data = await res.json();
      if (!data.ok) throw new Error(data.error || "Merge failed.");
      PAYLOAD = data;
      PIVOT_LAST = null; PIVOT_STATE.row = null;
      renderResults();
      agentSay(`Merge mukammal: ${data.join_report.result_rows.toLocaleString()} rows × ${data.join_report.result_columns} columns. Ab poora analysis merged data par chal raha hai — Excel report me "Merged_Data" sheet bhi hai.`);
    } catch (err) {
      alert("Merge failed: " + err.message);
      b.disabled = false; b.textContent = "Merge & Analyze";
    }
  }));
}

/* ================= VALIDATION ================= */
function tabValidation(p) {
  const rules = p.validation_rules || [];
  if (!rules.length) return '<div class="empty">No validation rules could be generated for this data.</div>';
  const rows = rules.map(r => {
    let detail = esc(r.rule);
    if (r.type === "list" && r.values) detail += `<br><span class="muted small">${r.values.slice(0, 8).map(esc).join(" · ")}${r.values.length > 8 ? " …" : ""}</span>`;
    return [`<span class="badge info">${esc(r.type)}</span>`, esc(r.column), detail, esc(r.purpose)];
  });
  return `
  <p class="muted small">Ye rules aap ke data ke asli values se generate hui hain aur <b>Cleaned_Data sheet me actually apply</b> ho chuki hain (dropdowns + range checks, data ke neeche 300 future rows par bhi). Naya entry galat ho to Excel reject karega.</p>
  ${tableHTML([{ label: "Type" }, { label: "Column" }, { label: "Rule" }, { label: "Purpose" }], rows)}
  <p class="muted small">Excel me khud dekhne ke liye: Cleaned_Data sheet → column select → Data → Data Validation. Dropdown lists "Validation_Rules" sheet se refer hoti hain (unhe delete na karein).</p>`;
}

/* ================= POWER QUERY ================= */
function tabPowerQuery(p) {
  const pq = p.powerquery || {};
  const m = pq.m_code || "-- Not available for this dataset.";
  const steps = pq.ui_steps || [];
  return `
  <p class="muted small">Power Query = refreshable cleaning pipeline. Ye script aap ke file ke <b>asli columns aur asli issues</b> se generate hui hai — same steps jo agent ne kiye. Ek dafa load karein, phir har mahine sirf <b>Refresh</b> dabaein.</p>
  <div class="cmd-row" style="margin:0 0 10px">
    <button class="btn btn-secondary" id="pqCopy">📋 Copy M Code</button>
    <span class="muted small" id="pqCopyMsg"></span>
  </div>
  <div class="pq-code" id="pqCode">${esc(m)}</div>
  <div class="section-h">🪜 Excel Steps (manual route)</div>
  <ol class="steps-list">${steps.map(s => `<li>${esc(s)}</li>`).join("")}</ol>
  <div class="section-h">🧭 Advanced Editor me paste kaise karein</div>
  <ol class="steps-list">
    <li>Data → Get Data → Launch Power Query Editor</li>
    <li>Home → Advanced Editor</li>
    <li>Poora code replace kar ke upar wala M code paste karein</li>
    <li>Done — file path (C:\\Data\\...) apne computer ke mutabiq badlein</li>
  </ol>`;
}

function wirePowerQuery() {
  const btn = $("pqCopy");
  if (!btn) return;
  btn.addEventListener("click", async () => {
    const code = PAYLOAD.powerquery ? PAYLOAD.powerquery.m_code : "";
    let ok = false;
    try { await navigator.clipboard.writeText(code); ok = true; } catch (e) {
      try {
        const ta = document.createElement("textarea");
        ta.value = code; document.body.appendChild(ta); ta.select();
        ok = document.execCommand("copy"); document.body.removeChild(ta);
      } catch (e2) { ok = false; }
    }
    $("pqCopyMsg").textContent = ok ? "✓ Copy ho gaya — Power Query Advanced Editor me paste karein" : "Copy nahi hua — code select kar ke Ctrl+C karein";
  });
}

/* ---------------- number format bar ---------------- */
document.querySelectorAll(".fmt-btn").forEach(b => b.addEventListener("click", () => {
  document.querySelectorAll(".fmt-btn").forEach(x => x.classList.remove("active"));
  b.classList.add("active");
  NUMFMT = b.dataset.fmt;
  renderTab(CURRENT_TAB);
}));

/* ---------------- AI agent command box ---------------- */
$("cmdGo").addEventListener("click", runCommand);
$("cmdInput").addEventListener("keydown", e => { if (e.key === "Enter") runCommand(); });
document.querySelectorAll("#cmdChips .chip").forEach(c => c.addEventListener("click", () => {
  $("cmdInput").value = c.dataset.cmd; runCommand();
}));

function agentSay(html) {
  const box = $("agentReply");
  box.hidden = false;
  box.innerHTML = "🤖 <b>Agent:</b> " + html;
}

function activateTab(name) {
  const t = document.querySelector(`.tab[data-tab="${name}"]`);
  if (t) t.click();
  document.querySelector(".tab-card").scrollIntoView({ behavior: "smooth", block: "start" });
}

function runCommand() {
  if (!PAYLOAD) { agentSay("Pehle koi file upload kar ke analysis chalaein."); return; }
  const raw = ($("cmdInput").value || "").trim().toLowerCase();
  if (!raw) return;
  $("cmdInput").value = "";
  const p = PAYLOAD;

  // number formatting
  if (/(million|millon|milli|m\b|mn|em)/.test(raw) && /(me|mein|in|display|dikh|show|format)/.test(raw)) {
    document.querySelector('.fmt-btn[data-fmt="m"]').click();
    agentSay("Numbers ab <b>Millions (M)</b> me display ho rahe hain — sirf display format badla hai, asal values wahi hain (Excel me bhi aise hi alag alag rakhi gayi hain).");
    return;
  }
  if (/(billion|billion|arab|bn\b)/.test(raw)) {
    document.querySelector('.fmt-btn[data-fmt="b"]').click();
    agentSay("Numbers ab <b>Billions (B)</b> me display ho rahe hain — display formatting only.");
    return;
  }
  if (/(thousand|k format|hazaar|hazaron)/.test(raw)) {
    document.querySelector('.fmt-btn[data-fmt="k"]').click();
    agentSay("Numbers ab <b>Thousands (K)</b> me hain.");
    return;
  }
  if (/(full|complete|poora|pura|asli|original)/.test(raw) && /(number|format|me|mein)/.test(raw)) {
    document.querySelector('.fmt-btn[data-fmt="full"]').click();
    agentSay("Numbers ab <b>full figures</b> me hain.");
    return;
  }
  // top / bottom N
  const topM = raw.match(/\btop\s*(\d+)?/);
  const botM = raw.match(/\b(bottom|last|sab se kam|sabse kam|lowest)\s*(\d+)?/);
  if (topM || botM) {
    const isTop = !!topM;
    const n = parseInt((isTop ? topM[1] : botM[1]) || "10", 10);
    if (!p.top50) { agentSay("Ranking possible nahi — is data me koi numeric amount column detect nahi hui."); return; }
    const obj = isTop ? p.top50 : p.bottom50;
    const rows = obj.rows.slice(0, Math.min(n, 50));
    const headers = [{ label: "Rank", num: true }].concat(obj.columns.map(c => ({ label: c, num: c === p.kpis.main_metric })));
    const trs = rows.map(r => [r.rank].concat(obj.columns.map(c => typeof r[c] === "number" ? fmtNum(r[c]) : esc(r[c]))));
    const title = isTop ? `Top ${rows.length} by ${p.kpis.main_metric}` : `Bottom ${rows.length} by ${p.kpis.main_metric}`;
    const box = $("agentReply");
    box.hidden = false;
    box.innerHTML = `🤖 <b>Agent:</b> ${esc(title)} — ${rows.length} records:<br><br>` + tableHTML(headers, trs);
    box.scrollIntoView({ behavior: "smooth" });
    return;
  }
  // department wise
  const dimHit = p.pivots.find(pv => /(department|dept|mahkama|mahakma)/.test(raw) && pv.dimension.toLowerCase().match(/depart|dept|wing|branch|ministry|office|section|directorate/))
    || (/(department|dept|mahkama)/.test(raw) ? p.pivots.find(pv => pv.dimension.toLowerCase() === "department") : null);
  if (dimHit) { activateTab("summary"); agentSay(`Department-wise summary khol di (${esc(dimHit.dimension)}). Puri detail "Summary & Pivots" tab aur Excel report me hai.`); return; }
  const distHit = p.pivots.find(pv => /(district|zilla|tehsil|city|region|division|area|location)/.test(raw) && pv.dimension.toLowerCase().match(/district|tehsil|zilla|city|region|division|area|location|province/));
  if (distHit) { activateTab("summary"); agentSay(`District/region-wise summary khol di (${esc(distHit.dimension)}).`); return; }
  // duplicates
  if (/(duplicate|duplicat|do bar|dobara|repeat|mukarrar)/.test(raw)) {
    activateTab("exceptions");
    agentSay(p.dup_rows ? `${p.dup_rows} exact duplicate rows mili hain — Exceptions tab me list hai. Ye remove NAHI ki gayi (audit ke liye rakhi gayi hain).${p.id_dup ? ` Iske ilawa "${esc(p.id_dup.column)}" me bhi duplicate values hain.` : ""}` : "Koi exact duplicate record nahi mili.");
    return;
  }
  // missing
  if (/(missing|ghaib|blank|khaali|khali|na ho)/.test(raw)) {
    activateTab("quality");
    const k = p.kpis;
    const worst = (p.columns || []).filter(c => c.missing > 0).sort((a, b) => b.missing - a.missing).slice(0, 5);
    agentSay(`Total ${k.missing_cells.toLocaleString()} missing cells (${k.missing_pct}% of all cells).` +
      (worst.length ? ` Sab se zyada missing: ${worst.map(w => `"${esc(w.name)}" (${w.missing})`).join(", ")}.` : "") +
      " Detail Data Quality tab me hai. Values invent NAHI kiye — sirf blank report kiye hain.");
    return;
  }
  // outliers
  if (/(outlier|anomal|ajeeb|unusual|ghair|anokha|ajab)/.test(raw)) {
    activateTab("exceptions");
    agentSay(p.outlier_total ? `${p.outlier_total} statistical anomalies (IQR method) flag hui hain — "requires review". Ye kisi ghalti ka saboot nahi, sirf review ke liye hain.` : "Koi statistical anomaly detect nahi hui.");
    return;
  }
  // trend
  if (/(trend|monthly|mahana|mahine|month)/.test(raw)) {
    if (p.trend) { activateTab("dashboard"); agentSay(`Monthly trend chart dikhla di — "${esc(p.trend.date_col)}" ke hisab se${p.trend.measure ? `, ${esc(p.trend.measure)} ka` : ""}.`); }
    else agentSay("Monthly trend nahi ban sakti — is data me koi date column detect nahi hui (ya dates parse nahi ho sakin).");
    return;
  }
  // budget
  if (/(budget|expenditure|kharch|utili|muqabla|compare)/.test(raw)) {
    if (p.budget) { activateTab("summary"); agentSay(`Budget vs Expenditure: total budget ${fmtNum(p.budget.total_budget)}, expenditure ${fmtNum(p.budget.total_expenditure)}, utilization ${fmtPct(p.budget.utilization)}. Detail Summary tab me hai.`); }
    else agentSay("Budget analysis possible nahi — data me budget aur expenditure dono columns nahi mile (ya wo numeric nahi hain).");
    return;
  }
  // summary / report
  if (/(summary|report|khulasa|overview|jalwa|sar)/.test(raw)) {
    activateTab("summary");
    agentSay(`Summary khol di: ${p.kpis.total_records.toLocaleString()} records, ${p.kpis.duplicate_rows} duplicates, ${p.kpis.missing_cells.toLocaleString()} missing cells${p.kpis.utilization != null ? `, utilization ${fmtPct(p.kpis.utilization)}%` : ""}. Poora Excel report download banner se le sakte hain.`);
    return;
  }
  // utilization
  if (/(utili)/.test(raw)) {
    agentSay(p.budget ? `Utilization ${fmtPct(p.budget.utilization)} (expenditure ${fmtNum(p.budget.total_expenditure)} ÷ budget ${fmtNum(p.budget.total_budget)}).` : "Utilization calculate nahi ho sakti — budget/expenditure columns missing hain.");
    return;
  }
  // pivot charts
  if (/(pivot ?charts?|pivotcharts|charts? banao|charts? dikhao|charts? dikha|graphs? banao|graphs? dikhao|pivot graph|chart ban)/.test(raw)) {
    activateTab("pivotcharts");
    agentSay(PIVOT_LAST
      ? `Pivot Charts khol di — <b>${esc(PIVOT_LAST.row_dim)}${PIVOT_LAST.col_dim ? " × " + esc(PIVOT_LAST.col_dim) : ""}</b> ka chart. Chart type (Bar / Column / Donut / Line / Stacked / Heatmap) upar se badal sakte hain — values wahi hain jo aap ke data se aayi hain.`
      : "Pivot Charts tab khol diya — fields chunein aur \"Build Chart\" dabaein. (Agar Pivot Builder me already pivot bana hua hai to wo yahan turant chart ban jayega.)");
    return;
  }
  // pivot builder
  if (/(pivot|peevo|cross tab|crosstab|matrix)/.test(raw)) {
    activateTab("pivot");
    agentSay("Pivot Builder khol diya — Rows, Columns, Values, Aggregation (Sum/Average/Min/Max/Median) aur Filter sab options aap ke data ke asli columns se chunein.");
    return;
  }
  // data model
  if (/(data model|model|relationship|rishta|join|merge|milao|jor)/.test(raw)) {
    activateTab("model");
    agentSay("Data Model tab khol diya — multi-sheet workbook me 'Detect Relationships' se common keys milengi, phir validated merge (join) ho sakta hai. Single-sheet me data dictionary dikhta hai.");
    return;
  }
  // power query
  if (/(power query|powerquery|pq\b|m code|m code|refresh)/.test(raw)) {
    activateTab("pq");
    agentSay("Power Query script khol di — aap ke file ke asli columns se bani hui M code. Copy kar ke Power Query → Advanced Editor me paste karein, phir sirf Refresh All chahiye hota hai.");
    return;
  }
  // validation
  if (/(validation|valid|rule|qaida|dro?p ?down|check)/.test(raw)) {
    activateTab("valid");
    agentSay(`Validation rules khol di — ${PAYLOAD.validation_rules ? PAYLOAD.validation_rules.length : 0} rules aap ke data se generate hui hain aur Excel report ki Cleaned_Data sheet me actually apply ho chuki hain (dropdowns + range checks).`);
    return;
  }
  // formula recipes
  if (/(formula|formulas|sumifs|countifs|rank)/.test(raw) && !/(vlookup|xlookup)/.test(raw)) {
    activateTab("lookup");
    agentSay("Formula recipes khol di — VLOOKUP / XLOOKUP / SUMIFS / COUNTIFS / RANK, sab aap ke workbook ke asli ranges se. Excel report ki \"Formulas\" sheet me live demo formula bhi hai.");
    return;
  }
  // lookup
  if (/(vlookup|xlookup|lookup|look ?up|dhoondo|nikalo|laao|bring)/.test(raw)) {
    activateTab("lookup");
    agentSay("Lookup (VLOOKUP/XLOOKUP) tool khol diya — key column, lookup sheet aur return column chunein. Preview ke baad 'Apply & Re-analyze' se naya column poore analysis me shamil ho jayega. Excel formulas bhi milengi jo aap khud paste kar sakein.");
    return;
  }
  // dashboard
  if (/(dashboard|dash board|board|summary view|manzar)/.test(raw)) {
    activateTab("dashboard");
    agentSay("Dashboard khol diya — KPI cards, charts, management summary aur quick tables ek jagah.");
    return;
  }
  agentSay('Ye command samajh nahi aayi. Yehi try karein: <b>"top 10"</b>, <b>"department wise summary"</b>, <b>"pivot banao"</b>, <b>"pivot chart banao"</b>, <b>"vlookup laga do"</b>, <b>"dashboard"</b>, <b>"data model"</b>, <b>"power query"</b>, <b>"validation rules"</b>, <b>"duplicates dikhao"</b>, <b>"missing values"</b>, <b>"outliers"</b>, <b>"monthly trend"</b>, <b>"million me dikhao"</b>, <b>"budget vs expenditure"</b>, <b>"summary"</b>.');
}

/* ================= DASHBOARD TAB ================= */
function tabDashboard(p) {
  let summaryHtml = "";
  const mg = p.mgmt_summary || [];
  if (mg.length) {
    summaryHtml = `<div class="section-h">📝 Management Summary (auto)</div>
      <div class="join-banner" style="display:flex;flex-direction:column;gap:4px">
        ${mg.map(line => {
          const cls = line.startsWith("POTENTIAL ISSUE") ? 'style="color:var(--red);font-weight:600"'
                       : line.startsWith("DATA LIMITATION") ? 'style="color:var(--amber)"'
                       : line.startsWith("OBSERVATION") ? 'style="color:var(--green-800)"' : "";
          return `<div ${cls}>${esc(line)}</div>`;
        }).join("")}
      </div>`;
  }
  let quickTables = "";
  if (p.status_dist) {
    const total = p.status_dist.rows.reduce((a, r) => a + r.count, 0) || 1;
    quickTables += `<div class="section-h">📌 ${esc(p.status_dist.column)} distribution</div>` +
      tableHTML([{ label: p.status_dist.column }, { label: "Records", num: true }, { label: "Share %", num: true }],
        p.status_dist.rows.map(r => [esc(r.label), r.count.toLocaleString(), fmtPct(r.count / total * 100)]));
  }
  if (p.top50 && p.top50.rows.length) {
    const showCols = p.top50.columns.slice(0, 4);
    quickTables += `<div class="section-h">🏆 Top 5 by ${esc(p.kpis.main_metric)}</div>` +
      tableHTML([{ label: "#" }, ...showCols.map(c => ({ label: c, num: c === p.kpis.main_metric }))],
        p.top50.rows.slice(0, 5).map(r => [r.rank, ...showCols.map(c => typeof r[c] === "number" ? fmtNum(r[c]) : esc(r[c]))]));
  }
  if (p.budget && p.budget.per_dept && p.budget.per_dept.length) {
    quickTables += `<div class="section-h">💰 Budget position by department (top 6)</div>` +
      tableHTML([{ label: "Department" }, { label: "Budget", num: true }, { label: "Expenditure", num: true },
                 { label: "Remaining", num: true }, { label: "Utilization %", num: true }],
        p.budget.per_dept.slice(0, 6).map(d => [esc(d.label), fmtNum(d.budget), fmtNum(d.expenditure), fmtNum(d.remaining), fmtPct(d.utilization)]));
  }
  return `
  <div class="kpi-grid" id="kpiGrid"></div>
  ${summaryHtml}
  <div class="charts-grid" id="chartsGrid"></div>
  ${quickTables}
  <p class="muted small">Charts aur KPIs Cleaned_Data se compute hui hain. Number format upar "Full / K / M / B" switch se badal sakte hain (sirf display — asal values wahi hain). Excel report ke "Dashboard" sheet me ye sab charts + management summary + conditional formatting shamil hai.</p>`;
}

/* ================= LOOKUP TOOL (VLOOKUP / XLOOKUP) ================= */
let LOOKUP_LAST = null;

function tabLookup(p) {
  if (p.is_merged) {
    return `<div class="empty">Lookup original workbook par available hai — ye dataset already merged/lookup se bana hai. Original file dobara upload karein.</div>`;
  }
  const myCols = (p.columns || []).map(c => c.name);
  if (!myCols.length) return '<div class="empty">No columns available.</div>';
  const sheets = p.sheets_info || [];
  const opt = (v, l, sel) => `<option value="${esc(v)}"${sel === v ? " selected" : ""}>${esc(l)}</option>`;
  const sheetOpts = sheets.map(s =>
    opt(s.name, `${s.name} (${s.rows.toLocaleString()}×${s.cols})${s.name === p.sheet_name ? " — same sheet" : ""}`, LOOKUP_LAST ? LOOKUP_LAST.lookup_sheet : (sheets.find(x => x.name !== p.sheet_name) || sheets[0]).name)).join("");
  return `
  <p class="muted small">Excel ka VLOOKUP/XLOOKUP — <b>exact match, case-insensitive, first match wins</b> (bilkul VLOOKUP FALSE jaisa). Do kaam ek saath:
  (1) <b>Preview + Apply</b> — naya column aap ke data me add kar ke poora analysis dobara chalega;
  (2) <b>Ready-made Excel formulas</b> — aap ke workbook ke asli ranges se bani hui, jo aap khud Excel me paste kar sakein.</p>
  <div class="builder">
    <div class="builder-grid">
      <div class="builder-field"><label>Aap ki key column</label>
        <select id="lkMyKey">${myCols.map(c => opt(c, c, LOOKUP_LAST ? LOOKUP_LAST.from_key : (p.columns.find(x => x.role === "id") || p.columns[0]).name)).join("")}</select></div>
      <div class="builder-field"><label>Lookup sheet</label>
        <select id="lkSheet">${sheetOpts}</select></div>
      <div class="builder-field"><label>Lookup key column</label>
        <select id="lkKey"><option value="">— pehle sheet chunein —</option></select></div>
      <div class="builder-field"><label>Return column (jo laani hai)</label>
        <select id="lkRet"><option value="">— pehle sheet chunein —</option></select></div>
      <button class="btn btn-primary" id="lkRun">🔍 Run Lookup (Preview)</button>
    </div>
  </div>
  ${formulaRecipesHtml(p)}
  <div id="lkResult">${LOOKUP_LAST ? renderLookupResult(LOOKUP_LAST) : '<div class="empty">Columns chunein aur "Run Lookup" dabaein.</div>'}</div>`;
}

function formulaRecipesHtml(p) {
  const rec = p.formulas || [];
  if (!rec.length) return "";
  return `<div class="section-h">🧮 Ready-made Excel formulas (aap ke workbook ke asli ranges se — bina preview ke bhi yahan hain)</div>
    <div class="table-wrap"><table class="data"><thead><tr>
      <th>Formula type</th><th>Formula (copy-paste)</th><th>Kya karta hai</th></tr></thead><tbody>
    ${rec.map(f => `<tr>
      <td><b>${esc(f.section)}</b></td>
      <td><code style="font-size:.78rem">${esc(f.formula)}</code></td>
      <td>${esc(f.purpose)}<br><span class="muted small">${esc(f.explain)}</span></td>
    </tr>`).join("")}
    </tbody></table></div>
    <p class="muted small">💡 XLOOKUP ke liye Excel 365/2021+ chahiye — purane versions ke liye VLOOKUP ya INDEX/MATCH use karein. Poori detail Excel report ki "Formulas" sheet me bhi hai (live demo formula ke saath).</p>`;
}

function wireLookup() {
  const p = PAYLOAD;
  const sheets = p.sheets_info || [];
  const fillCols = () => {
    const sname = $("lkSheet").value;
    const s = sheets.find(x => x.name === sname);
    const cols = s ? s.columns : [];
    const keep = LOOKUP_LAST && LOOKUP_LAST.lookup_sheet === sname ? LOOKUP_LAST : null;
    $("lkKey").innerHTML = cols.map(c => `<option value="${esc(c)}"${keep && keep.lookup_key === c ? " selected" : ""}>${esc(c)}</option>`).join("") || '<option value="">—</option>';
    $("lkRet").innerHTML = cols.map(c => `<option value="${esc(c)}"${keep && keep.return_col === c ? " selected" : ""}>${esc(c)}</option>`).join("") || '<option value="">—</option>';
  };
  fillCols();
  $("lkSheet").addEventListener("change", fillCols);
  $("lkRun").addEventListener("click", async () => {
    const btn = $("lkRun");
    btn.disabled = true; btn.textContent = "Looking up…";
    try {
      const res = await fetch("/api/lookup", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          job_id: PAYLOAD.job_id, cur_sheet: PAYLOAD.sheet_name,
          from_key: $("lkMyKey").value, lookup_sheet: $("lkSheet").value,
          lookup_key: $("lkKey").value, return_col: $("lkRet").value,
          apply: false,
        })
      });
      const data = await res.json();
      if (!data.ok) throw new Error(data.error || "Lookup failed.");
      LOOKUP_LAST = { ...data, from_key: $("lkMyKey").value };
      $("lkResult").innerHTML = renderLookupResult(LOOKUP_LAST);
      wireApply();
    } catch (err) {
      $("lkResult").innerHTML = `<div class="alert alert-error">⚠ ${esc(err.message)}</div>`;
    } finally {
      btn.disabled = false; btn.textContent = "🔍 Run Lookup (Preview)";
    }
  });
  wireApply();
}

function wireApply() {
  const btn = $("lkApply");
  if (!btn) return;
  btn.addEventListener("click", async () => {
    btn.disabled = true; btn.textContent = "Applying & re-analyzing…";
    try {
      const res = await fetch("/api/lookup", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          job_id: PAYLOAD.job_id, cur_sheet: PAYLOAD.sheet_name,
          from_key: LOOKUP_LAST.from_key, lookup_sheet: LOOKUP_LAST.lookup_sheet,
          lookup_key: LOOKUP_LAST.lookup_key, return_col: LOOKUP_LAST.return_col,
          apply: true,
          standardize_case: $("optStandardize").checked,
          fill_missing: $("optFill").checked,
          remove_duplicates: $("optDedup").checked,
        })
      });
      const data = await res.json();
      if (!data.ok) throw new Error(data.error || "Apply failed.");
      PAYLOAD = data;
      PIVOT_LAST = null; PIVOT_STATE.row = null; LOOKUP_LAST = null;
      renderResults();
      agentSay(`Lookup apply ho gaya — naya column "${esc(data.lookup_report.new_column)}" ab poore analysis, pivots aur dashboard mein shamil hai (${data.lookup_report.matched}/${data.lookup_report.total_rows} matched). Excel report mein "Merged_Data" sheet ye combined data rakhti hai.`);
    } catch (err) {
      alert("Apply failed: " + err.message);
      btn.disabled = false; btn.textContent = "✅ Apply & Re-analyze";
    }
  });
}

function renderLookupResult(L) {
  const s = L.stats, f = L.formulas || {};
  const statsCards = `
    <div class="kpi-grid" style="margin-bottom:14px">
      <div class="kpi good"><div class="kpi-label">Matched</div><div class="kpi-value">${s.matched.toLocaleString()}</div><div class="kpi-sub">${s.match_pct}% of rows</div></div>
      <div class="kpi ${s.not_found ? "warn" : "good"}"><div class="kpi-label">Not Found</div><div class="kpi-value">${s.not_found.toLocaleString()}</div><div class="kpi-sub">blank rahenge — kuch invent nahi hua</div></div>
      <div class="kpi ${s.duplicate_lookup_keys ? "warn" : "good"}"><div class="kpi-label">Duplicate keys (lookup)</div><div class="kpi-value">${s.duplicate_lookup_keys.toLocaleString()}</div><div class="kpi-sub">${s.duplicate_lookup_keys ? "first match use hua (VLOOKUP behaviour)" : "lookup key unique hai"}</div></div>
      <div class="kpi"><div class="kpi-label">Lookup rows</div><div class="kpi-value">${s.lookup_rows.toLocaleString()}</div><div class="kpi-sub">in "${esc(L.lookup_sheet)}"</div></div>
    </div>`;
  const preview = tableHTML(
    [{ label: `${L.from_key} (aap ki key)` }, { label: `${L.formulas ? esc(L.formulas.return_col) : ""} → naya column` }],
    (L.preview || []).map(r => [esc(r.key) || "—", esc(r.value) || '<span class="muted">Not Found</span>']));
  const fx = f.xlookup ? `
  <div class="section-h">🧮 Ready-made Excel formulas (aap ke workbook ke asli ranges se)</div>
  <div class="muted-box" style="margin-bottom:8px">Naye column ke pehle cell (row 2) me daalein aur neeche drag karein. "$" absolute references is liye hain ke drag karne par ranges fix rahen.</div>
  <table class="data"><tbody>
    <tr><td style="width:130px"><b>XLOOKUP</b><br><span class="muted small">Excel 365 / 2021+</span></td><td><code>${esc(f.xlookup)}</code></td></tr>
    <tr><td><b>VLOOKUP</b><br><span class="muted small">sab versions</span></td><td><code>${esc(f.vlookup)}</code></td></tr>
    <tr><td><b>INDEX/MATCH</b><br><span class="muted small">purane Excel</span></td><td><code>${esc(f.index_match)}</code></td></tr>
    <tr><td><b>IFERROR + VLOOKUP</b><br><span class="muted small">#N/A handle</span></td><td><code>${esc(f.iferror_vlookup)}</code></td></tr>
  </tbody></table>
  <div class="muted-box" style="margin-top:8px"><b>Samjhein:</b> <code>${esc(f.from_col)}</code> ki value ko lookup sheet ke <code>${esc(f.lookup_key)}</code> column me dhoondo (exact match), aur usi row ka <code>${esc(f.return_col)}</code> wapas lao. <code>FALSE</code> = exact match (approximate kabhi nahi). Not Found par XLOOKUP "Not Found" deta hai; VLOOKUP #N/A deta hai is liye IFERROR version use karein.</div>` : "";
  return statsCards +
    `<div class="section-h">👁 Preview (pehli 50 rows)</div>` + preview + fx +
    `<div style="margin-top:16px"><button class="btn btn-primary btn-lg" id="lkApply">✅ Apply & Re-analyze — naya column "${esc(f.return_col || "")}" poore analysis me add karo</button></div>`;
}
