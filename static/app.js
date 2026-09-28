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
let CURRENT_TAB = "quality";

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
  renderKpis();
  renderCharts();
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
        [{ label: pv.dimension }, { label: "Records", num: true }, { label: "Sum", num: true }, { label: "Average", num: true }, { label: "Share %", num: true }],
        pv.rows.map(r => [esc(r.label), r.count.toLocaleString(), fmtNum(r.sum), fmtNum(r.avg), fmtPct(r.pct)])
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

/* ---------------- number format bar ---------------- */
document.querySelectorAll(".fmt-btn").forEach(b => b.addEventListener("click", () => {
  document.querySelectorAll(".fmt-btn").forEach(x => x.classList.remove("active"));
  b.classList.add("active");
  NUMFMT = b.dataset.fmt;
  renderKpis(); renderCharts(); renderTab(CURRENT_TAB);
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
    if (p.trend) { document.querySelector(".charts-grid").scrollIntoView({ behavior: "smooth" }); agentSay(`Monthly trend chart dikhla di — "${esc(p.trend.date_col)}" ke hisab se${p.trend.measure ? `, ${esc(p.trend.measure)} ka` : ""}.`); }
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
  agentSay('Ye command samajh nahi aayi. Yehi try karein: <b>"top 10"</b>, <b>"department wise summary"</b>, <b>"duplicates dikhao"</b>, <b>"missing values"</b>, <b>"outliers"</b>, <b>"monthly trend"</b>, <b>"million me dikhao"</b>, <b>"budget vs expenditure"</b>, <b>"summary"</b>.');
}
