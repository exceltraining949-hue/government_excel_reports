# GovData Analytics — Government Data Analytics Portal

Ek web-based portal jahan user **Excel/CSV khud upload** karta hai aur **AI Agent** (automated analytics engine)
poora workflow khud chalata hai — bilkul usi methodology ke mutabiq jo government-office reporting mein use hoti hai.

## Workflow (jo har upload par automatically chalta hai)

```
Upload → Workbook Inspection → Data Quality Checks → Cleaning (audit-trail ke saath)
→ Duplicates & Missing Analysis → Standardization → Pivot Summaries → Rankings (Top/Bottom 50)
→ Budget vs Expenditure → Monthly Trends → IQR Anomaly Detection → KPI Dashboard
→ Government_Analytics_Report.xlsx (13 sheets) → Download
```

## Core Rules (engine mein enforced)

- **Raw data kabhi modify nahi hota** — `Raw_Data` sheet bilkul wahi hai jo upload hui.
- **Har change ka audit trail** — `Change_Log` sheet mein field/original/updated/reason/cells.
- **Koi data fabricate nahi hota** — blank sirf "MISSING" mark hota hai (categorical columns, optional).
- **Anomalies neutral language** — "requires review", kisi par ilzaam nahi.
- **Budget ≠ Expenditure** — dono alag concepts, kabhi interchangeable assume nahi.
- **Duplicates default mein delete NAHI hote** — sirf report hote hain (user option se remove kar sakta hai).

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3.13 + Flask (port 7860) |
| Data engine | pandas + numpy (`analyzer.py`) |
| Excel writer | openpyxl (`report_builder.py`) — native charts, conditional formatting, freeze panes |
| Frontend | Vanilla HTML/CSS/JS (`templates/index.html`, `static/`) |
| Charts | Pure SVG/CSS (koi external dependency nahi) |

## Features

- **Drag & drop upload** (.xlsx, .xls, .xlsm, .csv — max 25 MB)
- **Auto sheet detection** — sab se bari data sheet analyze hoti hai; dropdown se koi bhi sheet re-analyze
- **Cleaning options** (user control): case/spacing standardize, MISSING marking, duplicate removal
- **KPI cards, 4 chart types** (bar, donut, line trend, budget-vs-expenditure)
- **AI Agent command box** — Roman Urdu/English commands: "top 10 dikhao", "department wise summary",
  "million me dikhao", "duplicates dikhao", "missing values", "outliers", "monthly trend", "budget vs expenditure"
- **Number format switcher** — Full / Thousands (K) / Millions (M) / Billions (B) — display only, values asli
- **Bilingual UI** — English + اردو toggle
- **Excel report** — 13 sheets: Read_Me, Dashboard (KPIs + 4 native charts), Summary, Pivot_Analysis,
  Rankings, Exception_Report, Data_Quality, Change_Log, Assumptions, Calculations, Raw_Data, Cleaned_Data, Chart_Data

## Excel Report Sheets

1. **Read_Me** — methodology, source info, sheet index, confidentiality note
2. **Dashboard** — KPI cards + BarChart (dept-wise), PieChart (status), LineChart (monthly trend), Budget vs Expenditure
3. **Summary** — descriptive stats, budget position with 3-color utilization scale
4. **Pivot_Analysis** — department/district/status/grade-wise tables (count, sum, avg, share %)
5. **Rankings** — Top 50 / Bottom 50 records
6. **Exception_Report** — duplicates, duplicate keys, IQR anomalies, above-budget records
7. **Data_Quality** — har issue severity ke saath (CRITICAL/WARNING/INFO)
8. **Change_Log** — poori audit trail
9. **Assumptions** — saare interpretation decisions
10. **Calculations** — row-level: rank, duplicate flag, outlier flag, missing count, KEPT/REMOVED
11. **Raw_Data** — original, unmodified
12. **Cleaned_Data** — standardized copy
13. **Chart_Data** — dashboard charts ka backing data

## Files

```
gov-analytics-app/
├── app.py              # Flask server + API
├── analyzer.py         # analysis engine (inspection → cleaning → analysis)
├── report_builder.py   # Excel report generator
├── templates/index.html
├── static/style.css
├── static/app.js       # UI logic + SVG charts + command interpreter
├── uploads/            # per-job: source file + generated report + payload
├── make_sample.py      # test data generator
└── test_pipeline.py    # direct pipeline test
```

## Run

```bash
cd gov-analytics-app && python app.py   # http://0.0.0.0:7860
```

## Column-Role Auto-Detection

Engine column names se roles detect karta hai (English patterns):
`id, date, department, district, status, vendor, designation, grade, category, gender,
budget, expenditure, amount, percent, name` — isi se pivots, rankings, budget analysis
aur trend khud decide hote hain. Numeric text amounts ("Rs. 45,000", "(500)", "1,50,000")
carefully parse hote hain — multiplier words ("2.5 million") deliberately reject hote hain
taake galat value na bane.
