"""
GovData Analytics - Excel Report Builder
=========================================
Builds Government_Analytics_Report.xlsx with professional formatting:
Read_Me, Dashboard (KPIs + native charts), Summary, Pivot_Analysis,
Rankings, Exception_Report, Data_Quality, Change_Log, Assumptions,
Calculations, Raw_Data, Cleaned_Data, Chart_Data.

Rules: raw data written as-is; cleaned data separated; every sheet
formatted with frozen headers, filters and proper number formats.
"""

import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

# ------------------------- style tokens ------------------------- #
NAVY = '1F4E79'
GREEN = '2E7D32'
GOLD = 'C9A227'
RED = 'C0392B'
AMBER = 'F39C12'
LIGHT = 'DCE6F1'
GREY = 'F2F2F2'

THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

H_FONT = Font(bold=True, color='FFFFFF', size=11)
H_FILL = PatternFill('solid', fgColor=NAVY)
TITLE_FONT = Font(bold=True, size=16, color=NAVY)
SUB_FONT = Font(size=10, color='595959', italic=True)
SEC_FONT = Font(bold=True, size=12, color='FFFFFF')
SEC_FILL = PatternFill('solid', fgColor=GREEN)

SEV_FILL = {
    'CRITICAL': PatternFill('solid', fgColor='F8CBAD'),
    'WARNING': PatternFill('solid', fgColor='FFE699'),
    'INFO': PatternFill('solid', fgColor='E2EFDA'),
}


def _sheet_title(ws, title, subtitle=None, span=8):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(span, 2))
    c = ws.cell(row=1, column=1, value=title)
    c.font = TITLE_FONT
    ws.row_dimensions[1].height = 26
    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(span, 2))
        s = ws.cell(row=2, column=1, value=subtitle)
        s.font = SUB_FONT
    return 4 if subtitle else 3


GREY_FILL = PatternFill('solid', fgColor=GREY)
ZEBRA_THRESHOLD = 2000      # above this row-count, skip per-cell styling (speed)
FORMAT_THRESHOLD = 8000     # above this, skip per-cell number formats on data sheets


def _col_widths(df, headers_len=None):
    """Compute column widths from the data itself (no cell iteration)."""
    widths = {}
    for j, c in enumerate(df.columns):
        w = len(str(c))
        sample = df[c].head(300)
        for v in sample:
            if v is None:
                continue
            try:
                if v != v:
                    continue
            except Exception:
                pass
            if hasattr(v, 'strftime'):
                w = max(w, 11)
            else:
                w = max(w, len(str(v)))
        widths[j] = max(10, min(42, w + 3))
    return widths


def _write_table(ws, df, start_row, start_col=1, number_cols=None, pct_cols=None,
                 date_cols=None, max_rows=60000, zebra=True):
    """Write a DataFrame as a formatted table. Returns next free row."""
    number_cols = set(number_cols or [])
    pct_cols = set(pct_cols or [])
    date_cols = set(date_cols or [])
    if len(df) > max_rows:
        df = df.head(max_rows)
    headers = list(df.columns)
    widths = _col_widths(df)

    for j, h in enumerate(headers):
        c = ws.cell(row=start_row, column=start_col + j, value=str(h))
        c.font = H_FONT
        c.fill = H_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        letter = get_column_letter(start_col + j)
        if letter not in ws.column_dimensions or (ws.column_dimensions[letter].width or 0) < widths[j]:
            ws.column_dimensions[letter].width = widths[j]

    big = len(df) > ZEBRA_THRESHOLD
    if big:
        # fast path: bulk append, no per-cell borders/zebra
        df2 = df.astype(object).where(df.notna(), None)
        for rec in df2.itertuples(index=False, name=None):
            ws.append(list(rec))
        # number formats only when affordable
        if len(df) <= FORMAT_THRESHOLD:
            for j, h in enumerate(headers):
                if h in number_cols or h in pct_cols or h in date_cols:
                    fmt = '0.0"%"' if h in pct_cols else ('DD-MMM-YYYY' if h in date_cols else '#,##0.00')
                    col = start_col + j
                    for cell in ws.iter_rows(min_row=start_row + 1, max_row=start_row + len(df),
                                             min_col=col, max_col=col):
                        cell[0].number_format = fmt
        return start_row + len(df) + 2

    for i, (_, row) in enumerate(df.iterrows()):
        r = start_row + 1 + i
        for j, h in enumerate(headers):
            v = row[h]
            if isinstance(v, float) and v != v:
                v = None
            if isinstance(v, bool):
                v = 'YES' if v else 'NO'
            c = ws.cell(row=r, column=start_col + j, value=v)
            c.border = BORDER
            if zebra and i % 2 == 1:
                c.fill = GREY_FILL
            if h in date_cols or (hasattr(v, 'date') and not isinstance(v, str)):
                try:
                    c.number_format = 'DD-MMM-YYYY'
                except Exception:
                    pass
            elif h in pct_cols:
                c.number_format = '0.0"%"'
            elif h in number_cols:
                c.number_format = '#,##0.00' if isinstance(v, float) and abs(v - round(v)) > 0.001 else '#,##0'
    return start_row + len(df) + 2


def _autofit(ws, max_width=42, min_width=10):
    for col in ws.columns:
        best = 0
        letter = None
        for c in col:
            if c.value is None:
                continue
            letter = letter or c.column_letter
            try:
                l = len(str(c.value))
            except Exception:
                l = 10
            if l > best:
                best = l
        if letter:
            ws.column_dimensions[letter].width = max(min_width, min(max_width, best + 3))


# ------------------------------------------------------------------ #
def build_report(result, out_path):
    wb = Workbook()
    raw = result['raw']
    cleaned = result['cleaned']
    kpis = result['kpis']
    ctx = result.get('ctx', {})
    metric_cols = set()
    if result['main_metric']:
        metric_cols.add(result['main_metric'])
    for s in result['stats']:
        metric_cols.add(s['column'])
    if result['budget']:
        metric_cols.update(['Budget', 'Expenditure', 'Remaining', 'Utilization %'])
    date_cols = set(result['date_cols'])

    # ---------------- Read_Me ---------------- #
    ws = wb.active
    ws.title = 'Read_Me'
    r = _sheet_title(ws, 'GOVERNMENT ANALYTICS REPORT', 'Automated analysis generated by GovData Analytics Portal', span=6)
    info = [
        ('Source file', ctx.get('file_name', '-')),
        ('Sheet analysed', ctx.get('sheet_name', '-')),
        ('Records (rows)', kpis['total_records']),
        ('Fields (columns)', kpis['columns']),
        ('Report generated on', dt.datetime.now().strftime('%d-%b-%Y %H:%M')),
        ('Reporting basis', 'All figures computed directly from the uploaded file. No data was added, invented or estimated.'),
    ]
    for k, v in info:
        ws.cell(row=r, column=1, value=k).font = Font(bold=True)
        ws.cell(row=r, column=2, value=v)
        r += 1
    r += 1
    ws.cell(row=r, column=1, value='WORKBOOK CONTENTS').font = Font(bold=True, size=12, color=NAVY)
    r += 1
    contents = [
        ('Dashboard', 'Key performance indicators and charts (management view).'),
        ('Summary', 'Descriptive statistics, budget position, status distribution.'),
        ('Pivot_Analysis', 'Dimension-wise summaries: count, sum, average, min, max, median, share %.'),
        ('Data_Dictionary', 'Data model: every column with its detected role, type and notes.'),
        ('Rankings', 'Top 50 and Bottom 50 records by the main measure.'),
        ('Exception_Report', 'Duplicates, statistical anomalies, above-budget records.'),
        ('Validation_Rules', 'Data-validation rules + dropdown value lists (applied in Cleaned_Data).'),
        ('Power_Query', 'Refreshable Power Query (M) script + manual Excel steps.'),
        ('Formulas', 'VLOOKUP / XLOOKUP / SUMIFS / RANK recipes built from your real ranges + live demo.'),
        ('Data_Quality', 'Every data-quality issue detected, with severity.'),
        ('Change_Log', 'Audit trail of every change made during cleaning.'),
        ('Assumptions', 'Interpretation assumptions, stated explicitly.'),
        ('Calculations', 'Row-level flags: rank, duplicates, outliers, missing counts.'),
        ('Merged_Data', 'Present only when a sheet join was performed (Data Model feature).'),
        ('Raw_Data', 'Your original data, exactly as uploaded (unmodified).'),
        ('Cleaned_Data', 'Standardised copy (trimmed text, unified case, parsed dates) with data validation.'),
        ('Chart_Data', 'Backing data used by dashboard charts.'),
    ]
    for name, desc in contents:
        ws.cell(row=r, column=1, value=name).font = Font(bold=True)
        ws.cell(row=r, column=2, value=desc)
        r += 1
    r += 1
    ws.cell(row=r, column=1, value='METHODOLOGY').font = Font(bold=True, size=12, color=NAVY)
    r += 1
    method = [
        '1. The workbook was inspected sheet-by-sheet; the data sheet with the most rows was analysed by default.',
        '2. Exact duplicate rows and duplicate key values were detected and reported (never silently deleted).',
        '3. Text was trimmed and case-standardised; every change is listed in Change_Log.',
        '4. Dates were standardised; ambiguous dates were read day-first (dd/mm/yyyy) and the convention is stated in Assumptions.',
        '5. Numeric outliers were flagged using the IQR method. A flag means "requires review", not proof of error.',
        '6. Budget analysis treats Budget and Expenditure as distinct concepts and never assumes they are interchangeable.',
        '7. Raw_Data is a preserved copy; Cleaned_Data contains the standardised version used for analysis.',
    ]
    for m in method:
        ws.cell(row=r, column=1, value=m)
        r += 1
    r += 1
    ws.cell(row=r, column=1, value='CONFIDENTIALITY: This report contains the data you uploaded. Handle according to your office\'s official-record rules.').font = Font(italic=True, color=RED)
    _autofit(ws)
    ws.column_dimensions['B'].width = 80

    # ---------------- Dashboard ---------------- #
    dash = wb.create_sheet('Dashboard')
    r = _sheet_title(dash, 'MANAGEMENT DASHBOARD', ctx.get('file_name', ''), span=10)
    cards = [
        ('Total Records', kpis['total_records'], '#,##0'),
        ('Duplicate Rows', kpis['duplicate_rows'], '#,##0'),
        ('Missing Cells', kpis['missing_cells'], '#,##0'),
        ('Anomalies (IQR)', kpis['outliers'], '#,##0'),
    ]
    if kpis.get('main_metric') and kpis.get('main_metric_total') is not None:
        cards.append((f"Total {kpis['main_metric']}", kpis['main_metric_total'], '#,##0.00'))
    if kpis.get('utilization') is not None:
        cards.append(('Budget Utilization', kpis['utilization'], '0.0"%"'))
        cards.append(('Remaining Budget', kpis['total_budget'] - kpis['total_expenditure'], '#,##0.00'))
    if kpis.get('date_range'):
        cards.append(('Date Range', f"{kpis['date_range'][0]}  to  {kpis['date_range'][1]}", None))
    if kpis.get('top_category'):
        tc = kpis['top_category']
        cards.append((f"Top {tc['column']}", f"{tc['label']} ({tc['count']} records)", None))
    # draw cards 4 per row
    rr = r
    for i in range(0, len(cards), 4):
        batch = cards[i:i + 4]
        for j, (label, value, fmt) in enumerate(batch):
            c1 = dash.cell(row=rr, column=1 + j * 3, value=label.upper())
            c1.font = Font(bold=True, size=9, color='FFFFFF')
            c1.fill = PatternFill('solid', fgColor=NAVY)
            c1.alignment = Alignment(horizontal='center')
            c2 = dash.cell(row=rr + 1, column=1 + j * 3, value=value)
            c2.font = Font(bold=True, size=14, color=NAVY)
            c2.alignment = Alignment(horizontal='center')
            if fmt:
                c2.number_format = fmt
            dash.merge_cells(start_row=rr, start_column=1 + j * 3, end_row=rr, end_column=3 + j * 3)
            dash.merge_cells(start_row=rr + 1, start_column=1 + j * 3, end_row=rr + 1, end_column=3 + j * 3)
        rr += 3
    for col in range(1, 13):
        dash.column_dimensions[get_column_letter(col)].width = 12
    dash.sheet_view.showGridLines = False

    # management summary block (auto executive summary)
    mgmt = result.get('mgmt_summary') or []
    if mgmt:
        rr += 1
        c = dash.cell(row=rr, column=1, value='MANAGEMENT SUMMARY (auto-generated from the data)')
        c.font = Font(bold=True, size=12, color='FFFFFF')
        c.fill = SEC_FILL
        dash.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=12)
        rr += 1
        for line in mgmt:
            cell = dash.cell(row=rr, column=1, value='• ' + line)
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            dash.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=12)
            dash.row_dimensions[rr].height = 16
            rr += 1
        rr += 1

    # chart data sheet
    cd = wb.create_sheet('Chart_Data')
    cdr = _sheet_title(cd, 'Chart backing data', 'Referenced by Dashboard charts', span=6)

    charts_anchor_row = rr + 1

    # chart 1: dimension bar (top categories by metric) or pivot first dim
    anchors = []
    if result['pivots']:
        p = result['pivots'][0]
        cd.cell(row=cdr, column=1, value=f"{p['dimension']} - {'Sum of ' + p['measure'] if p['has_measure'] else 'Record count'}").font = Font(bold=True)
        hdr = cdr + 1
        cd.cell(row=hdr, column=1, value=p['dimension'])
        cd.cell(row=hdr, column=2, value=p['measure'] or 'Records')
        for i, rowd in enumerate(p['rows'][:10]):
            cd.cell(row=hdr + 1 + i, column=1, value=rowd['label'])
            cd.cell(row=hdr + 1 + i, column=2, value=rowd.get('sum') if p['has_measure'] else rowd.get('count'))
        n = min(10, len(p['rows']))
        bar = BarChart()
        bar.type = 'col'
        bar.title = f"{'Sum of ' + p['measure'] + ' by ' if p['has_measure'] else 'Records by '}{p['dimension']} (Top {n})"
        bar.height, bar.width = 8, 16
        data = Reference(cd, min_col=2, min_row=hdr, max_row=hdr + n)
        cats = Reference(cd, min_col=1, min_row=hdr + 1, max_row=hdr + n)
        bar.add_data(data, titles_from_data=True)
        bar.set_categories(cats)
        bar.legend = None
        dash.add_chart(bar, f'A{charts_anchor_row}')
        anchors.append(charts_anchor_row)
        cdr = hdr + n + 3

    # chart 2: status pie
    if result['status_dist']:
        sd = result['status_dist']
        cd.cell(row=cdr, column=1, value=f"Status distribution ({sd['column']})").font = Font(bold=True)
        hdr = cdr + 1
        cd.cell(row=hdr, column=1, value='Status')
        cd.cell(row=hdr, column=2, value='Records')
        for i, rowd in enumerate(sd['rows'][:8]):
            cd.cell(row=hdr + 1 + i, column=1, value=rowd['label'])
            cd.cell(row=hdr + 1 + i, column=2, value=rowd['count'])
        n = min(8, len(sd['rows']))
        pie = PieChart()
        pie.title = f'Status distribution ({sd["column"]})'
        pie.height, pie.width = 8, 8
        data = Reference(cd, min_col=2, min_row=hdr, max_row=hdr + n)
        cats = Reference(cd, min_col=1, min_row=hdr + 1, max_row=hdr + n)
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(cats)
        dash.add_chart(pie, f'H{charts_anchor_row}')
        cdr = hdr + n + 3

    # chart 3: monthly trend line
    if result['trend']:
        tr = result['trend']
        cd.cell(row=cdr, column=1, value=f"Monthly trend ({tr['measure'] or 'records'})").font = Font(bold=True)
        hdr = cdr + 1
        cd.cell(row=hdr, column=1, value='Month')
        cd.cell(row=hdr, column=2, value=tr['measure'] or 'Records')
        for i, lab in enumerate(tr['labels']):
            cd.cell(row=hdr + 1 + i, column=1, value=str(lab))
            cd.cell(row=hdr + 1 + i, column=2, value=tr['values'][i])
        n = len(tr['labels'])
        if n >= 2:
            line = LineChart()
            line.title = f"Monthly trend - {tr['measure'] or 'record count'}"
            line.height, line.width = 8, 16
            data = Reference(cd, min_col=2, min_row=hdr, max_row=hdr + n)
            cats = Reference(cd, min_col=1, min_row=hdr + 1, max_row=hdr + n)
            line.add_data(data, titles_from_data=True)
            line.set_categories(cats)
            line.legend = None
            dash.add_chart(line, f'A{charts_anchor_row + 17}')
            cdr = hdr + n + 3

    # chart 4: budget vs expenditure by dept
    if result['budget'] and result['budget'].get('per_dept'):
        bd = result['budget']['per_dept'][:10]
        cd.cell(row=cdr, column=1, value='Budget vs Expenditure (by department)').font = Font(bold=True)
        hdr = cdr + 1
        cd.cell(row=hdr, column=1, value='Department')
        cd.cell(row=hdr, column=2, value='Budget')
        cd.cell(row=hdr, column=3, value='Expenditure')
        for i, rowd in enumerate(bd):
            cd.cell(row=hdr + 1 + i, column=1, value=rowd['label'])
            cd.cell(row=hdr + 1 + i, column=2, value=rowd['budget'])
            cd.cell(row=hdr + 1 + i, column=3, value=rowd['expenditure'])
        n = len(bd)
        bar2 = BarChart()
        bar2.type = 'col'
        bar2.title = 'Budget vs Expenditure by department'
        bar2.height, bar2.width = 8, 16
        data = Reference(cd, min_col=2, max_col=3, min_row=hdr, max_row=hdr + n)
        cats = Reference(cd, min_col=1, min_row=hdr + 1, max_row=hdr + n)
        bar2.add_data(data, titles_from_data=True)
        bar2.set_categories(cats)
        dash.add_chart(bar2, f'H{charts_anchor_row + 17}')
    _autofit(cd)

    # ---------------- Summary ---------------- #
    ws = wb.create_sheet('Summary')
    r = _sheet_title(ws, 'SUMMARY & DESCRIPTIVE STATISTICS', 'Computed from Cleaned_Data', span=8)
    if result['stats']:
        ws.cell(row=r, column=1, value='Descriptive statistics (numeric fields)').font = Font(bold=True, size=12, color=NAVY)
        r += 1
        sdf = _stats_df(result['stats'])
        r = _write_table(ws, sdf, r, number_cols={'Sum', 'Mean', 'Median', 'Min', 'Max', 'Std Dev'})
        ws.freeze_panes = None
    if result['budget']:
        ws.cell(row=r, column=1, value='Budget position').font = Font(bold=True, size=12, color=NAVY)
        r += 1
        b = result['budget']
        rows = [('Total Budget (' + b['budget_col'] + ')', b['total_budget']),
                ('Total Expenditure (' + b['expenditure_col'] + ')', b['total_expenditure']),
                ('Remaining', b['remaining']),
                ('Utilization %', b['utilization'])]
        for k, v in rows:
            ws.cell(row=r, column=1, value=k).font = Font(bold=True)
            c = ws.cell(row=r, column=2, value=v)
            if 'Utilization' in k:
                c.number_format = '0.0"%"'
            else:
                c.number_format = '#,##0.00'
            r += 1
        if b['per_dept']:
            r += 1
            ws.cell(row=r, column=1, value='Department-wise budget utilisation').font = Font(bold=True, size=12, color=NAVY)
            r += 1
            hdr_row = r
            bdf = _budget_df(b['per_dept'])
            r = _write_table(ws, bdf, r, number_cols={'Budget', 'Expenditure', 'Remaining', 'Records'}, pct_cols={'Utilization %'})
            try:
                ws.conditional_formatting.add(
                    f'F{hdr_row + 1}:F{hdr_row + len(bdf)}',
                    ColorScaleRule(start_type='num', start_value=0, start_color='F8696B',
                                   mid_type='num', mid_value=50, mid_color='FFEB84',
                                   end_type='num', end_value=100, end_color='63BE7B'))
                ws.conditional_formatting.add(
                    f'F{hdr_row + 1}:F{hdr_row + len(bdf)}',
                    CellIsRule(operator='greaterThan', formula=['100'], fill=PatternFill('solid', fgColor='F8CBAD'), font=Font(color='9C0006', bold=True)))
            except Exception:
                pass
    if result['status_dist']:
        ws.cell(row=r, column=1, value=f"Status distribution ({result['status_dist']['column']})").font = Font(bold=True, size=12, color=NAVY)
        r += 1
        sdf2 = _status_df(result['status_dist'])
        r = _write_table(ws, sdf2, r, number_cols={'Records'}, pct_cols={'Share %'})
    _autofit(ws)

    # ---------------- Pivot_Analysis ---------------- #
    ws = wb.create_sheet('Pivot_Analysis')
    r = _sheet_title(ws, 'PIVOT-STYLE SUMMARIES', 'Group-wise aggregation of the cleaned data', span=8)
    for p in result['pivots']:
        ws.cell(row=r, column=1, value=f"{p['dimension'].upper()} - " +
                 (f"Sum / average / min / max / median of {p['measure']}" if p['has_measure'] else 'Record counts')).font = Font(bold=True, size=12, color='FFFFFF')
        ws.cell(row=r, column=1).fill = SEC_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        r += 1
        pdf = _pivot_df(p)
        r = _write_table(ws, pdf, r, number_cols={'Sum', 'Average', 'Min', 'Max', 'Median'}, pct_cols={'Share %'})
        r += 1
    _autofit(ws)

    # ---------------- Rankings ---------------- #
    ws = wb.create_sheet('Rankings')
    r = _sheet_title(ws, 'RANKINGS', f"Ranked by {result['rank_info']['metric']} - {result['rank_info']['direction']}" if result['rank_info'] else '', span=8)
    if result['top50']:
        ws.cell(row=r, column=1, value='TOP 50 (highest first)').font = Font(bold=True, size=12, color=GREEN)
        r += 1
        tdf = _rank_df(result['top50'])
        numc = {c for c in tdf.columns if c in metric_cols or c == 'Rank'}
        r = _write_table(ws, tdf, r, number_cols=numc)
        r += 1
    if result['bottom50']:
        ws.cell(row=r, column=1, value='BOTTOM 50 (lowest first)').font = Font(bold=True, size=12, color=RED)
        r += 1
        bdf = _rank_df(result['bottom50'])
        numc = {c for c in bdf.columns if c in metric_cols or c == 'Rank'}
        r = _write_table(ws, bdf, r, number_cols=numc)
    _autofit(ws)

    # ---------------- Exception_Report ---------------- #
    ws = wb.create_sheet('Exception_Report')
    r = _sheet_title(ws, 'EXCEPTION REPORT', 'Items flagged for review. A flag is not proof of any wrongdoing or error.', span=8)
    ws.cell(row=r, column=1, value='1. Exact duplicate records').font = Font(bold=True, size=12, color=NAVY)
    r += 1
    if result['dup_rows']:
        ws.cell(row=r, column=1, value=f"{result['dup_rows']} exact duplicate row(s) found in Raw_Data. First 10 listed below (row numbers refer to Raw_Data).").font = SUB_FONT
        r += 1
        ddf = _dup_df(result['dup_records'])
        r = _write_table(ws, ddf, r)
    else:
        ws.cell(row=r, column=1, value='None detected.')
        r += 2
    if result.get('id_dup_info'):
        d = result['id_dup_info']
        ws.cell(row=r, column=1, value=f"2. Duplicate key values in \"{d['column']}\" ({d['distinct']} value(s), {d['count']} rows involved)").font = Font(bold=True, size=12, color=NAVY)
        r += 1
        for t in d['top']:
            ws.cell(row=r, column=1, value=f"Value \"{t['value']}\" appears {t['times']} time(s)")
            r += 1
        r += 1
    ws.cell(row=r, column=1, value='3. Statistical anomalies (IQR method)').font = Font(bold=True, size=12, color=NAVY)
    r += 1
    if result['outliers']:
        ws.cell(row=r, column=1, value=f"{result['outlier_total']} value(s) outside the typical range. First entries below - each requires review.").font = SUB_FONT
        r += 1
        odf = _outlier_df(result['outliers'])
        r = _write_table(ws, odf, r, number_cols={'Value'})
        for row_cells in ws.iter_rows(min_row=r - len(odf) - 1, max_row=r - 2, min_col=4, max_col=4):
            for c in row_cells:
                c.fill = PatternFill('solid', fgColor='FFF2CC')
    else:
        ws.cell(row=r, column=1, value='None detected.')
        r += 2
    if result['budget'] and result['budget']['above_budget_rows']:
        ws.cell(row=r, column=1, value='4. Records where expenditure exceeds budget').font = Font(bold=True, size=12, color=NAVY)
        r += 1
        abdf = _above_df(result['budget']['above_budget_rows'])
        r = _write_table(ws, abdf, r, number_cols={'Budget', 'Expenditure'})
    _autofit(ws)

    # ---------------- Formulas (VLOOKUP / XLOOKUP recipes) ---------------- #
    ws = wb.create_sheet('Formulas')
    r = _sheet_title(ws, 'EXCEL FORMULA RECIPES',
                     'Aap ke workbook ke ASLI sheet names, columns aur ranges se bani hui - copy kar ke paste karein', span=6)
    recipes = result.get('formulas') or []
    if recipes:
        # live demo: working formula inside THIS report (Raw_Data reference)
        try:
            raw_cols = list(raw.columns)
            key_col = next((c for c in raw_cols if result['col_meta'].get(c, {}).get('role') == 'id'), raw_cols[0])
            met = result.get('main_metric')
            if met and key_col in raw_cols and met in raw_cols:
                kL = get_column_letter(raw_cols.index(key_col) + 1)
                mL = get_column_letter(raw_cols.index(met) + 1)
                last = 3 + len(raw)
                sample = raw[key_col].dropna()
                sample_val = sample.iloc[0] if len(sample) else None
                if sample_val is not None:
                    ws.cell(row=r, column=1, value='LIVE DEMO (ye formula isi report me chalta hai):').font = Font(bold=True, size=11, color=GREEN)
                    r += 1
                    ws.cell(row=r, column=1, value=f'Key: {sample_val}')
                    demo = ws.cell(row=r, column=2, value=f'=VLOOKUP(A{r}, Raw_Data!${kL}$4:${kL}${last}, {raw_cols.index(met) + 1}, FALSE)')
                    demo.font = Font(name='Consolas', size=10, bold=True, color=NAVY)
                    ws.cell(row=r, column=3, value=f'<- "{met}" ki value "{key_col}" = {sample_val} ke liye (Raw_Data sheet se)')
                    r += 2
        except Exception:
            r += 1
        hdr = ['Formula type', 'Formula (copy-paste)', 'Kya karta hai', 'Wazahat / parts']
        for j, h in enumerate(hdr):
            c = ws.cell(row=r, column=1 + j, value=h)
            c.font = H_FONT; c.fill = H_FILL; c.border = BORDER
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        r += 1
        for i, rec in enumerate(recipes):
            c1 = ws.cell(row=r, column=1, value=rec['section'])
            c1.font = Font(bold=True, size=10)
            c2 = ws.cell(row=r, column=2, value=rec['formula'])
            c2.font = Font(name='Consolas', size=10, color='0B3D2E')
            ws.cell(row=r, column=3, value=rec['purpose'])
            ws.cell(row=r, column=4, value=rec['explain']).alignment = Alignment(wrap_text=True, vertical='top')
            for j in range(1, 5):
                ws.cell(row=r, column=j).border = BORDER
                if i % 2 == 1:
                    ws.cell(row=r, column=j).fill = GREY_FILL
            ws.row_dimensions[r].height = 42
            r += 1
        ws.column_dimensions['B'].width = 60
        ws.column_dimensions['A'].width = 26
        ws.column_dimensions['C'].width = 38
        ws.column_dimensions['D'].width = 70
        r += 1
        ws.cell(row=r, column=1, value='NOTE: XLOOKUP / MAXIFS / MINIFS ke liye Excel 365 ya 2019+ chahiye. VLOOKUP me hamesha FALSE (exact match). "$" absolute references drag karne par ranges fix rakhte hain.').font = SUB_FONT
    else:
        ws.cell(row=r, column=1, value='Is dataset ke liye lookup formulas generate nahi ho sakin (key/measure columns nahi mile).')
    ws.freeze_panes = 'A4'

    # ---------------- Data_Quality ---------------- #
    ws = wb.create_sheet('Data_Quality')
    r = _sheet_title(ws, 'DATA QUALITY REPORT', 'Every issue detected in the uploaded file, with severity', span=6)
    ws.cell(row=r, column=1, value='Severity')
    ws.cell(row=r, column=2, value='Area')
    ws.cell(row=r, column=3, value='Issue')
    ws.cell(row=r, column=4, value='Count')
    for j in range(1, 5):
        c = ws.cell(row=r, column=j)
        c.font = H_FONT
        c.fill = H_FILL
        c.border = BORDER
    r += 1
    for iss in result['issues']:
        ws.cell(row=r, column=1, value=iss['severity']).fill = SEV_FILL.get(iss['severity'], PatternFill())
        ws.cell(row=r, column=2, value=iss['area'])
        ws.cell(row=r, column=3, value=iss['description'])
        ws.cell(row=r, column=4, value=iss['count'])
        for j in range(1, 5):
            ws.cell(row=r, column=j).border = BORDER
            if j == 4:
                ws.cell(row=r, column=j).number_format = '#,##0'
        r += 1
    ws.column_dimensions['C'].width = 90
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['D'].width = 10
    ws.freeze_panes = 'A5'

    # ---------------- Change_Log ---------------- #
    ws = wb.create_sheet('Change_Log')
    r = _sheet_title(ws, 'CHANGE LOG (AUDIT TRAIL)', 'Every change made while producing Cleaned_Data. Raw_Data remains untouched.', span=6)
    if result['change_log']:
        cdf = _change_df(result['change_log'])
        r = _write_table(ws, cdf, r, number_cols={'Cells affected'})
    else:
        ws.cell(row=r, column=1, value='No changes were required - the data was already clean.')
    _autofit(ws)

    # ---------------- Assumptions ---------------- #
    ws = wb.create_sheet('Assumptions')
    r = _sheet_title(ws, 'ASSUMPTIONS & NOTES', 'Interpretation decisions made during the analysis', span=4)
    for i, a in enumerate(result['assumptions'], 1):
        ws.cell(row=r, column=1, value=i).font = Font(bold=True)
        ws.cell(row=r, column=2, value=a)
        ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical='top')
        r += 1
    ws.column_dimensions['B'].width = 100

    # ---------------- Data_Dictionary (data model) ---------------- #
    ws = wb.create_sheet('Data_Dictionary')
    r = _sheet_title(ws, 'DATA DICTIONARY / DATA MODEL', 'Auto-detected structure of the analysed sheet', span=8)
    idc = next((c for c in result['col_meta'] if result['col_meta'][c]['role'] == 'id'), None)
    dd_rows = []
    for c in raw.columns:
        meta = result['col_meta'].get(c, {})
        prof = next((p for p in result['columns_profile'] if p['name'] == c), {})
        notes = []
        if idc == c:
            uniq_pct = (prof.get('unique', 0) / max(len(raw), 1)) * 100
            notes.append('Primary key candidate' if uniq_pct >= 95 else 'Key field - duplicates found')
        if c == result.get('main_metric'):
            notes.append('Main measure for rankings/summaries')
        if result.get('primary_date_col') == c:
            notes.append('Reference date column')
        if result.get('budget') and c in (result['budget']['budget_col'], result['budget']['expenditure_col']):
            notes.append('Used in budget analysis')
        dd_rows.append({
            'Column': c, 'Detected role': meta.get('role', ''), 'Data type': meta.get('dtype', ''),
            'Unique values': prof.get('unique'), 'Missing': prof.get('missing'),
            'Missing %': prof.get('missing_pct'),
            'Sample / range': ', '.join(prof.get('sample', [])[:3]) if prof.get('sample')
                               else (f"{prof.get('min')} … {prof.get('max')}" if prof.get('min') is not None else '—'),
            'Notes': '; '.join(notes) if notes else '',
        })
    import pandas as _pd
    ddf = _pd.DataFrame(dd_rows)
    r = _write_table(ws, ddf, r, number_cols={'Unique values', 'Missing', 'Missing %'})
    _autofit(ws)

    # ---------------- Validation_Rules ---------------- #
    ws = wb.create_sheet('Validation_Rules')
    r = _sheet_title(ws, 'DATA VALIDATION RULES', 'Generated from the actual data - applied as dropdowns/range checks in Cleaned_Data (rows below the data too)', span=6)
    rules = result.get('validation_rules', [])
    list_src_cells = {}
    if rules:
        vdf = _pd.DataFrame([{
            'Column': x['column'], 'Rule type': x['type'], 'Rule': x['rule'],
            'Purpose': x['purpose'],
        } for x in rules])
        r = _write_table(ws, vdf, r)
        # allowed-value lists (referenced by the dropdown validations)
        list_rules = [x for x in rules if x['type'] == 'list']
        if list_rules:
            lr = r + 1
            ws.cell(row=lr, column=1, value='ALLOWED VALUES (dropdown sources - do not delete)').font = Font(bold=True, size=11, color=NAVY)
            lr += 1
            list_src_cells = {}
            for j, x in enumerate(list_rules):
                col = get_column_letter(1 + j)
                hc = ws.cell(row=lr, column=1 + j, value=x['column'])
                hc.font = H_FONT
                hc.fill = H_FILL
                hc.border = BORDER
                for i, v in enumerate(x.get('values', [])[:250]):
                    ws.cell(row=lr + 1 + i, column=1 + j, value=v).border = BORDER
                nvals = min(len(x.get('values', [])), 250)
                list_src_cells[x['column']] = f"Validation_Rules!${col}${lr + 1}:${col}${lr + nvals}"
                ws.column_dimensions[col].width = 24
    else:
        ws.cell(row=r, column=1, value='No validation rules could be generated for this data.')
    _autofit(ws)

    # ---------------- Power_Query ---------------- #
    ws = wb.create_sheet('Power_Query')
    r = _sheet_title(ws, 'POWER QUERY SCRIPT (M CODE)', 'Refreshable cleaning recipe matching this analysis - paste into Power Query Advanced Editor', span=6)
    pq = result.get('powerquery') or {}
    m_lines = (pq.get('m_code') or 'Not available for this dataset.').splitlines()
    mono = Font(name='Consolas', size=10)
    for i, line in enumerate(m_lines):
        c = ws.cell(row=r + i, column=1, value=line)
        c.font = mono
        c.alignment = Alignment(vertical='top')
    r += len(m_lines) + 2
    ws.cell(row=r, column=1, value='EXCEL STEPS (manual route)').font = Font(bold=True, size=12, color=NAVY)
    r += 1
    for i, step in enumerate(pq.get('ui_steps', []), 1):
        ws.cell(row=r, column=1, value=f'{i}. {step}').alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 18
        r += 1
    ws.column_dimensions['A'].width = 110
    ws.sheet_view.showGridLines = False

    # ---------------- Calculations ---------------- #
    ws = wb.create_sheet('Calculations')
    _sheet_title(ws, 'ROW-LEVEL CALCULATIONS', 'Rank, duplicate flag, outlier flag and missing-field count per record', span=8)
    calc = result['calc'].copy()
    _write_table(ws, calc, 3, number_cols={c for c in calc.columns if 'Rank' in str(c) or 'count' in str(c).lower()})
    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f'A3:{get_column_letter(len(calc.columns))}{3 + len(calc)}'

    # ---------------- Raw_Data ---------------- #
    ws = wb.create_sheet('Raw_Data')
    _sheet_title(ws, 'RAW DATA (UNMODIFIED)', 'Exactly as uploaded - preserved for audit', span=10)
    _write_table(ws, raw, 3, date_cols=date_cols)
    ws.freeze_panes = 'A4'
    if len(raw.columns):
        ws.auto_filter.ref = f'A3:{get_column_letter(len(raw.columns))}{3 + len(raw)}'

    # ---------------- Cleaned_Data ---------------- #
    ws = wb.create_sheet('Cleaned_Data')
    _sheet_title(ws, 'CLEANED DATA', 'Standardised copy: trimmed text, unified case, parsed dates, MISSING markers (see Change_Log). Data validation (dropdowns/ranges) applied for future edits.', span=10)
    _write_table(ws, cleaned, 3, date_cols=date_cols,
                 number_cols={c for c in cleaned.columns if str(cleaned[c].dtype) in ('int64', 'float64', 'Int64')})
    ws.freeze_panes = 'A4'
    if len(cleaned.columns):
        ws.auto_filter.ref = f'A3:{get_column_letter(len(cleaned.columns))}{3 + len(cleaned)}'
    # apply Excel data validation (dropdowns / range checks) to data + 300 future rows
    if rules:
        from openpyxl.worksheet.datavalidation import DataValidation
        end_row = 3 + len(cleaned) + 300
        for j, c in enumerate(cleaned.columns, 1):
            rule = next((x for x in rules if x['column'] == c), None)
            if not rule:
                continue
            letter = get_column_letter(j)
            try:
                if rule['type'] == 'list' and c in list_src_cells:
                    dv = DataValidation(type='list', formula1=list_src_cells[c], allow_blank=True,
                                        showErrorMessage=True, errorTitle='Invalid entry',
                                        error='Please pick a value from the dropdown list.')
                elif rule['type'] == 'decimal':
                    dv = DataValidation(type='decimal', operator='between',
                                        formula1=str(rule['min']), formula2=str(rule['max']), allow_blank=True,
                                        showErrorMessage=True, errorTitle='Out of range',
                                        error=f"Value must be between {rule['min']} and {rule['max']}.")
                elif rule['type'] == 'date':
                    y1, m1, d1 = rule['min'].split('-')
                    y2, m2, d2 = rule['max'].split('-')
                    dv = DataValidation(type='date', operator='between',
                                        formula1=f'DATE({y1},{int(m1)},{int(d1)})',
                                        formula2=f'DATE({y2},{int(m2)},{int(d2)})', allow_blank=True,
                                        showErrorMessage=True, errorTitle='Invalid date',
                                        error=f"Date must be between {rule['min']} and {rule['max']}.")
                elif rule['type'] == 'textLength':
                    dv = DataValidation(type='textLength', operator='between',
                                        formula1=str(rule['min']), formula2=str(rule['max']), allow_blank=True,
                                        showErrorMessage=True, errorTitle='Invalid length',
                                        error=f"Text length must be between {rule['min']} and {rule['max']} characters.")
                else:
                    continue
                ws.add_data_validation(dv)
                dv.add(f'{letter}4:{letter}{end_row}')
            except Exception:
                pass

    # ---------------- Merged_Data (join / lookup result) ---------------- #
    merged_df = result.get('merged_data')
    if merged_df is not None:
        jr = result.get('join_report') or {}
        lr = result.get('lookup_report') or {}
        if jr:
            subtitle = (f"{jr.get('left_sheet', '?')} [{jr.get('left_key', '?')}] "
                        f"{str(jr.get('join_type', 'left')).upper()}-joined with "
                        f"{jr.get('right_sheet', '?')} [{jr.get('right_key', '?')}] - {jr.get('result_rows', '?')} rows. "
                        f"Full report in Assumptions sheet.")
        elif lr:
            subtitle = (f"Lookup column \"{lr.get('new_column', '?')}\" fetched from \"{lr.get('lookup_sheet', '?')}\" "
                        f"({lr.get('lookup_key', '?')} = {lr.get('return_col', '?')}); "
                        f"matched {lr.get('matched', '?')}/{lr.get('total_rows', '?')} rows. "
                        f"Detail in Assumptions sheet.")
        else:
            subtitle = 'Combined dataset.'
        ws = wb.create_sheet('Merged_Data')
        _sheet_title(ws, 'COMBINED DATA (JOIN / LOOKUP RESULT)', subtitle, span=10)
        _write_table(ws, merged_df, 3,
                     number_cols={c for c in merged_df.columns if str(merged_df[c].dtype) in ('int64', 'float64', 'Int64')},
                     date_cols={c for c in merged_df.columns if 'datetime' in str(merged_df[c].dtype)})
        ws.freeze_panes = 'A4'
        if len(merged_df.columns):
            ws.auto_filter.ref = f'A3:{get_column_letter(len(merged_df.columns))}{3 + len(merged_df)}'

    # sheet order
    order = ['Read_Me', 'Dashboard', 'Summary', 'Pivot_Analysis', 'Data_Dictionary', 'Rankings',
             'Exception_Report', 'Validation_Rules', 'Power_Query', 'Formulas', 'Data_Quality', 'Change_Log',
             'Assumptions', 'Calculations', 'Merged_Data', 'Raw_Data', 'Cleaned_Data', 'Chart_Data']
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [s for s in wb._sheets if s.title not in order]
    wb.active = 0

    # print-friendly page setup: landscape, fit-to-width, repeated header rows
    from openpyxl.worksheet.properties import PageSetupProperties
    header_row = {'Data_Quality': '4:4', 'Raw_Data': '3:3', 'Cleaned_Data': '3:3', 'Calculations': '3:3'}
    for ws in wb.worksheets:
        try:
            ws.page_setup.orientation = 'landscape'
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
            if ws.title in header_row:
                ws.print_title_rows = header_row[ws.title]
        except Exception:
            pass

    wb.save(out_path)
    return out_path


# ------------------------- dataframe helpers ------------------------- #
def _stats_df(stats):
    import pandas as pd
    return pd.DataFrame([{
        'Field': s['column'], 'Count': s['count'], 'Sum': s['sum'], 'Mean': s['mean'],
        'Median': s['median'], 'Min': s['min'], 'Max': s['max'], 'Std Dev': s['std'],
    } for s in stats])


def _budget_df(per_dept):
    import pandas as pd
    return pd.DataFrame([{
        'Department': d['label'], 'Budget': d['budget'], 'Expenditure': d['expenditure'],
        'Remaining': d['remaining'], 'Utilization %': d['utilization'], 'Records': d['records'],
    } for d in per_dept])


def _status_df(sd):
    import pandas as pd
    total = sum(x['count'] for x in sd['rows']) or 1
    return pd.DataFrame([{'Status': x['label'], 'Records': x['count'],
                          'Share %': round(x['count'] / total * 100, 1)} for x in sd['rows']])


def _pivot_df(p):
    import pandas as pd
    if p['has_measure']:
        return pd.DataFrame([{
            p['dimension']: x['label'], 'Records': x['count'], 'Sum': x['sum'],
            'Average': x['avg'], 'Min': x.get('min'), 'Max': x.get('max'),
            'Median': x.get('median'), 'Share %': x['pct'],
        } for x in p['rows']])
    return pd.DataFrame([{
        p['dimension']: x['label'], 'Records': x['count'], 'Share %': x['pct'],
    } for x in p['rows']])


def _rank_df(top):
    import pandas as pd
    cols = ['Rank'] + [c for c in top['columns']]
    rows = []
    for r in top['rows']:
        rows.append({'Rank': r['rank'], **{c: r.get(c) for c in top['columns']}})
    return pd.DataFrame(rows)


def _outlier_df(outliers):
    import pandas as pd
    return pd.DataFrame([{'Row (Raw_Data)': o['row'], 'Field': o['column'],
                          'Value': o['value'], 'Review note': o['detail']} for o in outliers])


def _dup_df(dups):
    import pandas as pd
    if not dups:
        return pd.DataFrame()
    cols = ['Row (Raw_Data)'] + [k for k in dups[0].keys() if k != 'row']
    return pd.DataFrame([{'Row (Raw_Data)': d['row'], **{k: d[k] for k in cols[1:]}} for d in dups])


def _above_df(rows):
    import pandas as pd
    return pd.DataFrame([{'Row (Raw_Data)': a['row'], 'Budget': a['budget'],
                          'Expenditure': a['expenditure'], 'Review note': a['detail']} for a in rows])


def _change_df(log):
    import pandas as pd
    return pd.DataFrame([{
        'Field': c['field'], 'Original': c['original'], 'Updated': c['updated'],
        'Reason': c['reason'], 'Cells affected': c['cells'],
    } for c in log])
