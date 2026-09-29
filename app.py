"""
GovData Analytics Portal - Flask server
Upload -> automated analysis -> dashboard JSON -> Excel report download.
"""
import os
import json
import uuid
import traceback
import datetime as dt

import pandas as pd
from flask import Flask, request, jsonify, render_template, send_file, abort

from analyzer import load_file, analyze, to_payload, build_pivot, detect_relationships, merge_sheets
from report_builder import build_report

APP_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(APP_DIR, 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024  # 25 MB
ALLOWED_EXT = {'xlsx', 'xls', 'xlsm', 'csv'}

# small in-memory cache of analysis results (keyed by job+sheet) for pivot building
_RESULT_CACHE = {}
_RESULT_CACHE_ORDER = []
_CACHE_MAX = 4
import threading
_CACHE_LOCK = threading.Lock()


def _cache_get(key):
    with _CACHE_LOCK:
        return _RESULT_CACHE.get(key)


def _cache_put(key, value):
    with _CACHE_LOCK:
        if key not in _RESULT_CACHE:
            _RESULT_CACHE_ORDER.append(key)
        _RESULT_CACHE[key] = value
        while len(_RESULT_CACHE_ORDER) > _CACHE_MAX:
            old = _RESULT_CACHE_ORDER.pop(0)
            _RESULT_CACHE.pop(old, None)


def re_safe(job_id):
    import re
    return bool(re.fullmatch(r'[0-9a-f]{6,20}', job_id))


def _job_source(job_id):
    """Return (job_dir, src_path_or_None, display_filename, meta) for a job, or raises ValueError."""
    if not re_safe(job_id):
        raise ValueError('Unknown job. Please upload the file again.')
    job_dir = os.path.join(UPLOAD_DIR, job_id)
    if not os.path.isdir(job_dir):
        raise ValueError('Unknown job. Please upload the file again.')
    meta = {}
    meta_path = os.path.join(job_dir, 'meta.json')
    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding='utf-8') as fh:
                meta = json.load(fh)
        except Exception:
            meta = {}
    if meta.get('type') == 'merge':
        # merged jobs have no source file - they re-run from the parent job
        return job_dir, None, meta.get('filename', 'merged'), meta
    src = [x for x in os.listdir(job_dir) if x.startswith('source.')]
    if not src:
        raise ValueError('Original file no longer available. Please upload again.')
    disp = meta.get('filename', os.path.basename(src[0]))
    return job_dir, os.path.join(job_dir, src[0]), disp, meta


def _get_result(job_id, sheet_name=None, options=None):
    """Load (or reuse a cached) analysis result for a stored job."""
    job_dir, src_path, disp, meta = _job_source(job_id)
    if meta.get('type') == 'merge':
        key = (job_id, '__merge__')
        cached = _cache_get(key)
        if cached is not None:
            return cached
        result = _join_result(meta.get('parent', job_id), meta.get('params', {}),
                              options or {'standardize_case': True, 'fill_missing': True, 'remove_duplicates': False})
        _cache_put(key, result)
        return result
    key = (job_id, sheet_name or '__auto__')
    cached = _cache_get(key)
    if cached is not None:
        return cached
    sheets, notes, wb_meta = load_file(src_path, disp)
    auto = False
    if sheet_name is None:
        best = None
        for n, d in sheets.items():
            if d.shape[0] > 0 and (best is None or d.shape[0] > sheets[best].shape[0]):
                best = n
        sheet_name = best or (list(sheets)[0] if sheets else None)
        auto = True
    if sheet_name is None or sheet_name not in sheets or sheets[sheet_name].shape[0] == 0:
        raise ValueError('The selected sheet has no data rows.')
    ctx = {'file_name': disp, 'sheet_name': sheet_name, 'auto_selected': auto,
           'sheet_notes': notes, 'hidden_sheets': wb_meta.get('hidden_sheets', []),
           'merged_ranges': wb_meta.get('merged_ranges', {})}
    result = analyze(sheets[sheet_name],
                     options=options or {'standardize_case': True, 'fill_missing': True, 'remove_duplicates': False},
                     context=ctx)
    _cache_put(key, result)
    return result


def _default_options(form):
    def flag(name, default):
        v = form.get(name)
        if v is None:
            return default
        return str(v).lower() in ('1', 'true', 'on', 'yes')
    return {
        'standardize_case': flag('standardize_case', True),
        'fill_missing': flag('fill_missing', True),
        'remove_duplicates': flag('remove_duplicates', False),
    }


def _sheets_info(sheets, wb_meta, chosen=None):
    info = []
    for name, df in sheets.items():
        info.append({
            'name': name, 'rows': int(df.shape[0]), 'cols': int(df.shape[1]),
            'hidden': name in (wb_meta.get('hidden_sheets') or []),
            'merged_ranges': (wb_meta.get('merged_ranges') or {}).get(name, 0),
        })
    info.sort(key=lambda x: -x['rows'])
    return info


def _run_job(job_id, src_path, filename, sheets, notes, wb_meta, sheet_name, options):
    auto = False
    if sheet_name is None:
        best = None
        for n, d in sheets.items():
            if d.shape[0] > 0 and (best is None or d.shape[0] > sheets[best].shape[0]):
                best = n
        sheet_name = best or (list(sheets)[0] if sheets else None)
        auto = True
    if sheet_name is None or sheets[sheet_name].shape[0] == 0:
        raise ValueError('The selected sheet has no data rows.')
    raw_df = sheets[sheet_name]

    ctx = {
        'file_name': filename,
        'sheet_name': sheet_name,
        'auto_selected': auto,
        'sheet_notes': notes,
        'hidden_sheets': wb_meta.get('hidden_sheets', []),
        'merged_ranges': wb_meta.get('merged_ranges', {}),
    }
    result = analyze(raw_df, options=options, context=ctx)
    _cache_put((job_id, sheet_name), result)

    report_name = 'Government_Analytics_Report.xlsx'
    report_path = os.path.join(UPLOAD_DIR, job_id, report_name)
    build_report(result, report_path)

    base = os.path.splitext(os.path.basename(filename))[0][:40].replace(' ', '_') or 'report'
    download_url = f'/download/{job_id}'
    payload = to_payload(result, job_id, filename, sheet_name,
                         _sheets_info(sheets, wb_meta, sheet_name), download_url)
    payload['report_filename'] = f'Government_Analytics_Report_{base}.xlsx'
    with open(os.path.join(UPLOAD_DIR, job_id, 'meta.json'), 'w', encoding='utf-8') as f:
        json.dump({'filename': filename, 'filename_base': base}, f, ensure_ascii=False)
    # persist payload for re-analysis caching
    with open(os.path.join(UPLOAD_DIR, job_id, 'payload.json'), 'w', encoding='utf-8') as f:
        json.dump({'sheet_name': sheet_name, 'payload': payload}, f, ensure_ascii=False)
    return payload


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/analyze', methods=['POST'])
def api_analyze():
    if 'file' not in request.files:
        return jsonify({'ok': False, 'error': 'No file received. Please choose a file first.'}), 400
    f = request.files['file']
    if not f.filename:
        return jsonify({'ok': False, 'error': 'No file selected.'}), 400
    ext = f.filename.lower().rsplit('.', 1)[-1]
    if ext not in ALLOWED_EXT:
        return jsonify({'ok': False, 'error': f'Unsupported file type ".{ext}". Please upload .xlsx, .xls, .xlsm or .csv.'}), 400

    job_id = uuid.uuid4().hex[:12]
    job_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)
    src_path = os.path.join(job_dir, 'source.' + ext)
    f.save(src_path)
    base = os.path.splitext(os.path.basename(f.filename))[0][:40].replace(' ', '_') or 'report'
    with open(os.path.join(job_dir, 'meta.json'), 'w', encoding='utf-8') as fh:
        json.dump({'filename': f.filename, 'filename_base': base}, fh, ensure_ascii=False)

    options = _default_options(request.form)
    try:
        sheets, notes, wb_meta = load_file(src_path, f.filename)
        if not any(df.shape[0] > 0 for df in sheets.values()):
            return jsonify({'ok': False, 'error': 'The file contains no data rows.'}), 400
        sheet_name = request.form.get('sheet') or None
        if sheet_name and sheet_name not in sheets:
            return jsonify({'ok': False, 'error': f'Sheet "{sheet_name}" not found.'}), 400
        payload = _run_job(job_id, src_path, f.filename, sheets, notes, wb_meta, sheet_name, options)
        return jsonify({'ok': True, **payload})
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    except MemoryError:
        return jsonify({'ok': False, 'error': 'File too large to process. Please upload a smaller file (max ~25 MB).'}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({'ok': False, 'error': f'Processing failed: {type(e).__name__}. The file may be corrupted or in an unsupported format.'}), 500


@app.route('/api/reanalyze', methods=['POST'])
def api_reanalyze():
    data = request.get_json(silent=True) or {}
    job_id = str(data.get('job_id', ''))
    sheet_name = data.get('sheet')
    options = {
        'standardize_case': bool(data.get('standardize_case', True)),
        'fill_missing': bool(data.get('fill_missing', True)),
        'remove_duplicates': bool(data.get('remove_duplicates', False)),
    }
    try:
        job_dir, src_path, disp, meta = _job_source(job_id)
        if meta.get('type') == 'merge':
            # merged job: re-run the join with the stored parameters
            params = meta.get('params', {})
            payload = _run_join(meta.get('parent', job_id), params, options)
            return jsonify({'ok': True, **payload})
        sheets, notes, wb_meta = load_file(src_path, disp)
        if sheet_name and sheet_name not in sheets:
            return jsonify({'ok': False, 'error': f'Sheet "{sheet_name}" not found.'}), 400
        payload = _run_job(job_id, src_path, disp, sheets, notes, wb_meta, sheet_name or None, options)
        return jsonify({'ok': True, **payload})
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({'ok': False, 'error': f'Re-analysis failed: {type(e).__name__}.'}), 500


# ------------------------------------------------------------------ #
# PIVOT BUILDER - full pivot options on demand
# ------------------------------------------------------------------ #
@app.route('/api/pivot', methods=['POST'])
def api_pivot():
    data = request.get_json(silent=True) or {}
    job_id = str(data.get('job_id', ''))
    sheet = data.get('sheet') or None
    row = data.get('row')
    col = data.get('col') or None
    measure = data.get('measure') or None
    agg = data.get('agg') or 'sum'
    filter_col = data.get('filter_col') or None
    filter_val = data.get('filter_val') or None
    if not row:
        return jsonify({'ok': False, 'error': 'Choose a "Rows" field first.'}), 400
    try:
        result = _get_result(job_id, sheet)
        pivot = build_pivot(result['cleaned'], row, col, measure, agg, filter_col, filter_val)
        return jsonify({'ok': True, 'pivot': pivot})
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({'ok': False, 'error': f'Pivot failed: {type(e).__name__}.'}), 500


# ------------------------------------------------------------------ #
# DATA MODEL - relationships between sheets
# ------------------------------------------------------------------ #
@app.route('/api/model', methods=['POST'])
def api_model():
    data = request.get_json(silent=True) or {}
    job_id = str(data.get('job_id', ''))
    try:
        job_dir, src_path, disp, meta = _job_source(job_id)
        if meta.get('type') == 'merge':
            return jsonify({'ok': False, 'error': 'This is already a merged dataset - upload the original workbook to explore relationships.'}), 400
        sheets, notes, wb_meta = load_file(src_path, disp)
        rels = detect_relationships(sheets)
        return jsonify({'ok': True, 'relationships': rels,
                        'sheets': _sheets_info(sheets, wb_meta)})
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({'ok': False, 'error': f'Relationship detection failed: {type(e).__name__}.'}), 500


# ------------------------------------------------------------------ #
# DATA MODEL - validated merge of two sheets (join)
# ------------------------------------------------------------------ #
def _join_result(parent_job_id, params, options):
    """Run a validated join and return the analysis result (no files written)."""
    job_dir, src_path, disp, meta = _job_source(parent_job_id)
    sheets, notes, wb_meta = load_file(src_path, disp)
    merged, join_report = merge_sheets(
        sheets, params.get('left_sheet'), params.get('left_key'),
        params.get('right_sheet'), params.get('right_key'), params.get('how', 'left'))
    ctx = {
        'file_name': disp, 'sheet_name': f"Merged: {params.get('left_sheet')} + {params.get('right_sheet')}",
        'auto_selected': False, 'sheet_notes': [], 'hidden_sheets': [], 'merged_ranges': {},
        'join_note': (f"Data model join: \"{params.get('left_sheet')}\".{params.get('left_key')} was joined with "
                      f"\"{params.get('right_sheet')}\".{params.get('right_key')} using a {str(params.get('how','left')).upper()} join. "
                      f"Matched: {join_report['matched_left_rows']} of {join_report['left_rows']} left rows; "
                      f"unmatched left rows kept blank: {join_report['unmatched_left_rows']}."),
    }
    result = analyze(merged, options=options, context=ctx)
    result['merged_data'] = merged
    result['join_report'] = join_report
    assumptions = result['assumptions']
    assumptions.insert(0, ctx['join_note'])
    if join_report.get('note'):
        assumptions.insert(1, join_report['note'])
    result['assumptions'] = assumptions
    return result


def _run_join(parent_job_id, params, options):
    """Run a validated join and produce a full analysis payload + report."""
    result = _join_result(parent_job_id, params, options)
    merged = result['merged_data']
    join_report = result['join_report']
    parent_dir, _, disp, meta = _job_source(parent_job_id)
    new_job = uuid.uuid4().hex[:12]
    new_dir = os.path.join(UPLOAD_DIR, new_job)
    os.makedirs(new_dir, exist_ok=True)
    report_path = os.path.join(new_dir, 'Government_Analytics_Report.xlsx')
    build_report(result, report_path)
    base = (meta.get('filename_base') or 'merged')[:30]
    with open(os.path.join(new_dir, 'meta.json'), 'w', encoding='utf-8') as f:
        json.dump({'filename': disp, 'filename_base': base + '_merged', 'type': 'merge',
                   'parent': parent_job_id, 'params': params}, f, ensure_ascii=False)
    payload = to_payload(result, new_job, disp,
                         f"Merged: {params.get('left_sheet')} + {params.get('right_sheet')}",
                         [{'name': 'Merged_Data', 'rows': int(len(merged)), 'cols': int(len(merged.columns)),
                           'hidden': False, 'merged_ranges': 0}],
                         f'/download/{new_job}')
    payload['report_filename'] = f'Government_Analytics_Report_{base}_merged.xlsx'
    _cache_put((new_job, '__merge__'), result)
    return payload


@app.route('/api/join', methods=['POST'])
def api_join():
    data = request.get_json(silent=True) or {}
    job_id = str(data.get('job_id', ''))
    params = {
        'left_sheet': data.get('left_sheet'), 'left_key': data.get('left_key'),
        'right_sheet': data.get('right_sheet'), 'right_key': data.get('right_key'),
        'how': data.get('how') or 'left',
    }
    options = {
        'standardize_case': bool(data.get('standardize_case', True)),
        'fill_missing': bool(data.get('fill_missing', True)),
        'remove_duplicates': bool(data.get('remove_duplicates', False)),
    }
    if not all([params['left_sheet'], params['left_key'], params['right_sheet'], params['right_key']]):
        return jsonify({'ok': False, 'error': 'Both sheets and both key columns are required for a merge.'}), 400
    try:
        payload = _run_join(job_id, params, options)
        return jsonify({'ok': True, **payload})
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    except MemoryError:
        return jsonify({'ok': False, 'error': 'The merged result is too large to process.'}), 400
    except Exception as e:
        traceback.print_exc()
        return jsonify({'ok': False, 'error': f'Merge failed: {type(e).__name__}.'}), 500


@app.route('/download/<job_id>')
def download(job_id):
    if not re_safe(job_id):
        abort(404)
    path = os.path.join(UPLOAD_DIR, job_id, 'Government_Analytics_Report.xlsx')
    if not os.path.exists(path):
        abort(404)
    fname = 'Government_Analytics_Report.xlsx'
    meta_path = os.path.join(UPLOAD_DIR, job_id, 'meta.json')
    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding='utf-8') as fh:
                base = json.load(fh).get('filename_base', 'report')
                fname = f'Government_Analytics_Report_{base}.xlsx'
        except Exception:
            pass
    return send_file(path, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/health')
def health():
    return jsonify({'ok': True, 'time': dt.datetime.now().isoformat()})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=7860, threaded=True)
