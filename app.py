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

from analyzer import load_file, analyze, to_payload
from report_builder import build_report

APP_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(APP_DIR, 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024  # 25 MB
ALLOWED_EXT = {'xlsx', 'xls', 'xlsm', 'csv'}


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
    job_dir = os.path.join(UPLOAD_DIR, job_id)
    if not os.path.isdir(job_dir) or not re_safe(job_id):
        return jsonify({'ok': False, 'error': 'Unknown job. Please upload the file again.'}), 404
    src = [x for x in os.listdir(job_dir) if x.startswith('source.')]
    if not src:
        return jsonify({'ok': False, 'error': 'Original file no longer available. Please upload again.'}), 404
    src_path = os.path.join(job_dir, src[0])
    try:
        payload_cache = os.path.join(job_dir, 'payload.json')
        filename = os.path.basename(src_path)
        if payload_cache.endswith('.xlsx') and False:
            pass
        # get original display filename
        disp = filename
        meta_path = os.path.join(job_dir, 'meta.json')
        if os.path.exists(meta_path):
            with open(meta_path, encoding='utf-8') as fh:
                disp = json.load(fh).get('filename', filename)
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


def re_safe(job_id):
    import re
    return bool(re.fullmatch(r'[0-9a-f]{6,20}', job_id))


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
