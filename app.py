import os
import json
import tempfile
import traceback
from functools import wraps
from flask import (Flask, render_template, request, jsonify, send_from_directory,
                   send_file, session, redirect, url_for)
from cutoff_engine import scan_colleges_from_excel, parse_college_data, generate_word_document, fast_precache_all

app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = os.environ.get('SECRET_KEY', 'ame-cutoff-book-maker-2027-secret')

# ---------- Login credentials ----------
LOGIN_USER = 'ame2027'
LOGIN_PASS = 'ame@2027'

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
CET_DIR = os.path.join(DATA_DIR, 'cet')
AI_DIR = os.path.join(DATA_DIR, 'ai')

DEFAULT_EXCEL_PATH = r"E:\MHT_CET_Cutoff_PDF_to_Excel_Master_Extractor\CET DATA\ALL_COLLEGES_CONSOLIDATED_20260916_151333.xlsx"
DEFAULT_AI_EXCEL_PATH = r"E:\OFFICE PERSONAL DATA\ENGG 2027\MH_AI_OPEN_CLOSE_2026_SINGLE_ROUND.xlsx"
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')
CACHE_FILE = os.path.join(BASE_DIR, 'cache.json')

EXCEL_EXTS = ('.xlsx', '.xlsm', '.xls')


def find_data_file(folder):
    """Newest Excel dropped inside a bundled folder (data/cet or data/ai)."""
    try:
        names = [n for n in os.listdir(folder)
                 if n.lower().endswith(EXCEL_EXTS) and not n.startswith('~$')]
    except OSError:
        return None
    if not names:
        return None
    full = [os.path.join(folder, n) for n in names]
    try:
        full.sort(key=lambda p: (os.path.getmtime(p), os.path.basename(p).lower()), reverse=True)
    except OSError:
        full.sort(reverse=True)
    return full[0]


def resolve_data_path(path, kind='cet'):
    """Use the configured path when it exists on this machine; otherwise pick the
    bundled Excel from ./data/cet or ./data/ai (Vercel has no access to the PC)."""
    if path and os.path.exists(path) and os.path.isfile(path):
        return path
    bundled = find_data_file(CET_DIR if kind == 'cet' else AI_DIR)
    if bundled:
        return bundled
    return path


def pick_output_folder():
    candidates = [
        os.path.join(BASE_DIR, 'College_Data'),
        os.path.join(tempfile.gettempdir(), 'College_Data'),
    ]
    for folder in candidates:
        try:
            os.makedirs(folder, exist_ok=True)
            probe = os.path.join(folder, '.write_test')
            with open(probe, 'w') as f:
                f.write('ok')
            os.remove(probe)
            return folder
        except OSError:
            continue
    return tempfile.gettempdir()


OUTPUT_FOLDER = pick_output_folder()
LOGO_PATH = os.path.join(BASE_DIR, 'static', 'logo.jpg')

# In-memory cache for ultra-fast instant lookups
MEMORY_CACHE = {}

def default_config():
    return {'saved_path': DEFAULT_EXCEL_PATH, 'ai_excel_path': DEFAULT_AI_EXCEL_PATH}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            cfg.setdefault('saved_path', DEFAULT_EXCEL_PATH)
            cfg.setdefault('ai_excel_path', DEFAULT_AI_EXCEL_PATH)
            cfg['saved_path'] = resolve_data_path(cfg.get('saved_path'), 'cet')
            cfg['ai_excel_path'] = resolve_data_path(cfg.get('ai_excel_path'), 'ai')
            return cfg
        except Exception:
            pass
    cfg = default_config()
    cfg['saved_path'] = resolve_data_path(cfg['saved_path'], 'cet')
    cfg['ai_excel_path'] = resolve_data_path(cfg['ai_excel_path'], 'ai')
    return cfg

def clean_path(value):
    return str(value or '').strip().strip('"').strip("'")

def get_ai_path(data=None):
    p = clean_path((data or {}).get('ai_excel_path'))
    if p: return p
    return load_config().get('ai_excel_path', DEFAULT_AI_EXCEL_PATH)

def make_cache_key(excel_path, ai_path, query, mode):
    # basename keys => the SAME cache works locally (E:\ drive) and on Vercel (./data)
    return f"{os.path.basename(str(excel_path))}||{os.path.basename(str(ai_path))}_{query}_{mode}"

def _file_sig(path):
    try:
        st = os.stat(path)
        return str(st.st_mtime), int(st.st_size)
    except OSError:
        return None

def sig_dict(excel_path, ai_path):
    """mtime + size stamps stored with every cache entry (size also works on Vercel)."""
    m, sz = _file_sig(excel_path) or ("", 0)
    a = _file_sig(ai_path) if ai_path and os.path.exists(ai_path) else None
    am, asz = a or ("", 0)
    return {'mtime': m, 'size': sz, 'ai_mtime': am, 'ai_size': asz}

def cache_entry_is_valid(entry, excel_path, ai_path):
    sig = _file_sig(excel_path)
    if not sig:
        return False
    mtime, size = sig
    # mtime OR size may match (Vercel extracts the zip with fresh mtimes)
    if entry.get('mtime') != mtime and entry.get('size') != size:
        return False
    if ai_path and os.path.exists(ai_path):
        a_sig = _file_sig(ai_path)
        if a_sig and entry.get('ai_mtime') != a_sig[0] and entry.get('ai_size') != a_sig[1]:
            return False
    elif entry.get('ai_mtime'):
        return False
    return True

def save_config(config_data):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2)
    except Exception as e:
        print(f"Error saving config: {e}")

def load_cache():
    global MEMORY_CACHE
    if MEMORY_CACHE:
        return MEMORY_CACHE
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                MEMORY_CACHE = json.load(f)
                return MEMORY_CACHE
        except Exception as e:
            print(f"Error loading cache: {e}")
    return {}

def save_cache(cache_data):
    global MEMORY_CACHE
    MEMORY_CACHE = cache_data
    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f, indent=2)
    except Exception as e:
        print(f"Error saving cache: {e}")

# Pre-load cache into memory at startup
load_cache()

# ============================================================
# LOGIN
# ============================================================
def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get('logged_in'):
            if request.path.startswith('/api/') or request.path.startswith('/download/'):
                return jsonify({'status': 'auth', 'message': 'Login required'}), 401
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrapper

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = ''
    if request.method == 'POST':
        user = (request.form.get('username') or '').strip()
        pwd = request.form.get('password') or ''
        if user == LOGIN_USER and pwd == LOGIN_PASS:
            session['logged_in'] = True
            session['user'] = user
            nxt = request.args.get('next') or request.args.get('next_url') or '/'
            if not str(nxt).startswith('/'):
                nxt = '/'
            return redirect(nxt)
        error = 'Invalid username or password. Please try again.'
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/whoami')
def whoami():
    return jsonify({'logged_in': bool(session.get('logged_in')), 'user': session.get('user', '')})

@app.route('/')
@login_required
def index():
    cfg = load_config()
    current_path = cfg.get('saved_path', DEFAULT_EXCEL_PATH)
    ai_path = cfg.get('ai_excel_path', DEFAULT_AI_EXCEL_PATH)
    return render_template('index.html', default_path=current_path, ai_path=ai_path,
                           logged_user=session.get('user', ''))

@app.route('/api/config', methods=['GET', 'POST'])
@login_required
def api_config():
    if request.method == 'POST':
        data = request.json or {}
        cfg = load_config()
        response = {'status': 'success'}

        if 'excel_path' in data:
            new_path = clean_path(data.get('excel_path'))
            cfg['saved_path'] = new_path or DEFAULT_EXCEL_PATH
            response['saved_path'] = cfg['saved_path']
            response['message'] = 'CET path saved successfully!' if new_path else 'CET path reset to default.'

        if 'ai_excel_path' in data:
            new_ai = clean_path(data.get('ai_excel_path'))
            cfg['ai_excel_path'] = new_ai or DEFAULT_AI_EXCEL_PATH
            response['ai_excel_path'] = cfg['ai_excel_path']
            response['ai_message'] = 'AI (JEE) path saved successfully!' if new_ai else 'AI (JEE) path reset to default.'

        if not ('excel_path' in data or 'ai_excel_path' in data):
            cfg = default_config()
            response['saved_path'] = cfg['saved_path']
            response['ai_excel_path'] = cfg['ai_excel_path']
            response['message'] = 'All paths reset to default.'

        save_config(cfg)
        response.setdefault('saved_path', cfg.get('saved_path', DEFAULT_EXCEL_PATH))
        response.setdefault('ai_excel_path', cfg.get('ai_excel_path', DEFAULT_AI_EXCEL_PATH))
        return jsonify(response)
    else:
        cfg = load_config()
        return jsonify({
            'status': 'success',
            'saved_path': cfg.get('saved_path', DEFAULT_EXCEL_PATH),
            'ai_excel_path': cfg.get('ai_excel_path', DEFAULT_AI_EXCEL_PATH)
        })

@app.route('/api/scan-excel', methods=['POST'])
@login_required
def api_scan_excel():
    data = request.json or {}
    excel_path = clean_path(data.get('excel_path')) or load_config().get('saved_path', DEFAULT_EXCEL_PATH)
    ai_path = get_ai_path(data)

    if not os.path.exists(excel_path):
        return jsonify({'status': 'error', 'message': f'File not found: {excel_path}'}), 400

    try:
        colleges, total_rows = scan_colleges_from_excel(excel_path, ai_path)
        cache = load_cache()

        cached_count = 0
        ai_count = 0
        for c in colleges:
            if c.get('has_ai'): ai_count += 1
            query = c.get('code') or c.get('name')
            ck = make_cache_key(excel_path, ai_path, query, 'ALL')
            if ck in cache and cache_entry_is_valid(cache[ck], excel_path, ai_path):
                cached_count += 1

        return jsonify({
            'status': 'success',
            'excel_path': excel_path,
            'ai_excel_path': ai_path,
            'ai_file_ok': bool(ai_path and os.path.exists(ai_path)),
            'ai_colleges': ai_count,
            'total_rows': total_rows,
            'total_colleges': len(colleges),
            'cached_count': cached_count,
            'colleges': colleges
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/api/get-college-data', methods=['POST'])
@login_required
def api_get_college_data():
    """Ultra-fast instant retrieval for selected college."""
    data = request.json or {}
    college_query = str(data.get('college_query', '')).strip()
    excel_path = clean_path(data.get('excel_path')) or load_config().get('saved_path', DEFAULT_EXCEL_PATH)
    ai_path = get_ai_path(data)

    cache = load_cache()
    prefix = f"{os.path.basename(str(excel_path))}||{os.path.basename(str(ai_path))}_"

    # 1. Direct key search in memory cache
    for mode in ['ALL', 'HU', 'STATE', 'OHU']:
        direct_key = make_cache_key(excel_path, ai_path, college_query, mode)
        entry = cache.get(direct_key)
        if entry and cache_entry_is_valid(entry, excel_path, ai_path):
            return jsonify({
                'status': 'success',
                'data': entry.get('data'),
                'is_cached': True
            })

    # 2. Substring code search (only inside the current CET + AI key space)
    for k, v in cache.items():
        if not k.startswith(prefix):
            continue
        if f"_{college_query}_" in k or k.endswith(f"_{college_query}"):
            if cache_entry_is_valid(v, excel_path, ai_path):
                return jsonify({
                    'status': 'success',
                    'data': v.get('data'),
                    'is_cached': True
                })

    # 3. Fallback: Parse on demand
    try:
        parsed_data = parse_college_data(excel_path, college_query, mode="ALL", ai_path=ai_path)
        cache_key = make_cache_key(excel_path, ai_path, college_query, 'ALL')
        cache[cache_key] = {
            **sig_dict(excel_path, ai_path),
            'filename': f"{parsed_data.get('college_code') or 'college'}.docx",
            'data': parsed_data
        }
        save_cache(cache)
        return jsonify({
            'status': 'success',
            'data': parsed_data,
            'is_cached': False
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/precache-all', methods=['POST'])
@login_required
def api_precache_all():
    data = request.json or {}
    excel_path = clean_path(data.get('excel_path')) or load_config().get('saved_path', DEFAULT_EXCEL_PATH)
    ai_path = get_ai_path(data)

    if not os.path.exists(excel_path):
        return jsonify({'status': 'error', 'message': f'Excel file not found at: {excel_path}'}), 400
    if ai_path and not os.path.exists(ai_path):
        return jsonify({'status': 'error', 'message': f'AI (JEE) Excel file not found at: {ai_path}'}), 400

    try:
        colleges_dict = fast_precache_all(excel_path, ai_path)
        cache = load_cache()
        stamps = sig_dict(excel_path, ai_path)

        # Keep only entries of the current CET file (drops old key formats / old files)
        base = os.path.basename(str(excel_path))
        stale_keys = [k for k in cache if not k.startswith(f"{base}||")]
        for k in stale_keys:
            del cache[k]

        ai_branch_total = 0
        for key, parsed_data in colleges_dict.items():
            ai_branch_total += parsed_data.get('ai_branches', 0)
            for mode in ['ALL', 'HU', 'OHU', 'STATE', 'HU_OHU']:
                cache_key = make_cache_key(excel_path, ai_path, key, mode)
                cache[cache_key] = {
                    **stamps,
                    'filename': f"{parsed_data.get('college_code') or 'college'}.docx",
                    'data': parsed_data
                }

        save_cache(cache)
        return jsonify({
            'status': 'success',
            'message': f'Regenerated all tables: {len(colleges_dict)} colleges, {ai_branch_total} AI (JEE) branch cut-offs merged.',
            'total_cached': len(colleges_dict),
            'ai_branches': ai_branch_total
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/api/generate-table', methods=['POST'])
@login_required
def api_generate_table():
    data = request.json or {}
    excel_path = clean_path(data.get('excel_path')) or load_config().get('saved_path', DEFAULT_EXCEL_PATH)
    college_query = data.get('college_query', '').strip()
    mode = data.get('mode', 'ALL').strip().upper()
    force_refresh = data.get('force_refresh', False)
    ai_path = get_ai_path(data)

    if not os.path.exists(excel_path):
        return jsonify({'status': 'error', 'message': f'Excel file not found at: {excel_path}'}), 400

    try:
        cache = load_cache()
        cache_key = make_cache_key(excel_path, ai_path, college_query, mode)

        # Fast cached response if not forcing refresh
        if not force_refresh and cache_key in cache:
            cached_entry = cache[cache_key]
            if cache_entry_is_valid(cached_entry, excel_path, ai_path):
                parsed_data = cached_entry.get('data')
                filename = cached_entry.get('filename', f"{college_query}.docx")
                return jsonify({
                    'status': 'success',
                    'data': parsed_data,
                    'download_url': f'/api/download-docx?code={college_query}&mode={mode}',
                    'filename': filename,
                    'mode': mode,
                    'is_cached': True
                })

        # Fresh parse if force_refresh is True or not in cache
        parsed_data = parse_college_data(excel_path, college_query, mode, ai_path=ai_path)
        filename = f"{parsed_data.get('college_code') or 'college'}.docx"

        cache[cache_key] = {
            **sig_dict(excel_path, ai_path),
            'filename': filename,
            'data': parsed_data
        }
        save_cache(cache)

        return jsonify({
            'status': 'success',
            'data': parsed_data,
            'download_url': f'/api/download-docx?code={college_query}&mode={mode}',
            'filename': filename,
            'mode': mode,
            'is_cached': False
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/api/download-docx')
@login_required
def api_download_docx():
    """Generates and serves Word document on-demand when user clicks download."""
    code = request.args.get('code', '').strip()
    mode = request.args.get('mode', 'ALL').strip().upper()
    cfg = load_config()
    excel_path = cfg.get('saved_path', DEFAULT_EXCEL_PATH)
    ai_path = cfg.get('ai_excel_path', DEFAULT_AI_EXCEL_PATH)

    cache = load_cache()
    prefix = f"{os.path.basename(str(excel_path))}||{os.path.basename(str(ai_path))}_"

    parsed_data = None
    direct_key = make_cache_key(excel_path, ai_path, code, mode)
    if direct_key in cache:
        parsed_data = cache[direct_key].get('data')
    if not parsed_data:
        for k, v in cache.items():
            if k.startswith(prefix) and (f"_{code}_" in k or k.endswith(f"_{code}")):
                parsed_data = v.get('data')
                break

    if not parsed_data:
        parsed_data = parse_college_data(excel_path, code, mode, ai_path=ai_path)

    file_path, filename = generate_word_document(parsed_data, mode=mode, logo_path=LOGO_PATH, output_folder=OUTPUT_FOLDER)
    return send_file(file_path, as_attachment=True, download_name=filename)

@app.route('/download/<filename>')
@login_required
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    print("=" * 60)
    print("  Starting AME CUTOFF BOOK MAKER Web Application")
    print("  Running on: http://localhost:5000")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5000, debug=True)
