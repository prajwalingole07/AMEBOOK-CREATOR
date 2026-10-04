import os
import re
import shutil
import pandas as pd
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

SKIP_CATS = ["DEFOPENS", "DEFOBCS", "DEFROBCS", "DEFLOPEN", "DEFLOBC"]
TABLE_COLS = ["Branch (Intake)", "Gender", "OPEN", "SC", "ST", "VJ-DT",
              "NT-1 (B)", "NT-2 (C)", "NT-3 (D)", "OBC", "SEBC", "EWS", "TFWS", "AI (JEE)"]
CAT_ORDER = ["OPEN", "SC", "ST", "VJ-DT", "NT-1 (B)", "NT-2 (C)", "NT-3 (D)", "OBC", "SEBC", "EWS", "TFWS", "AI (JEE)"]
REGULAR_CATS = ["OPEN", "SC", "ST", "VJ-DT", "NT-1 (B)", "NT-2 (C)", "NT-3 (D)", "OBC", "SEBC"]
AI_COL = "AI (JEE)"
# Columns that hold ONE value per branch (not per gender) -> merged + centred
SHARED_COLS = ["EWS", "TFWS", "AI (JEE)"]
AI_EXCEL_DEFAULT = r"E:\OFFICE PERSONAL DATA\ENGG 2027\MH_AI_OPEN_CLOSE_2026_SINGLE_ROUND.xlsx"

def find_code_col(df):
    for c in df.columns:
        su = str(c).upper()
        if 'CODE' in su and ('INST' in su or 'COLLEGE' in su):
            return c
    for c in df.columns:
        su = str(c).upper()
        if 'CODE' in su and 'BRANCH' not in su and 'COURSE' not in su:
            return c
    return None

def find_inst_name_col(df):
    for c in df.columns:
        su = str(c).upper()
        if ('INSTITUTE' in su or 'COLLEGE' in su) and 'CODE' not in su:
            return c
    for c in df.columns:
        su = str(c).upper()
        if 'NAME' in su and 'BRANCH' not in su and 'COURSE' not in su and 'CITY' not in su and 'CODE' not in su:
            return c
    return None

def find_col(df, keywords):
    for col in df.columns:
        col_str = str(col).upper()
        if all(kw.upper() in col_str for kw in keywords):
            return col
    return None

def parse_int(val):
    if val is None or pd.isna(val): return None
    s = str(val).strip()
    if not s: return None
    match = re.search(r'-?\d+', s)
    if match:
        try: return int(match.group())
        except: return None
    return None

def get_cutoff_cell_from_row(row):
    """
    Picks the best round (highest percentile) across R1..R4 and returns
    "rank" or "rank\n(percentile%)" (same two-line layout as the AI (JEE) column).
    """
    pairs = []
    for idx in range(1, 5):
        rk_val = parse_int(row.get(f'R{idx} RANK'))
        pc_val = row.get(f'R{idx} PERCENTILE')
        if rk_val is not None:
            try:
                pc_num = float(pc_val) if pd.notna(pc_val) else -1.0
            except (TypeError, ValueError):
                pc_num = -1.0
            pairs.append((rk_val, pc_num))

    if not pairs:
        return "-----"

    pairs.sort(key=lambda x: (x[1], -x[0]), reverse=True)
    rank, pct = pairs[0]
    if pct and pct > 0:
        return f"{rank}\n({pct:.2f}%)"
    return str(rank)

# ============================================================
# AI (JEE) EXCEL SUPPORT  -  single round rank / percentile
# ============================================================
_AI_FILE_CACHE = {"key": None, "rows": {}, "colleges": {}}

def norm_ai_code(v):
    if v is None or (isinstance(v, float) and pd.isna(v)): return ""
    s = str(v).strip()
    if s.endswith('.0'): s = s[:-2]
    m = re.match(r'^0*(\d+)', s)
    return m.group(1) if m else s

def branch_key(s):
    if s is None or (isinstance(s, float) and pd.isna(s)): return ""
    s = str(s).lower().replace('&', ' and ')
    s = s.replace('-', '').replace(',', '')
    s = re.sub(r'[\(\)\[\]]', ' ', s)
    s = re.sub(r'[^a-z0-9 ]+', ' ', s)
    s = re.sub(r'\bengg\b', 'engineering', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def compact_key(s):
    return branch_key(s).replace(' ', '')

def names_match(cet_name, ai_name):
    ck = compact_key(cet_name)
    ak = compact_key(ai_name)
    if not ck or not ak: return False
    if ck == ak: return True
    raw = str(cet_name or '').strip()
    if raw.endswith('...'):
        prefix = compact_key(raw.rstrip('.'))
        if prefix and ak.startswith(prefix): return True
    raw2 = str(ai_name or '').strip()
    if raw2.endswith('...'):
        prefix2 = compact_key(raw2.rstrip('.'))
        if prefix2 and ck.startswith(prefix2): return True
    return False

def names_similar(cet_name, ai_name):
    """Same as names_match but also accepts a clear prefix form, e.g.
    'Electronics and Communication' vs 'Electronics and Communication Engineering'."""
    if names_match(cet_name, ai_name):
        return True
    ck = compact_key(cet_name)
    ak = compact_key(ai_name)
    if len(ck) < 14 or len(ak) < 14:
        return False
    return ck.startswith(ak) or ak.startswith(ck)

def read_ai_excel(ai_path):
    """Reads the AI (JEE) single-round excel. Cached by path + mtime."""
    if not ai_path or not os.path.exists(ai_path):
        return {}, {}
    try:
        mtime = str(os.path.getmtime(ai_path))
    except OSError:
        return {}, {}
    if _AI_FILE_CACHE["key"] == (ai_path, mtime):
        return _AI_FILE_CACHE["rows"], _AI_FILE_CACHE["colleges"]

    df = pd.read_excel(ai_path)
    code_col = find_col(df, ["INST", "CODE"]) or find_code_col(df)
    name_col = find_col(df, ["INSTITUTE", "NAME"]) or find_inst_name_col(df)
    city_col = find_col(df, ["CITY"])
    course_col = find_col(df, ["COURSE", "NAME"]) or find_col(df, ["BRANCH"])
    choice_col = find_col(df, ["CHOICE", "CODE"])
    rank_col = find_col(df, ["AI", "MERIT"]) or find_col(df, ["RANK"])
    pct_col = find_col(df, ["PERCENTILE"])
    status_col = find_col(df, ["STATUS"])
    uni_col = find_col(df, ["UNIVERSITY"])

    rows_by_college = {}
    colleges = {}
    for _, row in df.iterrows():
        code = norm_ai_code(row.get(code_col)) if code_col else ""
        if not code: continue
        course = str(row.get(course_col, '')).strip() if course_col else ""
        if not course or course.lower() == 'nan': continue

        rank = parse_int(row.get(rank_col)) if rank_col else None
        if rank is None: continue
        pct = None
        if pct_col and pd.notna(row.get(pct_col)):
            try: pct = float(row.get(pct_col))
            except (TypeError, ValueError): pct = None

        rows_by_college.setdefault(code, []).append({
            "course": course,
            "code": norm_ai_code(row.get(choice_col)) if choice_col else "",
            "rank": rank,
            "pct": pct,
        })

        if code not in colleges:
            def clean(col):
                if not col: return ""
                v = row.get(col)
                if pd.isna(v): return ""
                return str(v).strip()
            colleges[code] = {
                "status": clean(status_col),
                "university": clean(uni_col),
                "city": clean(city_col),
                "name": clean(name_col),
            }

    _AI_FILE_CACHE["key"] = (ai_path, mtime)
    _AI_FILE_CACHE["rows"] = rows_by_college
    _AI_FILE_CACHE["colleges"] = colleges
    return rows_by_college, colleges

def build_ai_index(cet_df, ai_path):
    """
    Matches every AI (JEE) row to the exact CET branch name of the same college.
    Pass 1: branch name match (exact or truncated).  Pass 2: choice code match,
    accepted only when the labels agree (exact or a clear prefix form).
    Returns: ({college_code: {cet_branch_name: "rank\n(percentile%)"}}, {college_code: meta})
    """
    rows_by_college, colleges_meta = read_ai_excel(ai_path)
    if not rows_by_college:
        return {}, {}

    code_col = find_code_col(cet_df)
    branch_col = find_col(cet_df, ["BRANCH", "NAME"]) or find_col(cet_df, ["COURSE", "NAME"]) or find_col(cet_df, ["BRANCH"])
    branch_code_col = find_col(cet_df, ["BRANCH", "CODE"])
    if not code_col or not branch_col:
        return {}, {}

    cet_branches = {}
    keep = [c for c in [code_col, branch_col, branch_code_col] if c]
    for _, r in cet_df[keep].drop_duplicates().iterrows():
        c_code = norm_ai_code(r.get(code_col))
        b_name = str(r.get(branch_col, '')).strip() if pd.notna(r.get(branch_col)) else ""
        if not c_code or not b_name or b_name.lower() == 'nan': continue
        entry = cet_branches.setdefault(c_code, {"names": [], "by_key": {}, "by_code": {}, "trunc": []})
        if b_name not in entry["names"]:
            entry["names"].append(b_name)
            entry["by_key"].setdefault(compact_key(b_name), b_name)
            if b_name.endswith('...'):
                entry["trunc"].append(b_name)
        if branch_code_col is not None and pd.notna(r.get(branch_code_col)):
            bc = norm_ai_code(r.get(branch_code_col))
            if bc:
                entry["by_code"].setdefault(bc, b_name)

    matches = {}
    for c_code, rows in rows_by_college.items():
        info = cet_branches.get(c_code)
        if not info: continue
        picked = {}
        claimed = set()
        pending = []
        for row in rows:
            target = info["by_key"].get(compact_key(row["course"]))
            if not target:
                ak = compact_key(row["course"])
                for t in info["trunc"]:
                    if ak and ak.startswith(compact_key(t.rstrip('.'))):
                        target = t
                        break
            if target:
                prev = picked.get(target)
                pct = row["pct"] if row["pct"] is not None else -1.0
                if prev is None or pct > prev["pct"]:
                    picked[target] = {"rank": row["rank"], "pct": pct}
                claimed.add(target)
            else:
                pending.append(row)
        for row in pending:
            cands = []
            if row["code"] and row["code"] in info["by_code"]:
                cands.append(info["by_code"][row["code"]])
            if not cands and row["code"]:
                suf = row["code"][-5:]
                for bc, bn in info["by_code"].items():
                    if bc.endswith(suf) and bn not in cands:
                        cands.append(bn)
            target = None
            exact = info["by_code"].get(row["code"]) if row["code"] else None
            for c in cands:
                if not names_similar(c, row["course"]):
                    continue
                if c in claimed and c != exact:
                    continue
                target = c
                break
            if not target:
                continue
            prev = picked.get(target)
            pct = row["pct"] if row["pct"] is not None else -1.0
            if prev is None or pct > prev["pct"]:
                picked[target] = {"rank": row["rank"], "pct": pct}
            claimed.add(target)
        if picked:
            matches[c_code] = {
                b: f"{v['rank']}\n({v['pct']:.2f}%)" if v["pct"] is not None else str(v["rank"])
                for b, v in picked.items()
            }
    return matches, colleges_meta

def apply_ai_to_branches(branches, ai_matches):
    """Writes the AI (JEE) rank+percentile into every quota table (both genders)."""
    if not ai_matches: return 0
    added = 0
    for b_name, b_data in branches.items():
        val = ai_matches.get(b_name)
        if not val: continue
        added += 1
        for key in ("HU_M", "HU_F", "OHU_M", "OHU_F", "STATE_M", "STATE_F"):
            b_data.setdefault(key, {})[AI_COL] = val
    return added

def check_quota_has_regular_data(branches, mode):
    regular_cols = ["OPEN", "SC", "ST", "VJ-DT", "NT-1 (B)", "NT-2 (C)", "NT-3 (D)", "OBC", "SEBC"]
    m_key = f"{mode}_M"
    f_key = f"{mode}_F"
    for b_name, b_data in branches.items():
        m_dict = b_data.get(m_key, {})
        f_dict = b_data.get(f_key, {})
        for col in regular_cols:
            v_m = m_dict.get(col)
            if v_m and str(v_m).strip() not in ("-----", "", "None", "nan"):
                return True
            v_f = f_dict.get(col)
            if v_f and str(v_f).strip() not in ("-----", "", "None", "nan"):
                return True
    return False

CATEGORY_MAPPING = {
    "OPEN": "OPEN",
    "SC": "SC",
    "ST": "ST",
    "VJ": "VJ-DT",
    "VJDT": "VJ-DT",
    "VJ-DT": "VJ-DT",
    "NT1": "NT-1 (B)",
    "NT-1": "NT-1 (B)",
    "NT2": "NT-2 (C)",
    "NT-2": "NT-2 (C)",
    "NT3": "NT-3 (D)",
    "NT-3": "NT-3 (D)",
    "OBC": "OBC",
    "SEBC": "SEBC",
}

def classify_and_route(cat):
    c = str(cat).strip().upper()
    if not c: return None
    
    # 1. STRICT FILTER: Discard ANY Defence, PWD, Orphan, Minority categories completely
    if c.startswith(("DEF", "PWD", "ORPHAN", "MI")): return None
    if "DEF" in c or "PWD" in c or "ORPHAN" in c or "MINORITY" in c: return None
    
    # 2. Special state-level quotas
    if "EWS" in c: return ("Male", "EWS", "SPECIAL")
    if "TFWS" in c: return ("Male", "TFWS", "SPECIAL")
    if "AI" in c or "JEE" in c: return ("Male", "AI (JEE)", "SPECIAL")
    
    # 3. Regular MHT-CET categories must start with G (General/Male) or L (Ladies/Female)
    if not (c.startswith("G") or c.startswith("L")): return None
    
    # 4. Must end with quota suffix H (Home University), O (Other than HU), or S (State Level)
    if not (c.endswith("H") or c.endswith("O") or c.endswith("S")): return None
    
    gender = "Female" if c.startswith("L") else "Male"
    dest = "HU" if c.endswith("H") else ("OHU" if c.endswith("O") else "STATE")
    
    mid = c[1:-1]
    base = CATEGORY_MAPPING.get(mid)
    if not base: return None
    
    return (gender, base, dest)

def fetch_college_web_info(college_name):
    """Fallback web scraper to enrich missing college information."""
    info = {
        "nba_grading": "Available in MHT-CET Record",
        "total_seats": "Available",
        "website": "",
        "summary": "Data extracted from MHT-CET Cutoff Database."
    }
    return info

def scan_colleges_from_excel(excel_path, ai_path=None):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"File not found at path: {excel_path}")
        
    df = pd.read_excel(excel_path)
    ai_matches, ai_meta = build_ai_index(df, ai_path)
    
    code_col = find_code_col(df)
    inst_col = find_inst_name_col(df)
    city_col = find_col(df, ["CITY"]) or find_col(df, ["LOCATION"])
    branch_col = find_col(df, ["BRANCH", "NAME"]) or find_col(df, ["COURSE", "NAME"]) or find_col(df, ["BRANCH"])
    cat_col = find_col(df, ["CATEGORY"])
    
    colleges_dict = {}
    
    for _, row in df.iterrows():
        c_code = str(row.get(code_col, '')).strip() if code_col and pd.notna(row.get(code_col)) else ""
        c_name = str(row.get(inst_col, '')).strip() if inst_col and pd.notna(row.get(inst_col)) else ""
        c_city = str(row.get(city_col, '')).strip() if city_col and pd.notna(row.get(city_col)) else ""
        b_name = str(row.get(branch_col, '')).strip() if branch_col and pd.notna(row.get(branch_col)) else ""
        cat_val = str(row.get(cat_col, '')).strip().upper() if cat_col and pd.notna(row.get(cat_col)) else ""
        
        if c_code.endswith('.0'): c_code = c_code[:-2]
        if not c_code and not c_name: continue
        
        key = c_code if c_code else c_name
        if key not in colleges_dict:
            colleges_dict[key] = {
                "code": c_code,
                "name": c_name if c_name else f"Institute {c_code}",
                "city": c_city,
                "branches": set(),
                "categories": set(),
                "has_hu": False,
                "has_ohu": False,
                "has_state": False,
                "has_ews": False,
                "has_tfws": False,
                "has_ai": norm_ai_code(c_code) in ai_matches or key in ai_matches,
                "display": f"{c_code} - {c_name}" if c_code and c_name and c_code != c_name else (c_name or c_code)
            }
            
        c_obj = colleges_dict[key]
        if b_name and b_name.lower() != "nan":
            c_obj["branches"].add(b_name)
        if cat_val:
            c_obj["categories"].add(cat_val)
            route = classify_and_route(cat_val)
            if route:
                _, _, dest = route
                if dest == "HU": c_obj["has_hu"] = True
                elif dest == "OHU": c_obj["has_ohu"] = True
                elif dest == "STATE": c_obj["has_state"] = True
                elif dest == "SPECIAL":
                    if "EWS" in cat_val: c_obj["has_ews"] = True
                    if "TFWS" in cat_val: c_obj["has_tfws"] = True
            
    colleges_list = []
    for k, v in colleges_dict.items():
        colleges_list.append({
            "code": v["code"],
            "name": v["name"],
            "city": v["city"],
            "branches_count": len(v["branches"]),
            "display": v["display"],
            "has_hu": v["has_hu"],
            "has_ohu": v["has_ohu"],
            "has_state": v["has_state"],
            "has_ews": v["has_ews"],
            "has_tfws": v["has_tfws"],
            "has_ai": v["has_ai"]
        })
        
    colleges_list.sort(key=lambda x: (x["code"] or "99999", x["name"]))
    return colleges_list, len(df)

def fast_precache_all(excel_path, ai_path=None):
    """
    Vectorized single-pass ultra-fast pre-cache engine for all colleges.
    Accurately routes EWS/TFWS only to available quotas and avoids phantom empty tables.
    """
    df = pd.read_excel(excel_path)
    ai_matches, ai_meta = build_ai_index(df, ai_path)
    code_col = find_code_col(df)
    inst_col = find_inst_name_col(df)
    city_col = find_col(df, ["CITY"]) or find_col(df, ["LOCATION"])
    branch_col = find_col(df, ["BRANCH", "NAME"]) or find_col(df, ["COURSE", "NAME"]) or find_col(df, ["BRANCH"])
    cat_col = find_col(df, ["CATEGORY"])

    # First pass: detect regular quota presence per college
    college_quota_meta = {}
    for _, row in df.iterrows():
        c_code = str(row.get(code_col, '')).strip().replace('.0','') if code_col and pd.notna(row.get(code_col)) else ''
        c_name = str(row.get(inst_col, '')).strip() if inst_col and pd.notna(row.get(inst_col)) else ''
        cat = str(row.get(cat_col, '')).strip().upper() if cat_col and pd.notna(row.get(cat_col)) else ''
        if not c_code and not c_name: continue
        key = c_code if c_code else c_name
        if key not in college_quota_meta:
            college_quota_meta[key] = {"has_hu": False, "has_ohu": False, "has_state": False, "has_ews": False, "has_tfws": False}
        m = college_quota_meta[key]
        route = classify_and_route(cat)
        if route:
            _, _, dest = route
            if dest == "HU": m["has_hu"] = True
            elif dest == "OHU": m["has_ohu"] = True
            elif dest == "STATE": m["has_state"] = True
            elif dest == "SPECIAL":
                if "EWS" in cat: m["has_ews"] = True
                if "TFWS" in cat: m["has_tfws"] = True

    colleges_cache = {}

    for _, row in df.iterrows():
        c_code = str(row.get(code_col, '')).strip().replace('.0','') if code_col and pd.notna(row.get(code_col)) else ''
        c_name = str(row.get(inst_col, '')).strip() if inst_col and pd.notna(row.get(inst_col)) else ''
        b_name = str(row.get(branch_col, '')).strip() if branch_col and pd.notna(row.get(branch_col)) else ''
        cat = str(row.get(cat_col, '')).strip() if cat_col and pd.notna(row.get(cat_col)) else ''
        
        if not c_code and not c_name: continue
        key = c_code if c_code else c_name
        
        if key not in colleges_cache:
            full_name = f"{c_code} - {c_name}" if c_code and c_name and c_code != c_name else (c_name or c_code)
            q_info = college_quota_meta.get(key, {})
            ext = ai_meta.get(norm_ai_code(c_code)) or ai_meta.get(key) or {}
            row_city = ""
            if city_col and pd.notna(row.get(city_col)):
                row_city = str(row.get(city_col)).strip()
                if row_city.lower() == 'nan': row_city = ""
            colleges_cache[key] = {
                "college_code": c_code,
                "college_name": c_name,
                "full_name": full_name,
                "city": row_city or ext.get("city", ""),
                "status": ext.get("status", ""),
                "university": ext.get("university", ""),
                "branches": {},
                "has_hu": q_info.get("has_hu", False),
                "has_ohu": q_info.get("has_ohu", False),
                "has_state": q_info.get("has_state", False),
                "has_ews": q_info.get("has_ews", False),
                "has_tfws": q_info.get("has_tfws", False),
                "has_ai": False,
                "ai_branches": 0,
                "blank_warnings": [],
                "web_meta": {
                    "nba_grading": "Available in MHT-CET Record",
                    "total_seats": "Available",
                    "website": "",
                    "university": ext.get("university", ""),
                    "status": ext.get("status", ""),
                    "city": ext.get("city", ""),
                    "summary": (f"{ext.get('university')} | {ext.get('status', '')}".strip(" |")
                                if ext.get("university") else f"MHT-CET Cutoff record for {full_name}.")
                }
            }
            
        c_entry = colleges_cache[key]
        if not b_name or b_name.lower() == "nan": continue
        
        res = classify_and_route(cat)
        if not res: continue
        
        gender, base, dest = res
        cell = get_cutoff_cell_from_row(row)
        if cell == "-----": continue
        
        b_dict = c_entry["branches"]
        if b_name not in b_dict:
            b_dict[b_name] = {"HU_M": {}, "HU_F": {}, "OHU_M": {}, "OHU_F": {}, "STATE_M": {}, "STATE_F": {}}
            
        b = b_dict[b_name]
        k = "HU_M" if gender == "Male" else "HU_F"
        k_ohu = "OHU_M" if gender == "Male" else "OHU_F"
        k_state = "STATE_M" if gender == "Male" else "STATE_F"
        
        q_meta = college_quota_meta.get(key, {})
        
        if dest == "SPECIAL":
            # Only put EWS/TFWS in tables that actually have regular quota!
            if q_meta.get("has_hu"):
                if base not in b[k]: b[k][base] = cell
            if q_meta.get("has_ohu"):
                if base not in b[k_ohu]: b[k_ohu][base] = cell
            if q_meta.get("has_state"):
                if base not in b[k_state]: b[k_state][base] = cell
            # If college has no regular HU/OHU/State (rare edge case), fallback to State
            if not q_meta.get("has_hu") and not q_meta.get("has_ohu") and not q_meta.get("has_state"):
                if base not in b[k_state]: b[k_state][base] = cell
        elif dest == "HU":
            if base not in b[k]: b[k][base] = cell
        elif dest == "OHU":
            if base not in b[k_ohu]: b[k_ohu][base] = cell
        elif dest == "STATE":
            if base not in b[k_state]: b[k_state][base] = cell

    for key, c_entry in colleges_cache.items():
        branches = c_entry["branches"]
        ai_map = ai_matches.get(norm_ai_code(key)) or ai_matches.get(key)
        added = apply_ai_to_branches(branches, ai_map)
        c_entry["has_ai"] = added > 0
        c_entry["ai_branches"] = added
        c_entry["has_hu"] = check_quota_has_regular_data(branches, "HU")
        c_entry["has_ohu"] = check_quota_has_regular_data(branches, "OHU")
        c_entry["has_state"] = check_quota_has_regular_data(branches, "STATE")
        
        warnings = []
        for b_name, b_data in branches.items():
            ews_val = b_data["HU_M"].get("EWS") or b_data["HU_F"].get("EWS") or b_data["STATE_M"].get("EWS") or b_data["STATE_F"].get("EWS")
            tfws_val = b_data["HU_M"].get("TFWS") or b_data["HU_F"].get("TFWS") or b_data["STATE_M"].get("TFWS") or b_data["STATE_F"].get("TFWS")
            if not ews_val or ews_val == "-----":
                warnings.append(f"EWS rank is blank for branch: {b_name}")
            if not tfws_val or tfws_val == "-----":
                warnings.append(f"TFWS rank is blank for branch: {b_name}")
        c_entry["blank_warnings"] = warnings

    return colleges_cache

def parse_college_data(excel_path, college_query, mode="ALL", ai_path=None):
    df = pd.read_excel(excel_path)
    ai_matches, ai_meta = build_ai_index(df, ai_path)
    
    code_col = find_code_col(df)
    inst_col = find_inst_name_col(df)
    city_col = find_col(df, ["CITY"]) or find_col(df, ["LOCATION"])
    branch_col = find_col(df, ["BRANCH", "NAME"]) or find_col(df, ["COURSE", "NAME"]) or find_col(df, ["BRANCH"])
    cat_col = find_col(df, ["CATEGORY"])
    
    college_df = df
    if college_query:
        q_str = str(college_query).strip().lower()
        matches = []
        if code_col:
            matches.append(df[code_col].astype(str).str.lower().str.replace('.0','',regex=False) == q_str)
        if inst_col:
            matches.append(df[inst_col].astype(str).str.lower().str.contains(q_str))
        if matches:
            combined_filter = matches[0]
            for m in matches[1:]:
                combined_filter = combined_filter | m
            filtered = df[combined_filter]
            if not filtered.empty:
                college_df = filtered

    code_val = ""
    inst_val = ""
    if code_col and not college_df[code_col].dropna().empty:
        code_val = str(college_df[code_col].dropna().iloc[0]).strip().replace('.0','')
    if inst_col and not college_df[inst_col].dropna().empty:
        inst_val = str(college_df[inst_col].dropna().iloc[0]).strip()
        
    full_name = f"{code_val} - {inst_val}" if code_val and inst_val and code_val != inst_val else (inst_val or code_val)
    if not full_name: full_name = "Selected College"

    # First pass: detect regular quota presence per college using classify_and_route
    college_df_cats = college_df[cat_col].dropna().astype(str).str.upper().tolist() if cat_col else []
    has_hu = False
    has_ohu = False
    has_state = False
    has_ews = any("EWS" in c for c in college_df_cats)
    has_tfws = any("TFWS" in c for c in college_df_cats)
    for c in college_df_cats:
        r = classify_and_route(c)
        if r:
            _, _, d = r
            if d == "HU": has_hu = True
            elif d == "OHU": has_ohu = True
            elif d == "STATE": has_state = True

    branches = {}
    cur_branch = ""
    
    for _, row in college_df.iterrows():
        if branch_col and pd.notna(row.get(branch_col)):
            val = str(row[branch_col]).strip()
            if val and val.lower() != "nan":
                cur_branch = val
                
        if not cur_branch: continue
        
        cat = row.get(cat_col)
        if pd.isna(cat): continue
        
        res = classify_and_route(cat)
        if not res: continue
        
        gender, base, dest = res
        cell = get_cutoff_cell_from_row(row)
        
        if cell == "-----": continue
        
        if cur_branch not in branches:
            branches[cur_branch] = {"HU_M": {}, "HU_F": {}, "OHU_M": {}, "OHU_F": {}, "STATE_M": {}, "STATE_F": {}}
        
        b = branches[cur_branch]
        k = "HU_M" if gender == "Male" else "HU_F"
        k_ohu = "OHU_M" if gender == "Male" else "OHU_F"
        k_state = "STATE_M" if gender == "Male" else "STATE_F"
        
        if dest == "SPECIAL":
            if has_hu:
                if base not in b[k]: b[k][base] = cell
            if has_ohu:
                if base not in b[k_ohu]: b[k_ohu][base] = cell
            if has_state:
                if base not in b[k_state]: b[k_state][base] = cell
            if not has_hu and not has_ohu and not has_state:
                if base not in b[k_state]: b[k_state][base] = cell
        elif dest == "HU":
            if base not in b[k]: b[k][base] = cell
        elif dest == "OHU":
            if base not in b[k_ohu]: b[k_ohu][base] = cell
        elif dest == "STATE":
            if base not in b[k_state]: b[k_state][base] = cell

    # AI (JEE) single-round data is merged into every quota table
    ai_key = norm_ai_code(code_val)
    ai_map = ai_matches.get(ai_key) or ai_matches.get(str(college_query).strip())
    ai_added = apply_ai_to_branches(branches, ai_map)
    ext = ai_meta.get(ai_key) or {}

    city_val = ""
    if city_col and not college_df[city_col].dropna().empty:
        city_val = str(college_df[city_col].dropna().iloc[0]).strip()
        if city_val.lower() == 'nan': city_val = ""
    if not city_val: city_val = ext.get("city", "")

    blank_warnings = []
    for b_name, b_data in branches.items():
        ews_val = b_data["HU_M"].get("EWS") or b_data["HU_F"].get("EWS") or b_data["STATE_M"].get("EWS") or b_data["STATE_F"].get("EWS")
        tfws_val = b_data["HU_M"].get("TFWS") or b_data["HU_F"].get("TFWS") or b_data["STATE_M"].get("TFWS") or b_data["STATE_F"].get("TFWS")
        
        if not ews_val or ews_val == "-----":
            blank_warnings.append(f"EWS rank is blank for branch: {b_name}")
        if not tfws_val or tfws_val == "-----":
            blank_warnings.append(f"TFWS rank is blank for branch: {b_name}")

    web_meta = fetch_college_web_info(full_name)
    web_meta["university"] = ext.get("university", "")
    web_meta["status"] = ext.get("status", "")
    web_meta["city"] = city_val
    if ext.get("university"):
        web_meta["summary"] = f"{ext.get('university')} | {ext.get('status', '')}".strip(" |")
    
    # Final verification: verify that regular data actually exists in table
    has_hu = check_quota_has_regular_data(branches, "HU")
    has_ohu = check_quota_has_regular_data(branches, "OHU")
    has_state = check_quota_has_regular_data(branches, "STATE")
    
    return {
        "college_code": code_val,
        "college_name": inst_val,
        "full_name": full_name,
        "city": city_val,
        "status": ext.get("status", ""),
        "university": ext.get("university", ""),
        "has_hu": has_hu,
        "has_ohu": has_ohu,
        "has_state": has_state,
        "has_ews": has_ews,
        "has_tfws": has_tfws,
        "has_ai": ai_added > 0,
        "ai_branches": ai_added,
        "branches": branches,
        "blank_warnings": blank_warnings,
        "web_meta": web_meta
    }

# ============================================================
# WORD DOCX GENERATOR
# ============================================================
def set_cell_shading(cell, color):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>'))

def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    b = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                  f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  f'</w:tcBorders>')
    tcPr.append(b)

def drop_border(cell, side):
    """Removes one side of a cell's border (used where two cells are merged)."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is None:
        return
    el = borders.find(qn(f'w:{side}'))
    if el is not None:
        el.set(qn('w:val'), 'nil')

def make_run(p, text, bold=False, size=8):
    run = p.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Arial'
    run.font._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    return run

def create_docx_header_logo(doc, logo_path):
    if logo_path and os.path.exists(logo_path):
        try:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run()
            run.add_picture(logo_path, width=Inches(1.0))
        except Exception as e:
            print(f"Error adding logo to docx: {e}")

def create_info_box(doc, name, web_meta=None):
    tbl = doc.add_table(rows=3, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    nba = web_meta.get("nba_grading", "Available in MHT-CET Record") if web_meta else "Available in MHT-CET Record"
    seats = web_meta.get("total_seats", "Available") if web_meta else "Available"
    university = (web_meta or {}).get("university", "")
    status = (web_meta or {}).get("status", "")
    city = (web_meta or {}).get("city", "")
    
    texts = [
        [("College Code & Name: ", False), (name, False)],
        [("NBA Grading: ", True), (nba, False)],
        [("Total Seats: ", True), (seats, False)],
        [("Available Cut-off: ", True), ("CET R1-R4 + AI (JEE) (Rank + %ile)", False)]
    ]
    widths = [Inches(4.0), Inches(2.0), Inches(1.5), Inches(2.5)]
    
    for i, parts in enumerate(texts):
        cell = tbl.cell(0, i)
        cell.width = widths[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
        for txt, b in parts: make_run(p, txt, bold=b)
        set_cell_shading(cell, "F2F2F2"); set_cell_border(cell)

    merged = tbl.cell(1, 0).merge(tbl.cell(2, 3))
    merged.text = ""
    p0 = merged.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0); p0.paragraph_format.space_after = Pt(0)
    parts = []
    if university: parts.append(("University: ", True)); parts.append((university, False))
    if status: parts.append(("    Status: ", True)); parts.append((status, False))
    if city: parts.append(("    City: ", True)); parts.append((city, False))
    for txt, b in parts:
        make_run(p0, txt, bold=b, size=7)
    for p in merged.paragraphs[1:]:
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
    set_cell_shading(merged, "F2F2F2"); set_cell_border(merged)

def create_rank_table(doc, title, branches, mode):
    m_key, f_key = ("HU_M", "HU_F") if mode == "HU" else ("OHU_M", "OHU_F") if mode == "OHU" else ("STATE_M", "STATE_F")
    
    # Check if there is ANY regular data for this table
    has_any_regular = False
    for b_name, b_data in branches.items():
        m_data, f_data = b_data[m_key], b_data[f_key]
        if any(m_data.get(c) for c in REGULAR_CATS) or any(f_data.get(c) for c in REGULAR_CATS):
            has_any_regular = True
            break
            
    if not has_any_regular:
        return # Skip generating table if all regular cells are empty/dashed!

    def fill_cell(p, val, bold=False, left=False):
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if left else WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        if "\n" in str(val):
            top, bottom = str(val).split("\n", 1)
            make_run(p, top, bold=bold)
            br = p.add_run()
            br.font.size = Pt(7)
            br.add_break()
            make_run(p, bottom, bold=bold, size=7)
        else:
            make_run(p, val, bold=bold)

    shared_cols = [2 + CAT_ORDER.index(c) for c in SHARED_COLS]
    data_rows = []
    shared_pairs = []          # (male_data_row, female_data_row, [shared values])
    for b_name, b_data in branches.items():
        m_data, f_data = b_data[m_key], b_data[f_key]
        has_m = any(m_data.get(c) for c in CAT_ORDER)
        has_f = any(f_data.get(c) for c in CAT_ORDER)
        
        if has_m:
            row = [b_name, "Male"] + [m_data.get(c, "-----") for c in CAT_ORDER]
            data_rows.append(row)
        if has_f:
            branch_label = "" if has_m else b_name
            row = [branch_label, "Female"] + [f_data.get(c, "-----") for c in CAT_ORDER]
            if has_m:
                # EWS / TFWS / AI (JEE) are one value per branch -> shown once
                # in the vertically merged cell created below.
                shared_vals = []
                for c in SHARED_COLS:
                    v = m_data.get(c) or f_data.get(c) or "-----"
                    shared_vals.append(v)
                    row[2 + CAT_ORDER.index(c)] = ""
                shared_pairs.append((len(data_rows) - 1, len(data_rows), shared_vals))
            data_rows.append(row)
            
    if not data_rows: return

    tbl = doc.add_table(rows=len(data_rows)+1, cols=14)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    w = [2.7, 0.7, 0.72, 0.72, 0.65, 0.65, 0.8, 0.8, 0.8, 0.65, 0.72, 0.65, 0.65, 0.65]
    for r in tbl.rows:
        for i, wi in enumerate(w): r.cells[i].width = Inches(wi)

    tbl._tbl.addprevious(
        parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="left"/><w:spacing w:before="120" w:after="60"/></w:pPr>'
                  f'<w:r><w:rPr><w:b/><w:sz w:val="22"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial"/></w:rPr>'
                  f'<w:t>{title}</w:t></w:r></w:p>'))

    for i, col in enumerate(TABLE_COLS):
        c = tbl.rows[0].cells[i]; c.text = ""
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
        make_run(p, col, bold=True)
        set_cell_shading(c, "4472C4"); set_cell_border(c)
        for run in p.runs:
            run._element.get_or_add_rPr().append(parse_xml(f'<w:color {nsdecls("w")} w:val="FFFFFF"/>'))

    for ri, rd in enumerate(data_rows):
        bg = "FFFFFF" if ri % 2 == 0 else "F2F7FC"
        for ci in range(14):
            c = tbl.rows[ri+1].cells[ci]; c.text = ""
            p = c.paragraphs[0]
            val = rd[ci] if ci < len(rd) else ""
            fill_cell(p, val, bold=(ci == 0), left=(ci == 0))
            set_cell_shading(c, bg); set_cell_border(c)

    # EWS / TFWS / AI (JEE): one value, centred across the Male+Female rows
    for m_idx, f_idx, vals in shared_pairs:
        for k, ci in enumerate(shared_cols):
            top = tbl.rows[m_idx + 1].cells[ci]
            bot = tbl.rows[f_idx + 1].cells[ci]
            merged = top.merge(bot)
            merged.text = ""
            fill_cell(merged.paragraphs[0], vals[k])
            merged.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            drop_border(top, 'bottom')     # no line through the merged cell
            drop_border(bot, 'top')

def add_page_break(doc):
    p = doc.add_paragraph()
    p._element.clear()
    p._element.append(parse_xml(f'<w:r {nsdecls("w")}><w:br w:type="page"/></w:r>'))

def generate_word_document(college_data, mode="ALL", logo_path=None, output_folder="College_Data"):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    doc = Document()
    for s in doc.sections:
        s.page_width, s.page_height = Cm(29.7), Cm(21.0)
        s.top_margin, s.bottom_margin = Inches(0.5), Inches(0.45)
        s.left_margin, s.right_margin = Inches(0.5), Inches(0.5)

    st = doc.styles['Normal']
    st.font.name = 'Arial'; st.font.size = Pt(9)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    st.paragraph_format.space_before = Pt(0); st.paragraph_format.space_after = Pt(0)

    name = college_data["full_name"]
    branches = college_data["branches"]
    web_meta = college_data.get("web_meta", {})

    has_hu = college_data.get("has_hu", True)
    has_ohu = college_data.get("has_ohu", True)
    has_state = college_data.get("has_state", True)

    create_docx_header_logo(doc, logo_path)
    create_info_box(doc, name, web_meta)

    if mode == "HU" and has_hu:
        create_rank_table(doc, "Home University Cut-Off", branches, "HU")
    elif mode == "OHU" and has_ohu:
        create_rank_table(doc, "Other Than Home University Cut-Off", branches, "OHU")
    elif mode == "STATE" and has_state:
        create_rank_table(doc, "State Level Cut-Off", branches, "STATE")
    elif mode in ("HU_OHU", "4"):
        if has_hu:
            create_rank_table(doc, "Home University Cut-Off", branches, "HU")
        if has_ohu:
            if has_hu: add_page_break(doc); create_info_box(doc, name, web_meta)
            create_rank_table(doc, "Other Than Home University Cut-Off", branches, "OHU")
    else: # ALL
        rendered_count = 0
        if has_hu:
            create_rank_table(doc, "Home University Cut-Off", branches, "HU")
            rendered_count += 1
        if has_ohu:
            if rendered_count > 0: add_page_break(doc); create_info_box(doc, name, web_meta)
            create_rank_table(doc, "Other Than Home University Cut-Off", branches, "OHU")
            rendered_count += 1
        if has_state:
            if rendered_count > 0: add_page_break(doc); create_info_box(doc, name, web_meta)
            create_rank_table(doc, "State Level Cut-Off", branches, "STATE")

    safe_name = re.sub(r'[<>:"/\\|?*]', '', name).strip()
    if not safe_name: safe_name = "College_Cutoff"
    
    file_name = f"{safe_name}.docx"
    file_path = os.path.join(output_folder, file_name)
    
    if os.path.exists(file_path):
        counter = 1
        while True:
            new_name = f"{safe_name}_{counter}.docx"
            new_path = os.path.join(output_folder, new_name)
            if not os.path.exists(new_path):
                shutil.move(file_path, new_path)
                break
            counter += 1
            
    doc.save(file_path)
    return file_path, os.path.basename(file_path)
