import os
import time
from cutoff_engine import fast_precache_all
from app import save_cache, load_config, make_cache_key, sig_dict

cfg = load_config()
excel_path = cfg.get('saved_path')
ai_path = cfg.get('ai_excel_path', '')

t0 = time.time()
print(f"CET  : {excel_path}")
print(f"AI   : {ai_path}")
print("Rebuilding cache with CET + AI (JEE) data...")
colleges_dict = fast_precache_all(excel_path, ai_path)

cache = {}
stamps = sig_dict(excel_path, ai_path)
ai_total = 0
for key, parsed_data in colleges_dict.items():
    code_val = parsed_data.get('college_code') or 'college'
    ai_total += parsed_data.get('ai_branches', 0)
    for mode in ['ALL', 'HU', 'OHU', 'STATE', 'HU_OHU']:
        cache_key = make_cache_key(excel_path, ai_path, key, mode)
        cache[cache_key] = {
            **stamps,
            'filename': f"{code_val}.docx",
            'data': parsed_data
        }
save_cache(cache)
print(f"DONE in {time.time()-t0:.2f}s! Colleges: {len(colleges_dict)} | AI (JEE) branch cut-offs: {ai_total}")
