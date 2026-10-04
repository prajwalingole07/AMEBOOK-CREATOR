import os
import time
from cutoff_engine import scan_colleges_from_excel, parse_college_data
from app import load_cache, save_cache, load_config, make_cache_key, cache_entry_is_valid, sig_dict

cfg = load_config()
excel_path = cfg.get('saved_path')
ai_path = cfg.get('ai_excel_path', '')

print("Starting bulk pre-cache for all colleges...")
t0 = time.time()
colleges, total_rows = scan_colleges_from_excel(excel_path, ai_path)
print(f"Scanned {len(colleges)} colleges from {total_rows} rows in {time.time()-t0:.2f}s")

cache = load_cache()
stamps = sig_dict(excel_path, ai_path)

count = 0
for idx, c in enumerate(colleges):
    query = c.get('code') or c.get('name')
    for mode in ['ALL', 'HU', 'OHU', 'STATE', 'HU_OHU']:
        cache_key = make_cache_key(excel_path, ai_path, query, mode)
        if cache_key in cache and cache_entry_is_valid(cache[cache_key], excel_path, ai_path):
            continue
        parsed_data = parse_college_data(excel_path, query, mode, ai_path=ai_path)
        cache[cache_key] = {
            **stamps,
            'filename': f"{c.get('code', 'col')}.docx",
            'data': parsed_data
        }
        count += 1

    if (idx + 1) % 50 == 0:
        print(f"Progress: Cached {idx + 1}/{len(colleges)} colleges...")

save_cache(cache)
print(f"SUCCESS! Pre-cached {count} table entries for {len(colleges)} colleges in {time.time()-t0:.2f}s!")
