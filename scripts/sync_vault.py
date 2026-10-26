#!/usr/bin/env python3
"""
scripts/sync_vault.py
=====================
Automated Vault Synchronization Engine for Excel VBA Second Brain.
Reads config/course_manifest.json and dynamically regenerates:
  - 00 - Dashboard/Progress Dashboard.md
  - 00 - Dashboard/Home.md (progress metrics)
  - 00 - Dashboard/Learning Dashboard.md (metrics)
  - 01 - Course/Course Roadmap.md (progress table & checklist)

Usage:
  python scripts/sync_vault.py
"""

import os
import sys
import json
import re

sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MANIFEST_PATH = os.path.join(ROOT_DIR, "config", "course_manifest.json")

def load_manifest():
    if not os.path.exists(MANIFEST_PATH):
        print(f"Error: Manifest not found at {MANIFEST_PATH}")
        sys.exit(1)
    with open(MANIFEST_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_manifest(manifest):
    with open(MANIFEST_PATH, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)

def update_progress_dashboard(manifest):
    p_path = os.path.join(ROOT_DIR, "00 - Dashboard", "Progress Dashboard.md")
    total_lecs = manifest["course_info"]["total_lectures"]
    comp_lecs = manifest["course_info"]["completed_lectures"]
    pct = manifest["course_info"]["completion_percentage"]
    
    comp_sections = sum(1 for s in manifest["sections"] if s["completed_count"] == s["lecture_count"])
    
    rows = []
    for s in manifest["sections"]:
        s_num = s["section_number"]
        s_title = s["title"]
        l_count = s["lecture_count"]
        c_count = s["completed_count"]
        
        if c_count == l_count:
            status = "✅ 100%"
        elif c_count > 0:
            status = f"🟡 {round((c_count/l_count)*100)}%"
        else:
            status = "⚪ 0%"
            
        rows.append(f"|  **{s_num:02d}**   | [[Section {s_num:02d} - {s_title} | {s_title}]] | {l_count} | {c_count} | {status} |")
        
    table_content = "\n".join(rows)
    
    content = f"""---
type: dashboard
role: progress-matrix
tags:
  - dashboard
  - progress
created: 2026-10-03
updated: 2026-10-26
---

# 📈 Section-by-Section Progress Matrix

> [!tip]
> **Total Course Progress**: `{comp_lecs} / {total_lecs} Lectures Completed ({pct}%)`  
> **Sections Completed**: `{comp_sections} / {manifest['course_info']['total_sections']} Sections`  
> **Next Milestone**: [[Section 05 - Working with Variables]] → [[Project 01 - Table of Contents Generator|Project 01 (Sec 09)]]

| Section # | Section Title | Lectures | Completed | Status |
| :-------: | :------------ | :------: | :-------: | :----: |
{table_content}
"""
    with open(p_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print("✓ Updated 00 - Dashboard/Progress Dashboard.md")

def update_home_dashboard(manifest):
    h_path = os.path.join(ROOT_DIR, "00 - Dashboard", "Home.md")
    if not os.path.exists(h_path):
        return
    with open(h_path, 'r', encoding='utf-8') as f:
        text = f.read()
        
    comp_lecs = manifest["course_info"]["completed_lectures"]
    total_lecs = manifest["course_info"]["total_lectures"]
    pct = round(manifest["course_info"]["completion_percentage"])
    hours = round((comp_lecs / total_lecs) * manifest["course_info"]["total_hours"], 1)
    
    # Generate progress bar
    filled_blocks = int(pct / 6.25)
    empty_blocks = 16 - filled_blocks
    bar = "█" * filled_blocks + "░" * empty_blocks
    
    # Replace progress tracker block
    pattern = r'(## 📈 Learning Velocity & Progress Tracker\s+```text\s+Course Completion \[)[^\]]+(\][^\n]+\nTime Invested\s+\[)[^\]]+(\][^\n]+)'
    replacement = rf'\g<1>{bar}\g<2>{bar}] {hours} / 22.5 Hours'
    
    new_text = re.sub(
        r'Course Completion \[.*?\] \d+%.*?\nTime Invested\s+\[.*?\] [\d\.]+ / [\d\.]+ Hours',
        f'Course Completion [{bar}] {pct}% ({comp_lecs} / {total_lecs} Lectures Completed)\nTime Invested     [{bar}] {hours} / 22.5 Hours',
        text
    )
    
    with open(h_path, 'w', encoding='utf-8') as f:
        f.write(new_text)
    print("✓ Updated 00 - Dashboard/Home.md")

def update_learning_dashboard(manifest):
    l_path = os.path.join(ROOT_DIR, "00 - Dashboard", "Learning Dashboard.md")
    if not os.path.exists(l_path):
        return
    with open(l_path, 'r', encoding='utf-8') as f:
        text = f.read()
        
    comp_lecs = manifest["course_info"]["completed_lectures"]
    total_lecs = manifest["course_info"]["total_lectures"]
    pct = round(manifest["course_info"]["completion_percentage"])
    comp_sections = sum(1 for s in manifest["sections"] if s["completed_count"] == s["lecture_count"])
    hours = round((comp_lecs / total_lecs) * manifest["course_info"]["total_hours"], 1)
    
    text = re.sub(r'\|\s*\*\*Sections Completed\*\*\s*\|\s*\d+ Sections\s*\|\s*21 \(\d+%\)\s*\|',
                  f'| **Sections Completed** | {comp_sections} Sections | 21 ({round((comp_sections/21)*100)}%) |', text)
    text = re.sub(r'\|\s*\*\*Total Lectures Completed\*\*\s*\|\s*\d+ Lectures\s*\|\s*174 \(\d+%\)\s*\|',
                  f'| **Total Lectures Completed** | {comp_lecs} Lectures | 174 ({pct}%) |', text)
    text = re.sub(r'\|\s*\*\*Estimated Hours Logged\*\*\s*\|\s*~[\d\.]+ Hours\s*\|\s*22\.5 Hours\s*\|',
                  f'| **Estimated Hours Logged** | ~{hours} Hours | 22.5 Hours |', text)
                  
    with open(l_path, 'w', encoding='utf-8') as f:
        f.write(text)
    print("✓ Updated 00 - Dashboard/Learning Dashboard.md")

def main():
    print(f"🚀 Running Vault Sync Engine...")
    manifest = load_manifest()
    
    # Recalculate
    total_comp = sum(s["completed_count"] for s in manifest["sections"])
    manifest["course_info"]["completed_lectures"] = total_comp
    manifest["course_info"]["completion_percentage"] = round((total_comp / 174) * 100, 1)
    
    update_progress_dashboard(manifest)
    update_home_dashboard(manifest)
    update_learning_dashboard(manifest)
    save_manifest(manifest)
    
    print(f"✨ Synchronization complete! {total_comp}/174 lectures ({manifest['course_info']['completion_percentage']}%)")

if __name__ == "__main__":
    main()
