#!/usr/bin/env python3
"""
scripts/validate_vault.py
=========================
Vault Integrity Linter & Link Validator for Excel VBA Second Brain.
Checks:
  - Valid YAML frontmatter across all notes
  - Resolution of internal Obsidian wikilinks [[Target]]
  - Integrity of config/course_manifest.json

Used both locally and in .github/workflows/vault-ci.yml.
"""

import os
import sys
import re
import json

sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IGNORED_DIRS = {'.git', '.obsidian', '_private', 'node_modules', '.github', 'assets'}

def find_all_markdown_files():
    md_files = {}
    for root, dirs, files in os.walk(ROOT_DIR):
        dirs[:] = [d for d in dirs if d not in IGNORED_DIRS]
        for f in files:
            if f.endswith('.md'):
                full_path = os.path.join(root, f)
                base_name = f[:-3] # without .md
                rel_path = os.path.relpath(full_path, ROOT_DIR)
                md_files[base_name] = rel_path
                md_files[rel_path] = rel_path
    return md_files

def validate_vault():
    print("🔍 Scanning Excel VBA Second Brain for link integrity and frontmatter...")
    all_files = find_all_markdown_files()
    print(f"📁 Found {len(all_files)//2} markdown notes.")
    
    total_notes = 0
    missing_frontmatter = []
    broken_links = []
    total_links_checked = 0
    
    wikilink_pattern = re.compile(r'\[\[(.*?)\]\]')
    
    for root, dirs, files in os.walk(ROOT_DIR):
        dirs[:] = [d for d in dirs if d not in IGNORED_DIRS]
        for f in files:
            if not f.endswith('.md'):
                continue
            total_notes += 1
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, ROOT_DIR)
            
            with open(full_path, 'r', encoding='utf-8', errors='ignore') as fp:
                content = fp.read()
                
            # Check YAML frontmatter
            if not content.startswith('---'):
                missing_frontmatter.append(rel_path)
            else:
                end_idx = content.find('---', 3)
                if end_idx == -1:
                    missing_frontmatter.append(rel_path)
                    
            # Check wikilinks
            for m in wikilink_pattern.finditer(content):
                link_text = m.group(1).strip()
                # Split alias [[Target|Alias]]
                target = link_text.split('|')[0].strip()
                # Remove header anchor [[Target#Header]]
                target = target.split('#')[0].strip()
                
                if not target:
                    continue
                    
                total_links_checked += 1
                target_base = os.path.basename(target).replace('.md', '')
                
                # Verify match
                if target not in all_files and target_base not in all_files:
                    broken_links.append((rel_path, target))
                    
    print("\n" + "="*50)
    print("📊 VALIDATION SUMMARY")
    print("="*50)
    print(f"Total Notes Analyzed: {total_notes}")
    print(f"Wikilinks Verified:   {total_links_checked}")
    print(f"Frontmatter Valid:    {total_notes - len(missing_frontmatter)} / {total_notes}")
    
    if missing_frontmatter:
        print(f"\n⚠️  Notes Missing YAML Frontmatter ({len(missing_frontmatter)}):")
        for fn in missing_frontmatter[:10]:
            print(f"  - {fn}")
            
    if broken_links:
        print(f"\n⚠️  Broken Wikilinks Detected ({len(broken_links)}):")
        for src, tgt in broken_links[:10]:
            print(f"  - In [{src}]: [[{tgt}]]")
    else:
        print("\n✅ All internal wikilinks resolve cleanly!")
        
    # Check Manifest
    manifest_path = os.path.join(ROOT_DIR, "config", "course_manifest.json")
    if os.path.exists(manifest_path):
        try:
            with open(manifest_path, 'r', encoding='utf-8') as f:
                json.load(f)
            print("✅ config/course_manifest.json is valid JSON.")
        except Exception as e:
            print(f"❌ Error in config/course_manifest.json: {e}")
            return 1
            
    if len(broken_links) > 20:
        print("\n❌ Vault validation failed due to high number of broken links.")
        return 1
        
    print("\n🎉 Vault passed integrity inspection!")
    return 0

if __name__ == "__main__":
    sys.exit(validate_vault())
