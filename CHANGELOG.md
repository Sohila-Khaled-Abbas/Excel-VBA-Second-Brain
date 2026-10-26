# Changelog

All notable changes to **Excel VBA Second Brain** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

### Planned
- Dataview-powered auto-generated MOC for each section
- Anki flashcard export script for the Interview Question Bank
- GitHub Actions badge for note count and wikilink health

---

## [v2.0.0] — 2026-10-26

### 🚀 Major Release — Dynamic Non-Hardcoded Architecture & Full Curriculum

#### Added
- **Dynamic Engine**: `config/course_manifest.json` — single source of truth for all 21 sections, 174 lectures, and completion status. No more hardcoded progress percentages in dashboards.
- **Dynamic Engine**: `config/vault_config.json` — vault-level configuration (vault root, author, course metadata).
- **Dynamic Sync**: `scripts/sync_vault.py` — regenerates all 4 dashboard files from `course_manifest.json` on demand.
- **Vault Linter**: `scripts/validate_vault.py` — YAML frontmatter checker and wikilink integrity validator (verified: 0 broken links across 339 notes, 2,565 wikilinks).
- **Udemy Progress Sync**: `scripts/udemy_progress_sync.py` — CLI tool to mark lectures/sections complete from Excel outline, JSON export, or per-lecture commands.
- **Browser Bookmarklet**: `scripts/udemy_progress_export.js` — 1-click console exporter to capture live Udemy progress to JSON.
- **GitHub Workflows**: `.github/workflows/vault-ci.yml` (automated wikilink & frontmatter validation on push/PR).
- **GitHub Workflows**: `.github/workflows/sync-progress.yml` (manual workflow dispatch for dashboard regeneration).
- **GitHub Community Files**: Issue templates (bug report, course question, snippet request), PR template, `CODE_OF_CONDUCT.md`, `SECURITY.md`.
- **Advanced Code Cookbook**: 5 new production-ready snippets — Universal Speed Booster Wrapper, Dynamic 2D Array Memory Processor, RegEx Text Parser, Safe File Export with Overwrite Guard, Multi-File Batch Processor Lab.
- **Advanced Architecture Notes**: Early vs Late Binding Architecture, Bridging VBA with Power Query and Python.
- **Interview Question Bank**: 7 core VBA interview questions with folded answers and self-assessment tracker.
- **Failed Attempt Post-Mortem**: Worksheet_Change Infinite Loop — root cause, fix, and generalized rule.
- **Resources**: Udemy Progress Sync Guide, VBA Q&A and Debugging Strategy, Course Materials & Workbooks Guide.

#### Changed
- All 174 lecture notes enriched with authentic VBA code extracted from course workbooks.
- All 21 section files rewritten with curriculum objectives, architectural takeaways, pitfalls, and practice file mappings.
- `README.md` upgraded with shields, Mermaid architecture diagram, full 21-section roadmap table, and dynamic architecture documentation.
- `CONTRIBUTING.md` expanded with style rules, frontmatter format, and script usage instructions.

---

## [v1.5.0] — 2026-10-25

### 🧠 Second Brain Enrichment — Cheat Sheets & Projects

#### Added
- All 6 cheat sheets enriched: VBA Syntax, Excel Object Model, Loops, Arrays, Debugging, Common VBA Patterns.
- Project 01 (Table of Contents), Project 02 (Regional Reporting ETL), Project 03 (UserForm Invoicing) — architectural documentation with authentic course code.
- `16 - Resources/References.md` expanded with categorized external resources.

---

## [v1.0.0] — 2026-10-24

### 🎉 Initial Release — Vault Foundation

#### Added
- Full vault directory structure (00–21 folders + `99 - Templates/`).
- `.obsidian/` configuration with Minimal theme, custom CSS snippets (Excel green `#107C41` palette).
- Initial set of concept notes across VBA Fundamentals, Excel Object Model, Control Flow, and Debugging.
- 5 Dashboard files: `Home.md`, `Learning Dashboard.md`, `Progress Dashboard.md`, `Knowledge Map.md`, `Course Roadmap.md`.
- `README.md`, `LICENSE` (MIT), `CONTRIBUTING.md`, `.gitignore`.
- Git repository initialized and linked to `https://github.com/Sohila-Khaled-Abbas/Excel-VBA-Second-Brain`.

[Unreleased]: https://github.com/Sohila-Khaled-Abbas/Excel-VBA-Second-Brain/compare/v2.0.0...HEAD
[v2.0.0]: https://github.com/Sohila-Khaled-Abbas/Excel-VBA-Second-Brain/compare/v1.5.0...v2.0.0
[v1.5.0]: https://github.com/Sohila-Khaled-Abbas/Excel-VBA-Second-Brain/compare/v1.0.0...v1.5.0
[v1.0.0]: https://github.com/Sohila-Khaled-Abbas/Excel-VBA-Second-Brain/releases/tag/v1.0.0
