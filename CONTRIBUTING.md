# Contributing to Excel VBA Second Brain

> **Welcome!** This vault is a structured Second Brain for Excel VBA mastery. Contributions that add technical depth, new production patterns, or improved wikilink connections are highly valued.

---

## 📋 Table of Contents

- [Repository Philosophy](#-repository-philosophy)
- [Public vs Private Content Rules](#-public-vs-private-content-rules)
- [Note Frontmatter Standard](#-note-frontmatter-standard)
- [Code Block Standards](#-code-block-standards)
- [Running Validation Scripts](#-running-validation-scripts)
- [Pull Request Checklist](#-pull-request-checklist)
- [Folder Map & Where to Add Notes](#-folder-map--where-to-add-notes)

---

## 🎯 Repository Philosophy

This vault applies the **5-Layer Understanding Model** to every note:

| Layer | Description |
| :---: | :--- |
| **1** | Technical Definition (precise, unambiguous) |
| **2** | Personal Understanding (in your own words) |
| **3** | Simple Analogy (non-technical explanation) |
| **4** | Real Example (working VBA code from an actual use case) |
| **5** | Analytics / Production Application (how it fits in a data pipeline or BI tool) |

---

## 🛡️ Public vs Private Content Rules

| ✅ Commit This | ❌ Never Commit This |
| :--- | :--- |
| Original synthesis, mental models, diagrams | Udemy video files or transcripts |
| Your own VBA code snippets & patterns | Official course slide PDFs |
| Architecture notes, debugging post-mortems | Proprietary `.xlsm` practice workbooks |
| Wikilinked concept notes | Verbatim instructor text copied from Udemy |

> All binary course files must remain in the local, git-ignored `_private/` directory.

---

## 📝 Note Frontmatter Standard

Every Markdown note **must** include valid YAML frontmatter:

```yaml
---
title: Descriptive Note Title
type: concept | snippet | experiment | project | reference | post-mortem
domain: vba_fundamentals | advanced_vba | code_cookbook | projects | labs
status: draft | active | evergreen
tags:
  - vba/topic
  - category/subtopic
related:
  - "[[Wikilink to Related Note]]"
created: YYYY-MM-DD
updated: YYYY-MM-DD
---
```

**Valid `type` values:**
- `concept` — Evergreen technical concept note
- `snippet` — Production-ready, copy-pasteable code pattern
- `experiment` — Lab result with hypothesis, code, and findings
- `project` — Project architectural documentation
- `reference` — External resource index or cheat sheet
- `post-mortem` — Failed attempt with root cause analysis

---

## 💻 Code Block Standards

All VBA code blocks must follow these rules:

```vba
Option Explicit          ' Required at top of every module

Public Sub ProcedureName()    ' Descriptive, PascalCase name
    ' ✅ Explicit variable declaration
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' ✅ Error handling architecture
    On Error GoTo ErrorHandler
    
    ' ✅ Speed optimization wrapper
    Application.ScreenUpdating = False
    
    ' --- Business Logic ---
    
CleanExit:
    Application.ScreenUpdating = True  ' Always restore state
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

| Rule | Requirement |
| :--- | :--- |
| `Option Explicit` | Mandatory in every module |
| Naming | `ProcedureName` (PascalCase), `myVariable` (camelCase), `MY_CONST` (SCREAMING_SNAKE_CASE) |
| Error Handling | All Subs >10 lines must include `On Error GoTo` |
| COM Cleanup | All `Set obj = CreateObject(...)` must end with `Set obj = Nothing` |
| Speed Wrapper | All data-heavy routines must toggle `ScreenUpdating`, `EnableEvents`, `Calculation` |

---

## 🔧 Running Validation Scripts

Before submitting a PR, run the vault linter to verify zero broken wikilinks:

```bash
# Navigate to vault root
cd "d:\courses\Data Science\Data Engineering\Udemy Courses\VBA & Marcos Course\Excel-VBA-Second-Brain"

# Validate all notes (wikilinks + frontmatter)
python scripts/validate_vault.py

# Regenerate all dashboard files from course_manifest.json
python scripts/sync_vault.py

# Mark a specific lecture as complete in course_manifest.json
python scripts/udemy_progress_sync.py --complete S04-03

# Mark an entire section as complete
python scripts/udemy_progress_sync.py --section 3
```

> **CI Requirement**: The `vault-ci.yml` GitHub Actions workflow runs `validate_vault.py` automatically on every push and pull request. PRs will not be merged if validation fails.

---

## ✅ Pull Request Checklist

Before opening a PR, verify each item:

- [ ] **No proprietary files**: No `.xlsm`, `.xlsx` (course files), PDFs, or video files added
- [ ] **Frontmatter valid**: Every new `.md` note has correct YAML frontmatter (all required fields present)
- [ ] **Wikilinks verified**: `python scripts/validate_vault.py` exits with code 0 (zero broken links)
- [ ] **Code standards**: All VBA code uses `Option Explicit`, PascalCase procedures, and error handling
- [ ] **No hardcoded paths**: No absolute Windows paths in notes; use `[[wikilink]]` or `config/` variables
- [ ] **Descriptive commit message**: Follow [Conventional Commits](https://www.conventionalcommits.org/) format

```
feat(cookbook): add regex text parser with UDF support
fix(links): correct broken wikilink target in Section 09 note
docs(readme): add dynamic architecture section
chore(scripts): update udemy_progress_sync CLI argument parsing
```

---

## 🗂️ Folder Map & Where to Add Notes

| Your Contribution | Target Folder |
| :--- | :--- |
| New VBA concept (variables, scope, etc.) | `02 - VBA Fundamentals/<topic>/` |
| Excel Object Model reference | `03 - Excel Object Model/<object>/` |
| Reusable production snippet | `12 - Code Cookbook/<category>/` |
| Lab experiment or benchmark | `11 - Labs & Experiments/Experiments/` |
| Failed attempt post-mortem | `11 - Labs & Experiments/Failed Attempts/` |
| Interview question | `14 - Questions & Insights/` |
| New cheat sheet | `13 - Cheat Sheets/` |
| External resource or link | `16 - Resources/` |
| Course section architecture note | `01 - Course/Course Sections/` |
