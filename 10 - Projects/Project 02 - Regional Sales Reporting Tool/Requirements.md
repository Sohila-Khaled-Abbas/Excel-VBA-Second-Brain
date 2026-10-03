---
type: project-doc
project: Project 02 - Regional Sales Reporting Tool
doc_type: requirements
status: in-progress
created: 2026-10-03
updated: 2026-10-03
---

# Project 02: Requirements & Problem Statement

## 🏢 Business Problem
Every month-end, the regional sales organization receives decentralized Excel workbooks from independent retail partners (BereKleid, LucasBasics, MetaCreations, UrbanRight). Analysts waste 6-8 hours manually opening files, copying data into a consolidation sheet, refreshing PivotTables, generating executive PDFs, and formatting CSV exports with semicolon delimiters for the corporate SAP ERP system.

## 🎯 Solution Pipeline Requirements
Build an end-to-end automated ETL reporting pipeline that:
1. Prompts the user to multi-select regional source files (`Application.FileDialog(msoFileDialogFilePicker)` with `AllowMultiSelect = True`).
2. Clears previous staging tables and imports all transactions dynamically.
3. Automatically populates a standardized master template.
4. Generates both interactive PivotTable reports and frozen (hardcoded) executive distribution files.
5. Exports a clean, standardized text file with a fixed **semicolon (`;`) delimiter** for ERP ingestion.
6. Records an audit trail of completed tasks on an activity log tab.
7. Protects sensitive tabs with a password and hides technical calculations.
