---
type: project-doc
project: Project 01 - Table of Contents Generator
doc_type: requirements
status: completed
created: 2026-10-03
updated: 2026-10-03
---

# Project 01: Requirements & Problem Statement

## 🏢 Business Problem
Large corporate Excel models frequently contain 20 to 80 separate tabs (e.g., department budgets, cost centers, raw feeds). Navigating these files manually is frustrating and inefficient. Users get lost, hidden sheets remain forgotten, and executive review is hindered.

## 🎯 Solution Objectives
Build a one-click macro that:
1. Dynamically scans the active workbook for all worksheets.
2. Creates or refreshes a dedicated "Table of Contents" tab at the very front (index 1).
3. Inserts clickable hyperlinks pointing to each worksheet.
4. Inserts a "Back to TOC" return hyperlink on cell A1 of every target worksheet.
5. Respects hidden sheets or handles them gracefully without crashing.
6. Integrates into the user's `PERSONAL.XLSB` so it is available globally in any open Excel file.
