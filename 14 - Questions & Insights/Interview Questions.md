---
type: interview-prep
topic: Technical Interview
tags:
  - interview
  - career
  - vba
created: 2026-10-03
updated: 2026-10-03
---

# 💼 VBA Technical Interview Preparation

## 1. Fundamentals & Object Model
### Q: Explain the hierarchy of the Excel Object Model.
- **Strong Answer**: Starts at `Application` (global settings), descending into `Workbooks` collection, individual `Workbook`, `Worksheets` collection, `Worksheet`, and finally `Range` / `Cells`. Advanced answers highlight specialized child objects like `ListObjects`, `PivotTables`, and `ChartObjects`.

### Q: Why is `Option Explicit` considered non-negotiable in production VBA?
- **Strong Answer**: It enforces strict variable declaration at compile time. Without it, misspelled variable names create implicit `Variant` variables initialized to empty/0, causing silent mathematical errors that are extremely hard to track down.

---

## 2. Performance & Memory
### Q: How do you optimize a slow VBA macro processing 100,000 rows?
- **Strong Answer**: 
  1. Disable UI overhead: `ScreenUpdating = False`, `DisplayAlerts = False`.
  2. Suspend calculations and events: `Calculation = xlCalculationManual`, `EnableEvents = False`.
  3. Avoid COM roundtrips by loading the entire data range into an in-memory 2D `Variant` array (`arr = Range.Value`), iterating in RAM, and dumping back in one shot.
  4. Eliminate `.Select` and `.Activate` statements.

---

## 3. Reliability & Architecture
### Q: How do you handle errors in production VBA code?
- **Strong Answer**: Implement structured error handling via `On Error GoTo ErrorHandler`. Route execution through a centralized `CleanExit` label where global application states (`ScreenUpdating`, `Calculation`) are guaranteed to reset before `Exit Sub`.
