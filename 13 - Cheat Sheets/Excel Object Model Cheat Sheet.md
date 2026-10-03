---
type: cheat-sheet
domain: VBA
topic: Object Model
tags:
  - cheat-sheet
  - object-model
  - vba
created: 2026-10-03
updated: 2026-10-03
---

# 🏛️ Excel Object Model Cheat Sheet

## Hierarchy Quick Map
`Application` → `Workbooks("Report.xlsx")` → `Worksheets("Data")` → `Range("A1:D10")`

## Core Range Operations
| Goal | Code Pattern |
| :--- | :--- |
| **Referencing** | `ws.Range("A1")` or `ws.Cells(r, c)` |
| **Last Row** | `ws.Cells(ws.Rows.Count, "A").End(xlUp).Row` |
| **Last Column** | `ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column` |
| **Current Block** | `ws.Range("A1").CurrentRegion` (Mimics Ctrl+A) |
| **Fast Value Transfer** | `wsTarget.Range("A1:C50").Value = wsSource.Range("A1:C50").Value` |
| **Clear Data Only** | `ws.Range("A2:Z100").ClearContents` (Preserves formats) |
| **Clear Everything** | `ws.Range("A2:Z100").Clear` (Deletes data & formatting) |

## Sheet Referencing
- `ThisWorkbook`: The workbook running the macro.
- `ActiveWorkbook`: The workbook currently visible on top.
- `Worksheets("Name")`: Tab name reference.
- `wsCodeName`: VBE CodeName reference (immune to user tab renaming).
