---
type: snippet
language: VBA
category: patterns
difficulty: beginner
use_case: Dynamically detect data boundaries to avoid hardcoded ranges
objects:
  - Range
  - Worksheet
tags:
  - vba
  - snippet
  - cookbook
  - patterns
created: 2026-10-03
updated: 2026-10-03
---

# Snippet: Dynamic Last Used Row and Column

## 🎯 Problem Statement
Hardcoding row boundaries (`Range("A2:D50")`) breaks macros whenever row volumes expand or shrink. This pattern provides bulletproof, dynamic boundary detection.

---

## 💻 VBA Code Implementation

```vba
Option Explicit

Public Function GetLastRow(ws As Worksheet, Optional colAnchor As Variant = 1) As Long
    ' Purpose: Finds last non-empty row in a specific column (default column A)
    GetLastRow = ws.Cells(ws.Rows.Count, colAnchor).End(xlUp).Row
End Function

Public Function GetLastColumn(ws As Worksheet, Optional rowAnchor As Long = 1) As Long
    ' Purpose: Finds last non-empty column in a specific row (default row 1)
    GetLastColumn = ws.Cells(rowAnchor, ws.Columns.Count).End(xlToLeft).Column
End Function

Public Function GetSheetTrueLastRow(ws As Worksheet) As Long
    ' Purpose: Finds the absolute last used row across ANY column on the sheet
    Dim rngFind As Range
    Set rngFind = ws.Cells.Find(What:="*", After:=ws.Range("A1"), _
                                LookIn:=xlFormulas, LookAt:=xlPart, _
                                SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not rngFind Is Nothing Then
        GetSheetTrueLastRow = rngFind.Row
    Else
        GetSheetTrueLastRow = 1
    End If
End Function
```

---

## 🔍 Line-by-Line Breakdown
- `ws.Rows.Count`: Returns maximum row count of the sheet (1,048,576 in modern Excel), ensuring compatibility across older and newer Excel formats.
- `.End(xlUp)`: Equivalent to clicking the bottom-most cell in the column and pressing `Ctrl + Up Arrow`.
- `.Find(What:="*", ...)`: Comprehensive wildcard search scanning backwards from bottom-right to top-left.

---

## 🔗 Related Notes
- [[Range]], [[Cells]], [[Common VBA Patterns]]
