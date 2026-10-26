---
title: Dynamic 2D Array Memory Processor
type: snippet
domain: code_cookbook
status: evergreen
tags:
  - vba/arrays
  - vba/memory
  - vba/performance
  - vba/snippet
related:
  - "[[Arrays]]"
  - "[[Performance Optimization]]"
created: 2026-10-26
updated: 2026-10-26
---

# Dynamic 2D Array Memory Processor

> **Purpose**: Ingest an entire worksheet range into memory in a single COM call, manipulate values in high-speed RAM, and flush results back to the sheet in one write operation. 50x–100x faster than cell-by-cell loops.

---

## Production Snippet

```vba
Option Explicit

Public Sub FastArrayDataTransform()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim sourceData As Variant
    Dim r As Long, c As Long
    Dim startTime As Double
    
    startTime = Timer
    Set ws = ThisWorkbook.Sheets("SalesData")
    
    ' 1. Identify active data boundary dynamically
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 2 Then Exit Sub ' No data rows
    
    ' 2. ONE COM READ: Ingest entire range into 1-based 2D Variant array
    ' Dimensions: sourceData(RowIndex, ColumnIndex)
    sourceData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
    
    ' 3. Process records completely in RAM
    For r = 1 To UBound(sourceData, 1)
        ' Example: Apply 10% discount if Quantity (Col 3) > 50
        If IsNumeric(sourceData(r, 3)) And sourceData(r, 3) > 50 Then
            sourceData(r, 4) = sourceData(r, 4) * 0.9 ' Unit Price (Col 4)
            sourceData(r, 5) = "Volume Discount Applied" ' Notes (Col 5)
        End If
    Next r
    
    ' 4. ONE COM WRITE: Dump modified array back to the grid in 1 operation
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value = sourceData
    
    MsgBox "Processed " & (lastRow - 1) & " rows in " & _
           Format(Timer - startTime, "0.00") & " seconds!", vbInformation, "RAM Processing Complete"
End Sub
```

---

## Key Array Memory Rules

1. **Range to Variant is always 1-based 2D**: Even if reading a single column (e.g., `Range("A1:A10")`), the resulting array has 2 dimensions: `(1 To 10, 1 To 1)`.
2. **UBound(arr, 1) vs UBound(arr, 2)**: Dimension 1 is Rows; Dimension 2 is Columns.
3. **Target Range sizing must match**: The destination range must have identical row and column counts as the array bounds, or Excel will truncate or tile values.
