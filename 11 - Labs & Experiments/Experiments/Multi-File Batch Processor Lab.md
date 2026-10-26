---
title: Multi-File Batch Processor Lab
type: experiment
domain: labs
status: active
tags:
  - vba/files
  - vba/experiment
  - vba/dir
  - vba/arrays
related:
  - "[[File Automation]]"
  - "[[Dynamic 2D Array Memory Processor]]"
  - "[[Batch Folder Scanner]]"
created: 2026-10-26
updated: 2026-10-26
---

# Multi-File Batch Processor Lab

> **Lab Goal**: Benchmark and compare two strategies for processing 50+ Excel files — cell-by-cell loop vs. Variant Array loading — measuring time, file handles, and error robustness.

---

## Background & Hypothesis
In regional reporting automation (Project 02), the system opens each monthly file individually, copies data, closes the file, and loops. For large file sets (30–50 workbooks), each `Workbooks.Open` call takes ~1.2–2.0 seconds including the COM load time.

**Hypothesis**: Pre-staging file paths into a string array and batch-validating with `Dir()` before opening will reduce total processing time by eliminating redundant error checks inside the main loop.

---

## Experiment Code: Strategy A — Naive Sequential Open

```vba
Option Explicit

Public Sub Strategy_A_NaiveOpen()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim destRow As Long
    Dim startTime As Double
    
    startTime = Timer
    folderPath = "C:\Data\MonthlySales\"
    fileName = Dir(folderPath & "*.xlsx")
    destRow = 2
    
    Do While fileName <> ""
        ' No pre-validation: just open and copy
        Set wb = Workbooks.Open(folderPath & fileName)
        Set ws = wb.Sheets(1)
        
        ' Copy first 10 columns of data row 2 to summary
        ThisWorkbook.Sheets("Summary").Cells(destRow, 1).Resize(1, 10).Value = _
            ws.Range("A2:J2").Value
        
        destRow = destRow + 1
        wb.Close SaveChanges:=False
        
        fileName = Dir()
    Loop
    
    Debug.Print "Strategy A Time: " & Format(Timer - startTime, "0.000") & " seconds"
End Sub
```

---

## Experiment Code: Strategy B — Pre-staged Array + Batch Open

```vba
Public Sub Strategy_B_BatchArray()
    Dim folderPath As String
    Dim fileNames() As String
    Dim fileCount As Long
    Dim i As Long
    Dim wb As Workbook
    Dim destRow As Long
    Dim startTime As Double
    
    startTime = Timer
    folderPath = "C:\Data\MonthlySales\"
    
    ' Pass 1: Collect all valid file paths into array
    ReDim fileNames(0 To 200)
    fileCount = 0
    Dim f As String
    f = Dir(folderPath & "*.xlsx")
    Do While f <> ""
        fileNames(fileCount) = folderPath & f
        fileCount = fileCount + 1
        f = Dir()
    Loop
    ReDim Preserve fileNames(0 To fileCount - 1)
    
    Debug.Print "Found " & fileCount & " files."
    
    ' Pass 2: Process validated list
    destRow = 2
    For i = 0 To fileCount - 1
        Set wb = Workbooks.Open(fileNames(i))
        ThisWorkbook.Sheets("Summary").Cells(destRow, 1).Resize(1, 10).Value = _
            wb.Sheets(1).Range("A2:J2").Value
        destRow = destRow + 1
        wb.Close SaveChanges:=False
    Next i
    
    Debug.Print "Strategy B Time: " & Format(Timer - startTime, "0.000") & " seconds"
End Sub
```

---

## Results Log

| Run | Strategy | File Count | Total Time | Avg/File | Notes |
| :-- | :------- | ---------: | ---------: | -------: | :---- |
| 1 | A — Naive | 42 | ~58.2s | ~1.386s | Dir loop overhead noticed |
| 2 | B — Array | 42 | ~52.1s | ~1.240s | Marginal but consistent |
| 3 | B — Array | 42 | ~51.4s | ~1.223s | Confirmed: ~11% faster at scale |

---

## Conclusion
The pre-staging approach (Strategy B) shows consistent improvement (~11%) at 42-file scale. The main bottleneck remains `Workbooks.Open` COM call itself. For >100 files, consider using `Power Query` batch import via `VBA.Refresh` trigger instead of opening each workbook individually.

**Next Experiment**: Test `GetObject` on already-open workbooks to skip the COM open overhead.
