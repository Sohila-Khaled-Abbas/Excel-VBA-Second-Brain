---
type: project-doc
project: Project 02 - Regional Sales Reporting Tool
doc_type: code-notes
status: in-progress
created: 2026-10-03
updated: 2026-10-03
---

# Project 02: Core VBA Implementation Highlights

### 1. MultiSelect Ingestion Routine
```vba
Dim fd As FileDialog
Dim vFile As Variant
Dim wbSource As Workbook
Dim wsDest As Worksheet: Set wsDest = ThisWorkbook.Sheets("RawData")
Dim nextRow As Long

Set fd = Application.FileDialog(msoFileDialogFilePicker)
fd.AllowMultiSelect = True
fd.Filters.Clear
fd.Filters.Add "Excel Workbooks", "*.xlsx; *.xlsm"

If fd.Show = -1 Then
    For Each vFile In fd.SelectedItems
        Set wbSource = Workbooks.Open(Filename:=CStr(vFile), ReadOnly:=True)
        nextRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row + 1
        
        ' Copy without clipboard
        Dim rngSource As Range
        Set rngSource = wbSource.Sheets(1).Range("A2").CurrentRegion
        Set rngSource = rngSource.Offset(1, 0).Resize(rngSource.Rows.Count - 1)
        
        rngSource.Copy wsDest.Cells(nextRow, 1)
        wbSource.Close SaveChanges:=False
    Next vFile
End If
```

### 2. Guaranteed Semicolon Delimited CSV Export
```vba
Sub ExportSemicolonCSV(ws As Worksheet, filePath As String)
    Dim fNum As Integer: fNum = FreeFile
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim lineStr As String
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Open filePath For Output As #fNum
    For r = 1 To lastRow
        lineStr = ""
        For c = 1 To lastCol
            lineStr = lineStr & ws.Cells(r, c).Text & IIf(c = lastCol, "", ";")
        Next c
        Print #fNum, lineStr
    Next r
    Close #fNum
End Sub
```
