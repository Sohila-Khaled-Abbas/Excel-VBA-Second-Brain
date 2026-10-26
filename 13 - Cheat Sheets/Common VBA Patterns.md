---
type: cheat-sheet
domain: VBA
topic: Reusable Automation Patterns
tags:
  - cheat-sheet
  - patterns
  - boilerplate
created: 2026-10-03
updated: 2026-10-26
---

# 📐 Common VBA Patterns & Boilerplates

---

## 1. High-Performance Screen & Calculation Boost

```vba
Public Sub BoostSpeed(ByVal enableBoost As Boolean)
    With Application
        .ScreenUpdating = Not enableBoost
        .DisplayAlerts = Not enableBoost
        .EnableEvents = Not enableBoost
        .Calculation = IIf(enableBoost, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub
```

---

## 2. Check If Sheet Exists (Safe Function)

```vba
Public Function SheetExists(ByVal sheetName As String, Optional ByVal wb As Workbook = Nothing) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
```

---

## 3. Check If File / Folder Exists

```vba
' Check File
If Dir("C:\Reports\Sales.xlsx") <> "" Then
    ' File exists!
End If

' Check Folder
If Dir("C:\Reports", vbDirectory) <> "" Then
    ' Folder exists!
End If
```

---

## 4. Multi-File Dialog File Picker

```vba
Dim fd As FileDialog
Dim vFile As Variant

Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
    .AllowMultiSelect = True
    .Filters.Clear
    .Filters.Add "Excel Files", "*.xlsx; *.xlsm"
    .Title = "Select Workbooks to Ingest"
    
    If .Show = -1 Then
        For Each vFile In .SelectedItems
            Dim wbTemp As Workbook
            Set wbTemp = Workbooks.Open(Filename:=CStr(vFile), ReadOnly:=True)
            ' ... process wbTemp ...
            wbTemp.Close SaveChanges:=False
        Next vFile
    End If
End With
```

---

## 5. Low-Level Semicolon CSV Export via FreeFile

```vba
Public Sub ExportSemicolonCSV(ByVal ws As Worksheet, ByVal filePath As String)
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
