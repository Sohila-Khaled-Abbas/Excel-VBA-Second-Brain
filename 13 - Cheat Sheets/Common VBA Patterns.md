---
type: cheat-sheet
domain: VBA
topic: Best Practices
tags:
  - cheat-sheet
  - patterns
created: 2026-10-03
updated: 2026-10-03
---

# 📐 Common VBA Patterns & Boilerplates

## 1. Production Standard Submarine Boilerplate
```vba
Public Sub StandardPattern()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Logic...
    
CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

## 2. Check if Sheet Exists
```vba
Public Function SheetExists(sheetName As String, Optional wb As Workbook = Nothing) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
```
