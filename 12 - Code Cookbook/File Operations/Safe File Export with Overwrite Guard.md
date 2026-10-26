---
title: Safe File Export with Overwrite Guard
type: snippet
domain: code_cookbook
status: evergreen
tags:
  - vba/files
  - vba/export
  - vba/defensive-programming
  - vba/snippet
related:
  - "[[File Operations]]"
  - "[[Defensive Programming]]"
created: 2026-10-26
updated: 2026-10-26
---

# Safe File Export with Overwrite Guard

> **Purpose**: Export dynamic workbooks, CSVs, or PDFs while defensively checking if the target file path already exists, prompting the user for confirmation, and handling file lock / permission exceptions cleanly.

---

## Production Procedure

```vba
Option Explicit

Public Sub ExportSheetSafely(ByVal ws As Worksheet, ByVal TargetPath As String)
    Dim userResponse As VbMsgBoxResult
    Dim targetWb As Workbook
    
    On Error GoTo ErrorHandler
    
    ' 1. Check if target file already exists
    If Dir(TargetPath) <> "" Then
        userResponse = MsgBox("The file already exists:" & vbCrLf & TargetPath & vbCrLf & vbCrLf & _
                              "Do you want to overwrite it?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, _
                              "Confirm Overwrite")
        If userResponse = vbNo Then
            MsgBox "Export cancelled by user.", vbInformation, "Cancelled"
            Exit Sub
        End If
    End If
    
    ' 2. Copy worksheet into a new standalone workbook
    ws.Copy
    Set targetWb = ActiveWorkbook
    
    ' 3. Prevent Excel prompt since we already handled user confirmation
    Application.DisplayAlerts = False
    targetWb.SaveAs Filename:=TargetPath, FileFormat:=xlOpenXMLWorkbook ' .xlsx
    targetWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    MsgBox "File exported successfully to:" & vbCrLf & TargetPath, vbInformation, "Export Succeeded"
    
CleanExit:
    Application.DisplayAlerts = True
    Set targetWb = Nothing
    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    If Err.Number = 1004 Then
        MsgBox "Cannot write file. The file may be currently open in another program.", vbCritical, "File Locked"
    Else
        MsgBox "Export failed: " & Err.Description, vbCritical, "Export Error"
    End If
    Resume CleanExit
End Sub
```
