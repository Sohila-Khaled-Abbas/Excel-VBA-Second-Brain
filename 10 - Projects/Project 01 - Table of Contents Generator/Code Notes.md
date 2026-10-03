---
type: project-doc
project: Project 01 - Table of Contents Generator
doc_type: code-notes
status: completed
created: 2026-10-03
updated: 2026-10-03
---

# Project 01: Core VBA Implementation Notes

```vba
Option Explicit

Public Sub GenerateTableOfContents()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wsTOC As Worksheet
    Dim ws As Worksheet
    Dim r As Long
    Dim sheetLink As String
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Check if TOC sheet exists; delete if present
    On Error Resume Next
    Set wsTOC = wb.Sheets("Table of Contents")
    On Error GoTo ErrorHandler
    
    If Not wsTOC Is Nothing Then
        wsTOC.Delete
    End If
    
    ' Add new TOC sheet at the first position
    Set wsTOC = wb.Sheets.Add(Before:=wb.Sheets(1))
    wsTOC.Name = "Table of Contents"
    
    ' Header formatting
    With wsTOC.Range("A1:B1")
        .Value = Array("#", "Worksheet Name")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(16, 124, 65) ' Excel Green
    End With
    
    r = 2
    For Each ws In wb.Worksheets
        If ws.Name <> wsTOC.Name And ws.Visible = xlSheetVisible Then
            ' Item number
            wsTOC.Cells(r, 1).Value = r - 1
            
            ' Two-way hyperlink target (quoted for spaces)
            sheetLink = "'" & ws.Name & "'!A1"
            
            wsTOC.Hyperlinks.Add _
                Anchor:=wsTOC.Cells(r, 2), _
                Address:="", _
                SubAddress:=sheetLink, _
                TextToDisplay:=ws.Name
                
            ' Optional: Add return link to target sheet
            ws.Range("Z1").Value = "Back to TOC"
            ws.Hyperlinks.Add _
                Anchor:=ws.Range("Z1"), _
                Address:="", _
                SubAddress:="'" & wsTOC.Name & "'!A1", _
                TextToDisplay:="« Back to TOC"
                
            r = r + 1
        End If
    Next ws
    
    wsTOC.Columns("A:B").AutoFit
    wsTOC.Activate
    wsTOC.Range("A1").Select
    
CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
ErrorHandler:
    MsgBox "TOC Generator error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```
