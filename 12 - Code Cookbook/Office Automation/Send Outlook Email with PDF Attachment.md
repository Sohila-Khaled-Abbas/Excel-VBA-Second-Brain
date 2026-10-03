---
type: snippet
language: VBA
category: office-automation
difficulty: advanced
use_case: Export active sheet to PDF and dispatch via Outlook email
objects:
  - Worksheet
  - Outlook.Application
tags:
  - vba
  - snippet
  - outlook
  - pdf
created: 2026-10-03
updated: 2026-10-03
---

# Snippet: Automated PDF Export and Outlook Email Dispatch

## 🎯 Problem Statement
Eliminates manual printing to PDF, opening Outlook, writing repetitive email bodies, and attaching documents.

---

## 💻 VBA Code Implementation

```vba
Option Explicit

Public Sub ExportSheetToPDFAndEmail(wsTarget As Worksheet, toEmail As String, reportSubject As String)
    Dim pdfPath As String
    Dim outApp As Object
    Dim outMail As Object
    
    On Error GoTo ErrorHandler
    
    ' Build temp file path
    pdfPath = Environ("TEMP") & "\" & wsTarget.Name & "_" & Format(Now, "YYYYMMDD_HHMMSS") & ".pdf"
    
    ' 1. Export target sheet to PDF
    wsTarget.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    ' 2. Instantiate Outlook via Late-Binding
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0) ' 0 = olMailItem
    
    With outMail
        .To = toEmail
        .Subject = reportSubject
        .Body = "Hi team," & vbCrLf & vbCrLf & _
                "Please review the attached automated executive report." & vbCrLf & vbCrLf & _
                "Generated on: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
        .Attachments.Add pdfPath
        .Display ' Change to .Send for unattended distribution
    End With
    
CleanExit:
    Set outMail = Nothing
    Set outApp = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "Failed to export and send email: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

---

## 🔗 Related Notes
- [[Outlook Automation]], [[PDF Generation]], [[Project 03 - Invoice Generation Tool]]
