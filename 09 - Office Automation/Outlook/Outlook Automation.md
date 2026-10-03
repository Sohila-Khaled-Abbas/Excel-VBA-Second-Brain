---
type: concept
domain: VBA
topic: Office Suite
difficulty: advanced
status: learning
source: Udemy
course_section: "[[Section 15 - Interacting with Other Applications (Word, PowerPoint etc.)]]"
related_concepts:
  - "[[PDF Generation]]"
related_projects:
  - "[[Project 03 - Invoice Generation Tool]]"
tags:
  - vba
  - outlook
  - email
  - office-automation
created: 2026-10-03
updated: 2026-10-03
---

# Outlook Automation

> [!important]
> **Summary Statement**: Automates Microsoft Outlook via COM automation to generate personalized emails, attach dynamically exported PDF reports or invoices, and display or silently send messages.

---

## 💻 Late-Binding Outlook Email Pattern
Late binding (`CreateObject("Outlook.Application")`) does not require setting references in VBE, ensuring code runs without missing reference errors across different Office versions:

```vba
Sub SendAutomatedReport(recipientEmail As String, pdfPath As String)
    Dim outApp As Object
    Dim outMail As Object
    
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0) ' 0 = olMailItem
    
    With outMail
        .To = recipientEmail
        .Subject = "Monthly Performance Report - " & Format(Date, "mmmm yyyy")
        .Body = "Hello," & vbCrLf & vbCrLf & _
                "Please find attached your monthly performance statement." & vbCrLf & _
                "Best regards," & vbCrLf & "Finance Analytics Team"
        .Attachments.Add pdfPath
        .Display ' Or .Send for unattended background distribution
    End With
    
    Set outMail = Nothing
    Set outApp = Nothing
End Sub
```

---

## 🔗 Connected Notes
- [[PDF Generation]], [[Project 03 - Invoice Generation Tool]]
