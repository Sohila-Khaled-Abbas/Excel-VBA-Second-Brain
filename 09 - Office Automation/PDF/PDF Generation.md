---
type: concept
domain: VBA
topic: Office Suite
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 15 - Interacting with Other Applications (Word, PowerPoint etc.)]]"
related_concepts:
  - "[[Outlook Automation]]"
related_projects:
  - "[[Project 03 - Invoice Generation Tool]]"
tags:
  - vba
  - pdf
  - reporting
created: 2026-10-03
updated: 2026-10-03
---

# PDF Generation (`ExportAsFixedFormat`)

> [!important]
> **Summary Statement**: Exports worksheets or specific ranges to locked, publication-ready PDF documents using Excel's native `ExportAsFixedFormat` method.

---

## 💻 Exporting a Single Sheet to PDF
```vba
Sub ExportInvoicePDF(ws As Worksheet, targetPath As String)
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=targetPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
End Sub
```

---

## 🔗 Connected Notes
- [[Outlook Automation]], [[Project 03 - Invoice Generation Tool]]
