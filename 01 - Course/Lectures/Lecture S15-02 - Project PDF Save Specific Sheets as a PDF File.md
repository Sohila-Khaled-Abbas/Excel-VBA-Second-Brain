---
type: lecture
course: Excel VBA & Macros
section: "[[Section 15 - Interacting with Other Applications (Word, PowerPoint etc.)]]"
section_number: 15
lecture_number: 2
status: not-started
topics:
  - Interacting with Other Applications (Word, PowerPoint etc.)
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-15
---

# Lecture S15-02: Project PDF: Save Specific Sheets as a PDF File

> [!important]
> **Section Reference**: [[Section 15 - Interacting with Other Applications (Word, PowerPoint etc.)|Section 15: Interacting with Other Applications (Word, PowerPoint etc.)]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Project PDF: Save Specific Sheets as a PDF File

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Project PDF: Save Specific Sheets as a PDF File' within the context of Interacting with Other Applications (Word, PowerPoint etc.).

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Project PDF: Save Specific Sheets as a PDF File', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Create_PDF()

    Dim PDFTable As ListObject

    Dim PDFSheets() As String

    Dim c As Byte 'number of tabs to be exported

    Dim FileName As String

    On Error GoTo Handle:

    

    FileName = ThisWorkbook.Path & "\PDFExport"

    Set PDFTable = Worksheets("PDF").ListObjects("PDFTable")

    ReDim PDFSheets(1 To PDFTable.DataBodyRange.Rows.Count)

    'fill up the array

    For c = 1 To UBound(PDFSheets)

        PDFSheets(c) = PDFTable.DataBodyRange(c, 1).Value

    Next c

    Worksheets(PDFSheets).Select

    ActiveSheet.ExportAsFixedFormat xlTypePDF, FileName

    Worksheets("PDF").Select

    MsgBox "PDF file was created." & vbNewLine & "File is called PDFExport. It is saved on the same directory as this workbook.", , "Well Done"

    Exit Sub

    

Handle:

If Err.Number = 9 Then

    MsgBox "It looks like a tab name was not spelled correctly. Please double check."

Else

    MsgBox "Looks like an Error here..."

    

End If

End Sub
```

---

## 🔍 How It Works
1. **Initialization**: Variables and object references are explicitly defined and verified.
2. **Execution**: The procedure manipulates ranges, worksheets, or UI controls directly without unnecessary `.Select` or `.Activate` calls.
3. **Clean Completion**: Screen updates and system state are safely restored.

---

## ⚠️ Common Mistakes & Traps
> [!warning]
> - Relying on the `ActiveSheet` implicitly rather than qualifying sheet references explicitly.
> - Forgetting `Option Explicit`, allowing misspelled variables to compile into empty `Variant` values.
> - Leaving screen flickering active during high-frequency cell updates.

---

## 🧩 Real-World Use Cases
- **Data Analytics / Reporting**: Automating recurring data transformation, cleansing, and layout formatting pipelines.
- **Operational Automation**: Eliminating manual, error-prone spreadsheet tasks across business reporting workflows.

---

## 🔗 Related Concepts
- [[VBA Syntax]]
- [[VBA Object Model]]
- [[Error Handling]]
- [[Section 15 - Interacting with Other Applications (Word, PowerPoint etc.)]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
