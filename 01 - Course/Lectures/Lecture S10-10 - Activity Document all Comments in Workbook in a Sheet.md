---
type: lecture
course: Excel VBA & Macros
section: "[[Section 10 - Looping in VBA (controlling the flow of code)]]"
section_number: 10
lecture_number: 10
status: not-started
topics:
  - Looping in VBA (controlling the flow of code)
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-10
---

# Lecture S10-10: Activity: Document all Comments in Workbook in a Sheet

> [!important]
> **Section Reference**: [[Section 10 - Looping in VBA (controlling the flow of code)|Section 10: Looping in VBA (controlling the flow of code)]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Activity: Document all Comments in Workbook in a Sheet

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Activity: Document all Comments in Workbook in a Sheet' within the context of Looping in VBA (controlling the flow of code).

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Activity: Document all Comments in Workbook in a Sheet', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Document_All_Comments()

    Dim Sh As Worksheet

    Dim Cmt As Comment

    Dim r As Long 'counting rows

    Dim w As Byte ' counting sheets

    Application.ScreenUpdating = False

    ' Create a new sheet

    Set Sh = Worksheets.Add

    With Sh

        'Put header "Comment", "Address" & "Author" in A1, B1 & C1 respectively.

        .Cells(1, 1).Value = "Comment"

        .Cells(1, 2).Value = "Address"

        .Cells(1, 3).Value = "Author"

        

        'Put the content of each comment in the workbook in a separate cell - starting from A2.

        'Put the address of each comment - such as the worksheet name and cell address in a separate cell starting from B2 and author from C2.

        r = 2

        For w = 1 To Worksheets.Count

        

            For Each Cmt In Worksheets(w).Comments

                .Cells(r, 1).Value = Cmt.Text

                .Cells(r, 2).Value = Worksheets(w).Name & "! " & Cmt.Parent.Address

                .Cells(r, 3).Value = Cmt.Author

                r = r + 1

            Next Cmt

        Next w

        .Columns.AutoFit

        

    End With

      Application.ScreenUpdating = True

      

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
- [[Section 10 - Looping in VBA (controlling the flow of code)]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
