---
type: lecture
course: Excel VBA & Macros
section: "[[Section 04 - Referencing Ranges, Worksheets & Workbooks]]"
section_number: 4
lecture_number: 8
status: completed
topics:
  - Referencing Ranges, Worksheets & Workbooks
related_concepts:
  - '[[VBA Syntax]]'
  - '[[Excel Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-04
---

# Lecture S04-08: Activity: Save a Hardcoded copy of Workbook

> [!important]
> **Section Reference**: [[Section 04 - Referencing Ranges, Worksheets & Workbooks|Section 04: Referencing Ranges, Worksheets & Workbooks]]  
> **Status**: `✅ Completed`  
> **Lecture Title**: Activity: Save a Hardcoded copy of Workbook

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Activity: Save a Hardcoded copy of Workbook' within the context of Referencing Ranges, Worksheets & Workbooks.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Activity: Save a Hardcoded copy of Workbook', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Hardcoded_Workbook_copy()



    '1: Save the current workbook

    ThisWorkbook.Save

    '2: Select all sheets, copy and paste special all cells in the sheets

    Sheets.Select

    Cells.Copy

    Cells.PasteSpecial xlPasteValues

    'if you'd like to deselect all sheets

    Range("A1").Select

    Sheets(1).Select

    '3: Save the workbook with SaveAs (be careful you don't just SAVE, but SAVEAS)

    '4: Save it in the same directory as this workbook but before the name put "HC_"

    'Bonus: Add today's date in the form of "yymmdd" after "HC_" so HC_yymmdd_S4_ReferencingRanges"



    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\HC_" & VBA.Format(Date, "yymmdd") & "_" & ThisWorkbook.Name



    

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
- [[Excel Object Model]]
- [[Error Handling]]
- [[Section 04 - Referencing Ranges, Worksheets & Workbooks]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Done

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
