---
type: lecture
course: Excel VBA & Macros
section: "[[Section 05 - Working with Variables]]"
section_number: 5
lecture_number: 6
status: not-started
topics:
  - Working with Variables
related_concepts:
  - '[[VBA Syntax]]'
  - '[[Excel Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-05
---

# Lecture S05-06: Activity: Revisit Hard-coded Workbook Project & Use Variable Assignment

> [!important]
> **Section Reference**: [[Section 05 - Working with Variables|Section 05: Working with Variables]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Activity: Revisit Hard-coded Workbook Project & Use Variable Assignment

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Activity: Revisit Hard-coded Workbook Project & Use Variable Assignment' within the context of Working with Variables.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Activity: Revisit Hard-coded Workbook Project & Use Variable Assignment', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Hardcoded_Workbook_copy()

    'Use appropriate variables in this example
    
    'this macro should first_
    '1: Save the current workbook
    '2: Select all sheets, copy and paste special all cells in the sheets
    '3: Save the workbook with SaveAs (be careful you don't just SAVE, but SAVEAS)
    '4: Save it in the same directory as this workbook but before the name put "HC_"
    'Bonus: Add today's date in the form of "yymmdd" after "HC_" so HC_yymmdd_S4_ReferencingRanges"

    Dim FilePath As String
    Dim HCName As String
    
    FilePath = ThisWorkbook.Path & "\"
    HCName = "HC_" & VBA.Format(Date, "yymmdd") & "_" & ThisWorkbook.Name
    ThisWorkbook.Save
    Sheets.Select
    Cells.Copy
    Cells.PasteSpecial xlPasteValues
    
    Range("A1").Select
    Sheets(1).Select
    ThisWorkbook.SaveAs Filename:=FilePath & HCName
    
    
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
- [[Section 05 - Working with Variables]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
