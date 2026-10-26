---
type: lecture
course: Excel VBA & Macros
section: "[[Section 08 - Debugging, Error Handling & Procedure Scope]]"
section_number: 8
lecture_number: 8
status: not-started
topics:
  - Debugging, Error Handling & Procedure Scope
related_concepts:
  - '[[VBA Syntax]]'
  - '[[Excel Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-08
---

# Lecture S08-08: Project: Get the Total Number of Formulas (or Comments) Used in a Workbook

> [!important]
> **Section Reference**: [[Section 08 - Debugging, Error Handling & Procedure Scope|Section 08: Debugging, Error Handling & Procedure Scope]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Project: Get the Total Number of Formulas (or Comments) Used in a Workbook

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Project: Get the Total Number of Formulas (or Comments) Used in a Workbook' within the context of Debugging, Error Handling & Procedure Scope.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Project: Get the Total Number of Formulas (or Comments) Used in a Workbook', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Count_WB_Comments_error_handling()
'for tab activity
    Dim Sh As Worksheet
    Dim CounterWS As Double
    Dim CounterGlobal As Double
    
    For Each Sh In ThisWorkbook.Worksheets
        On Error GoTo ErrorHandle
        CounterWS = Sh.Cells.SpecialCells(xlCellTypeComments).Count
Label1:        CounterGlobal = CounterGlobal + CounterWS
ErrorHandle:
    If Err.Number = 1004 Then
        CounterWS = 0
        Resume Label1
    End If
    Next Sh
    MsgBox "There are " & CounterGlobal & " comments in this workbook!"
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
- [[Section 08 - Debugging, Error Handling & Procedure Scope]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
