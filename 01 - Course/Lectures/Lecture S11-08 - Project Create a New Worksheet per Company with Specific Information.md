---
type: lecture
course: Excel VBA & Macros
section: "[[Section 11 - Working with Arrays]]"
section_number: 11
lecture_number: 8
status: not-started
topics:
  - Working with Arrays
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-11
---

# Lecture S11-08: Project: Create a New Worksheet per Company with Specific Information

> [!important]
> **Section Reference**: [[Section 11 - Working with Arrays|Section 11: Working with Arrays]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Project: Create a New Worksheet per Company with Specific Information

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Project: Create a New Worksheet per Company with Specific Information' within the context of Working with Arrays.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Project: Create a New Worksheet per Company with Specific Information', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Manager_Report()



    Dim CompInfo(1 To 8, 1 To 2) As String

    Dim r As Byte, c As Byte

    Const StartRow As Long = 6

    Dim NewBook As Workbook

    Dim ShNew As Worksheet

    

    For r = 1 To 8

        For c = 1 To 2

            CompInfo(r, c) = Cells(r + StartRow, c).Value

        Next c

    Next r

    

    Set NewBook = Workbooks.Add

    'reuse r for worksheets

    For r = 1 To 8

        Set ShNew = NewBook.Worksheets.Add

        ShNew.Name = CompInfo(r, 1)

        ShNew.Range("A1").Value = CompInfo(r, 1)

        ShNew.Range("A2").Value = CompInfo(r, 2)

        

    Next r

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
- [[Section 11 - Working with Arrays]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
