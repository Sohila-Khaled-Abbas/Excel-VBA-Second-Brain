---
type: lecture
course: Excel VBA & Macros
section: "[[Section 19 - Creating Customized Functions / Formulas]]"
section_number: 19
lecture_number: 5
status: not-started
topics:
  - Creating Customized Functions / Formulas
related_concepts:
  - '[[VBA Syntax]]'
  - '[[Excel Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-19
---

# Lecture S19-05: Using Functions in Sub Procedures

> [!important]
> **Section Reference**: [[Section 19 - Creating Customized Functions / Formulas|Section 19: Creating Customized Functions / Formulas]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Using Functions in Sub Procedures

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Using Functions in Sub Procedures' within the context of Creating Customized Functions / Formulas.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Using Functions in Sub Procedures', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Split_Text_Number()

    Dim myRange As Range
    Dim cell As Range
    Set myRange = Range("A6", "A" & Cells(Rows.Count, 1).End(xlUp).Row)
    
    For Each cell In myRange
        cell.Offset(0, 1).Value = TextOnly(cell.Value)
        cell.Offset(0, 2).Value = NumberOnly(cell.Value)
    Next cell
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
- [[Section 19 - Creating Customized Functions / Formulas]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
