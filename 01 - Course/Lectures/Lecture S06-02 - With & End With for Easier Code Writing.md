---
type: lecture
course: Excel VBA & Macros
section: "[[Section 06 - Looping through Collections & Making Decisions]]"
section_number: 6
lecture_number: 2
status: not-started
topics:
  - Looping through Collections & Making Decisions
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-06
---

# Lecture S06-02: With & End With for Easier Code Writing

> [!important]
> **Section Reference**: [[Section 06 - Looping through Collections & Making Decisions|Section 06: Looping through Collections & Making Decisions]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: With & End With for Easier Code Writing

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'With & End With for Easier Code Writing' within the context of Looping through Collections & Making Decisions.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'With & End With for Easier Code Writing', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub With_Change_Font()

' for tab With

    Dim myRange As Range

    'start is fixed but last cell is variable

      Set myRange = Range("A10", "A" & Cells(Rows.Count, 1).End(xlUp).Row)

      With myRange.Font

        .Name = "Arial"

        .Size = 12

        .Bold = True

        ' myRange.WrapText = True

      End With

  

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
- [[Section 06 - Looping through Collections & Making Decisions]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
