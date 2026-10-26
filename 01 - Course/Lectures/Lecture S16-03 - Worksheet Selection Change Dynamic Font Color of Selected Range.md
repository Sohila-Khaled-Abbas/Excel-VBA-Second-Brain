---
type: lecture
course: Excel VBA & Macros
section: "[[Section 16 - Worksheet and Workbook Events]]"
section_number: 16
lecture_number: 3
status: not-started
topics:
  - Worksheet and Workbook Events
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-16
---

# Lecture S16-03: Worksheet Selection Change: Dynamic Font Color of Selected Range

> [!important]
> **Section Reference**: [[Section 16 - Worksheet and Workbook Events|Section 16: Worksheet and Workbook Events]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Worksheet Selection Change: Dynamic Font Color of Selected Range

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Worksheet Selection Change: Dynamic Font Color of Selected Range' within the context of Worksheet and Workbook Events.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Worksheet Selection Change: Dynamic Font Color of Selected Range', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Cells.Font.Color = VBA.ColorConstants.vbBlack

    Target.Font.Color = VBA.ColorConstants.vbBlue

    

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
- [[Section 16 - Worksheet and Workbook Events]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
