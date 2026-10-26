---
type: lecture
course: Excel VBA & Macros
section: "[[Section 17 - Working with Userforms & ActiveX Controls]]"
section_number: 17
lecture_number: 6
status: not-started
topics:
  - Working with Userforms & ActiveX Controls
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-17
---

# Lecture S17-06: ActiveX to Show or Hide Other ActiveX controls (eg. Actual & Budget reporting)

> [!important]
> **Section Reference**: [[Section 17 - Working with Userforms & ActiveX Controls|Section 17: Working with Userforms & ActiveX Controls]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: ActiveX to Show or Hide Other ActiveX controls (eg. Actual & Budget reporting)

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'ActiveX to Show or Hide Other ActiveX controls (eg. Actual & Budget reporting)' within the context of Working with Userforms & ActiveX Controls.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'ActiveX to Show or Hide Other ActiveX controls (eg. Actual & Budget reporting)', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Private Sub CBYear_Change()

    If Val(CBYear.Value) = Year(Date) Then

        Me.OPBudget.Visible = True

        Me.OPFC.Visible = True

    Else

        Me.OPBudget.Visible = False

        Me.OPFC.Visible = False

        shData.Range("H4").Value = False

        shData.Range("H5").Value = False

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
- [[Section 17 - Working with Userforms & ActiveX Controls]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
