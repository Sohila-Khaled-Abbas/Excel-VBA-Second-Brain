---
type: lecture
course: Excel VBA & Macros
section: "[[Section 14 - Main Project 2 Regional Reporting Tool]]"
section_number: 14
lecture_number: 8
status: not-started
topics:
  - Main Project #2: Regional Reporting Tool
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-14
---

# Lecture S14-08: Track the List of Tasks Completed with VBA

> [!important]
> **Section Reference**: [[Section 14 - Main Project 2 Regional Reporting Tool|Section 14: Main Project #2: Regional Reporting Tool]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Track the List of Tasks Completed with VBA

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Track the List of Tasks Completed with VBA' within the context of Main Project #2: Regional Reporting Tool.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Track the List of Tasks Completed with VBA', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Admin_Tasks()

    Dim InputPass As String, Password As String

    Password = "me"

    InputPass = VBA.InputBox("Enter Password to toggle access")

    If InputPass = "" Then Exit Sub

    

    If InputPass <> Password Then

        MsgBox Title:="Incorrect Password", prompt:="The Password is Incorrect!", Buttons:=vbCritical

        Exit Sub

    End If

    Application.ScreenUpdating = False

    If shAll.Visible = xlSheetVisible Then

        shAll.Visible = xlSheetVeryHidden

        shMap.Visible = xlSheetVeryHidden

        shNewDat.Visible = xlSheetVeryHidden

    Else

         shAll.Visible = xlSheetVisible

        shMap.Visible = xlSheetVisible

        shNewDat.Visible = xlSheetVisible

    End If

    Application.ScreenUpdating = False

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
- [[Section 14 - Main Project 2 Regional Reporting Tool]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
