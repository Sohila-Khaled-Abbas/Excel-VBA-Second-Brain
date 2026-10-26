---
type: lecture
course: Excel VBA & Macros
section: "[[Section 19 - Creating Customized Functions Formulas]]"
section_number: 19
lecture_number: 3
status: not-started
topics:
  - Creating Customized Functions / Formulas
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-19
---

# Lecture S19-03: Function to Calculate Age (Result in Years & Months)

> [!important]
> **Section Reference**: [[Section 19 - Creating Customized Functions Formulas|Section 19: Creating Customized Functions / Formulas]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Function to Calculate Age (Result in Years & Months)

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Function to Calculate Age (Result in Years & Months)' within the context of Creating Customized Functions / Formulas.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Function to Calculate Age (Result in Years & Months)', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Function TheAge(myDate As Date) As String

    Dim myMonths As Long

    Dim myYears As Long

    myMonths = Month(Date) - Month(myDate)

    myYears = Year(Date) - Year(myDate)

     'to account for negative months

    If myMonths < 0 Then

        myYears = myYears - 1

        myMonths = 12 - VBA.Abs(myMonths)

        

    End If

    TheAge = myYears & " Years & " & myMonths & " Months"

    

End Function
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
- [[Section 19 - Creating Customized Functions Formulas]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
