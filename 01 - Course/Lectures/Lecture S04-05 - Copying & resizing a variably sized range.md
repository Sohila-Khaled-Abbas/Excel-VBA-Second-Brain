---
type: lecture
course: Excel VBA & Macros
section: "[[Section 04 - Referencing Ranges, Worksheets & Workbooks]]"
section_number: 4
lecture_number: 5
status: completed
topics:
  - Referencing Ranges, Worksheets & Workbooks
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-04
---

# Lecture S04-05: Copying & resizing a variably sized range

> [!important]
> **Section Reference**: [[Section 04 - Referencing Ranges, Worksheets & Workbooks|Section 04: Referencing Ranges, Worksheets & Workbooks]]  
> **Status**: `✅ Completed`  
> **Lecture Title**: Copying & resizing a variably sized range

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Copying & resizing a variably sized range' within the context of Referencing Ranges, Worksheets & Workbooks.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Copying & resizing a variably sized range', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Copy_Resize_Ranges()



    'this macro shows you how to copy ranges and resize them if needed

    'clear before copy

    Range("J4:O16,J20:O32,A20:O32").Clear



     '-------------------------

    'Copy Method

    '-------------------------

    Range("A4:E10").Copy Range("J4")

    Range("A4").CurrentRegion.Copy Range("J4")



    '-------------------------

    'PasteSpecial Method

    '-------------------------

    Range("A4").CurrentRegion.Copy

    Range("J20").PasteSpecial xlPasteValuesAndNumberFormats

    Range("J20").PasteSpecial xlPasteComments



    '-------------------------

    'Resize Property with Copy Method

    '-------------------------

    Range("A4").CurrentRegion.Offset(1, 0).Resize(Range("A4"). _

    CurrentRegion.Rows.Count - 1).Copy

    Range("A20").PasteSpecial xlPasteValuesAndNumberFormats

    Application.CutCopyMode = False

    Range("A1").Select

    



    

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
- [[Section 04 - Referencing Ranges, Worksheets & Workbooks]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Done

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
