---
type: lecture
course: Excel VBA & Macros
section: "[[Section 07 - Useful Built-in Functions]]"
section_number: 7
lecture_number: 5
status: not-started
topics:
  - Useful Built-in Functions
related_concepts:
  - '[[VBA Syntax]]'
  - '[[Excel Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-07
---

# Lecture S07-05: Input Box that can Select Ranges (Excel InputBox Method)

> [!important]
> **Section Reference**: [[Section 07 - Useful Built-in Functions|Section 07: Useful Built-in Functions]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Input Box that can Select Ranges (Excel InputBox Method)

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Input Box that can Select Ranges (Excel InputBox Method)' within the context of Useful Built-in Functions.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Input Box that can Select Ranges (Excel InputBox Method)', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Excel_InputBox_Range()
'for tab InputBox_Excel
    Dim myRange As Range
    Dim cellBlank As Long, cellNum As Long, cellOther As Long
    
    On Error GoTo Leave 'in case the user clicks on cancel
    Set myRange = Application.InputBox("Count the number of cells", "Checking Cells...", , , , , , 8)
    cellBlank = Excel.WorksheetFunction.CountBlank(myRange)
    cellNum = Excel.WorksheetFunction.Count(myRange)
    cellOther = Excel.WorksheetFunction.CountA(myRange) - cellNum
    'message box to inform
    MsgBox cellBlank & " cells are blank." & vbNewLine _
    & cellNum & " cells have numbers." & vbNewLine _
    & cellOther & " cells are non-numeric."
Leave:
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
- [[Section 07 - Useful Built-in Functions]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
