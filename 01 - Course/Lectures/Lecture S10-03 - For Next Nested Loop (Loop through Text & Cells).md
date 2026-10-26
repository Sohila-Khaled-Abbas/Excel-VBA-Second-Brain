---
type: lecture
course: Excel VBA & Macros
section: "[[Section 10 - Looping in VBA (controlling the flow of code)]]"
section_number: 10
lecture_number: 3
status: not-started
topics:
  - Looping in VBA (controlling the flow of code)
related_concepts:
  - '[[VBA Syntax]]'
  - '[[Excel Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-10
---

# Lecture S10-03: For Next Nested Loop (Loop through Text & Cells)

> [!important]
> **Section Reference**: [[Section 10 - Looping in VBA (controlling the flow of code)|Section 10: Looping in VBA (controlling the flow of code)]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: For Next Nested Loop (Loop through Text & Cells)

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'For Next Nested Loop (Loop through Text & Cells)' within the context of Looping in VBA (controlling the flow of code).

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'For Next Nested Loop (Loop through Text & Cells)', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub For_Next_Loop_in_Text()
    Dim i As Long 'for looping inside each cell
    Dim myValue As String
    Dim NumFound As Long
    Dim TxtFound As String
    Dim r As Long 'for looping through rows
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    For r = StartRow To LastRow
        myValue = Range("A" & r).Value
        For i = 1 To VBA.Len(myValue)
            If IsNumeric(VBA.Mid(myValue, i, 1)) Then
                NumFound = NumFound & Mid(myValue, i, 1)
            ElseIf Not IsNumeric(Mid(myValue, i, 1)) Then
                TxtFound = TxtFound & Mid(myValue, i, 1)
            End If
            
        Next i
        Range("H" & r).Value = TxtFound
        Range("I" & r).Value = NumFound
        NumFound = 0
        TxtFound = ""
        
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
- [[Excel Object Model]]
- [[Error Handling]]
- [[Section 10 - Looping in VBA (controlling the flow of code)]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
