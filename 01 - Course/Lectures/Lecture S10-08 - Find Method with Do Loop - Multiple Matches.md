---
type: lecture
course: Excel VBA & Macros
section: "[[Section 10 - Looping in VBA (controlling the flow of code)]]"
section_number: 10
lecture_number: 8
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

# Lecture S10-08: Find Method with Do Loop - Multiple Matches

> [!important]
> **Section Reference**: [[Section 10 - Looping in VBA (controlling the flow of code)|Section 10: Looping in VBA (controlling the flow of code)]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Find Method with Do Loop - Multiple Matches

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Find Method with Do Loop - Multiple Matches' within the context of Looping in VBA (controlling the flow of code).

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Find Method with Do Loop - Multiple Matches', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Many_Finds()
    Dim CompId As Range
    Dim i As Byte
    Dim FirstMatch As Variant
    
    Range("D3:D6").ClearContents
    i = 3
    Dim Start
    Start = VBA.Timer
    
    Set CompId = Range("A:A").Find(what:=Range("B3").Value, _
    LookIn:=xlValues, lookat:=xlWhole)
    If Not CompId Is Nothing Then
        Range("D" & i).Value = CompId.Offset(, 4).Value
        FirstMatch = CompId.Address
        Do
            Set CompId = Range("A:A").FindNext(CompId)
            If CompId.Address = FirstMatch Then Exit Do
            i = i + 1
            Range("D" & i).Value = CompId.Offset(, 4).Value
        Loop
    Else
    MsgBox "Company not found!"
    End If
    Debug.Print Round(Timer - Start, 3)
    
    
    'Application.Speech.Speak "Well Done. " & i - 2 & " matches were found."
    
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
