---
type: concept
domain: VBA
topic: Control Flow
difficulty: beginner
status: mastered
source: Udemy
course_section: "[[Section 10 - Looping in VBA (controlling the flow of code)]]"
related_concepts:
  - "[[For Each]]"
  - "[[Cells]]"
tags:
  - vba
  - loops
  - counter
created: 2026-10-03
updated: 2026-10-03
---

# For Next Loop

> [!important]
> **Summary Statement**: `For...Next` executes a block of code a specified number of times using an integer counter variable.

---

## ⚠️ The Reverse Loop Pattern (Crucial for Deleting Rows)

When deleting rows based on a condition, **you must loop backwards from bottom to top** (`Step -1`). Looping top-to-bottom skips rows because Excel shifts subsequent rows up!

```vba
Sub DeleteZeroValueRows()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long
    Dim r As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' REVERSE LOOP: Bottom to Top
    For r = lastRow To 2 Step -1
        If ws.Cells(r, "D").Value = 0 Or IsEmpty(ws.Cells(r, "D").Value) Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub
```

---

## 🔗 Connected Notes
- [[For Each]], [[Do While]], [[Loops Cheat Sheet]]
