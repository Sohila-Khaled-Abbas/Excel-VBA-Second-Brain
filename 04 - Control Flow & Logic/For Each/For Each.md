---
type: concept
domain: VBA
topic: Control Flow
difficulty: beginner
status: mastered
source: Udemy
course_section: "[[Section 06 - Looping through Collections & Making Decisions]]"
related_concepts:
  - "[[Collections]]"
  - "[[For Next]]"
  - "[[Worksheet]]"
tags:
  - vba
  - loops
  - collections
created: 2026-10-03
updated: 2026-10-03
---

# For Each Loop

> [!important]
> **Summary Statement**: `For Each...Next` iterates through every member of an object collection (e.g., sheets, cells, workbooks) without requiring a manual index counter.

---

## 💻 Core Use Case: Looping All Worksheets

```vba
Sub UnhideAllWorksheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub
```

---

## ⚖️ When to Use `For Each` vs. `For Next`
- Use **`For Each`** when visiting every element in a collection where the order does not matter.
- Use **`For Next`** when you need an integer counter, need to step backwards, or are deleting rows.

---

## 🔗 Connected Notes
- [[For Next]], [[Collections]], [[Loops Cheat Sheet]]
