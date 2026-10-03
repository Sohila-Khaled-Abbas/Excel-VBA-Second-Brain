---
type: concept
domain: VBA
topic: Syntax & Clean Code
difficulty: beginner
status: mastered
source: Udemy
course_section: "[[Section 06 - Looping through Collections & Making Decisions]]"
related_concepts:
  - "[[Range]]"
  - "[[Performance Optimization]]"
tags:
  - vba
  - syntax
  - with-statement
created: 2026-10-03
updated: 2026-10-03
---

# With Statement (`With...End With`)

> [!important]
> **Summary Statement**: `With...End With` executes a series of statements on a single object without re-evaluating or retyping the object reference, improving code readability and execution speed.

---

## 💻 Syntax & Comparison

```vba
' WITHOUT With Statement (Repetitive & Slower):
ws.Range("A1").Font.Bold = True
ws.Range("A1").Font.Size = 12
ws.Range("A1").Font.Color = RGB(0, 0, 0)
ws.Range("A1").Interior.Color = RGB(220, 230, 240)

' WITH With Statement (Clean, Fast, Concise):
With ws.Range("A1")
    With .Font
        .Bold = True
        .Size = 12
        .Color = RGB(0, 0, 0)
    End With
    .Interior.Color = RGB(220, 230, 240)
End With
```
