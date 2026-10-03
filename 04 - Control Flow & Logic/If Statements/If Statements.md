---
type: concept
domain: VBA
topic: Control Flow
difficulty: beginner
status: mastered
source: Udemy
course_section: "[[Section 06 - Looping through Collections & Making Decisions]]"
related_concepts:
  - "[[Select Case]]"
tags:
  - vba
  - conditionals
  - logic
created: 2026-10-03
updated: 2026-10-03
---

# If Statements (If...Then...Else)

> [!important]
> **Summary Statement**: `If...Then...Else` conditionally branches code execution based on boolean truth evaluations.

---

## 💻 Syntax Pattern
```vba
If score >= 90 Then
    grade = "A"
ElseIf score >= 80 Then
    grade = "B"
ElseIf score >= 70 Then
    grade = "C"
Else
    grade = "F"
End If
```

---

## 🔗 Connected Notes
- [[Select Case]]
