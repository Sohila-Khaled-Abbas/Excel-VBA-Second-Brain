---
type: concept
domain: VBA
topic: Control Flow
difficulty: beginner
status: mastered
source: Udemy
course_section: "[[Section 06 - Looping through Collections & Making Decisions]]"
related_concepts:
  - "[[If Statements]]"
tags:
  - vba
  - conditionals
  - clean-code
created: 2026-10-03
updated: 2026-10-03
---

# Select Case

> [!important]
> **Summary Statement**: `Select Case` evaluates a single expression against a series of possible values, providing cleaner, faster, and more maintainable code than nested `If...ElseIf` chains.

---

## 💻 Syntax Pattern
```vba
Select Case region
    Case "North", "Northwest"
        taxRate = 0.05
    Case "South"
        taxRate = 0.04
    Case "International"
        taxRate = 0.00
    Case Else
        taxRate = 0.06
End Select
```

---

## 🔗 Connected Notes
- [[If Statements]]
