---
type: concept
domain: VBA
topic: Control Flow
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 10 - Looping in VBA (controlling the flow of code)]]"
related_concepts:
  - "[[Do Until]]"
  - "[[For Next]]"
tags:
  - vba
  - loops
  - conditionals
created: 2026-10-03
updated: 2026-10-03
---

# Do While Loop

> [!important]
> **Summary Statement**: `Do While` continues looping as long as a specified condition evaluates to `True`.

---

## 💻 Syntax & Safety Guardrails
```vba
Sub ReadUntilBlank()
    Dim r As Long: r = 2
    
    ' Loop until first empty cell is reached
    Do While ActiveSheet.Cells(r, "A").Value <> ""
        ' Process cell
        r = r + 1
        
        ' Safety guardrail against infinite loop
        If r > 1000000 Then Exit Do
    Loop
End Sub
```

---

## 🔗 Connected Notes
- [[Do Until]], [[Loops Cheat Sheet]]
