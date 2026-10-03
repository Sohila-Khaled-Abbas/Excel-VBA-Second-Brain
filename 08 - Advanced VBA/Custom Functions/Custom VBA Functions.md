---
type: concept
domain: VBA
topic: Functions
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 19 - Creating Customized Functions Formulas]]"
related_concepts:
  - "[[Procedures & Functions]]"
tags:
  - vba
  - udfs
  - formulas
created: 2026-10-03
updated: 2026-10-03
---

# Custom VBA Functions (User-Defined Functions / UDFs)

> [!important]
> **Summary Statement**: UDFs are custom calculation functions written in standard VBA modules that can be called directly inside Excel worksheet formulas just like `SUM()` or `VLOOKUP()`.

---

## 💻 Classic UDF: Sum by Cell Color
```vba
Function SumByColor(sumRange As Range, colorCriteriaCell As Range) As Double
    Dim cell As Range
    Dim targetColor As Long
    Dim total As Double
    
    targetColor = colorCriteriaCell.Interior.Color
    
    For Each cell In sumRange
        If cell.Interior.Color = targetColor Then
            If IsNumeric(cell.Value) Then
                total = total + cell.Value
            End If
        End If
    Next cell
    
    SumByColor = total
End Function
```

---

## 🔗 Connected Notes
- [[Procedures & Functions]], [[Formula Automation]]
