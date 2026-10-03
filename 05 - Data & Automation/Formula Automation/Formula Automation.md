---
type: concept
domain: VBA
topic: Automation
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 13 - Excel Tables, Formulas & Pivot Tables]]"
related_concepts:
  - "[[Range]]"
tags:
  - vba
  - formulas
created: 2026-10-03
updated: 2026-10-03
---

# Formula Automation (`Formula` vs. `FormulaR1C1`)

> [!important]
> **Summary Statement**: VBA can dynamically inject, update, and hardcode formulas across large ranges using standard or R1C1 notation.

---

## 💻 Wrapping Existing Formulas with `IFERROR`
```vba
Sub WrapSelectionWithIFERROR()
    Dim cell As Range
    For Each cell In Selection.SpecialCells(xlCellTypeFormulas)
        cell.Formula = "=IFERROR(" & Mid(cell.Formula, 2) & ", 0)"
    Next cell
End Sub
```

---

## 🔗 Connected Notes
- [[Range]], [[Excel Tables]]
