---
type: concept
domain: VBA
topic: Object Model
difficulty: beginner
status: mastered
source: Udemy
course_section: "[[Section 04 - Referencing Ranges, Worksheets & Workbooks]]"
related_concepts:
  - "[[Range]]"
  - "[[For Next]]"
tags:
  - vba
  - cells
  - coordinate
created: 2026-10-03
updated: 2026-10-03
---

# Cells Property

> [!important]
> **Summary Statement**: `Cells(row, column)` references an individual cell using numeric coordinate indexes, making it ideal for dynamic row-and-column looping in `For Next` constructs.

---

## ⚖️ `Range` vs. `Cells`

| Syntax | Example | Best For |
| :--- | :--- | :--- |
| **`Range`** | `ws.Range("A1:D10")` | Multi-cell blocks, fixed column letters, readable addresses |
| **`Cells`** | `ws.Cells(r, c)` | Dynamic programmatic loops where rows and columns are integers |
| **Hybrid** | `ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 5))` | Dynamically defining a multi-column rectangular block |

---

## 🔗 Connected Notes
- [[Range]], [[For Next]], [[Loops Cheat Sheet]]
