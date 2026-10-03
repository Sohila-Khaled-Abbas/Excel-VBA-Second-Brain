---
type: concept
domain: VBA
topic: Performance
difficulty: intermediate
status: mastered
source: Udemy
course_section: "[[Section 08 - Debugging, Error Handling & Procedure Scope]]"
related_concepts:
  - "[[Application]]"
  - "[[Arrays]]"
tags:
  - vba
  - performance
  - optimization
created: 2026-10-03
updated: 2026-10-03
---

# Performance Optimization

> [!important]
> **Summary Statement**: Professional VBA optimization focuses on suppressing unnecessary Excel UI updates, minimizing COM interface roundtrips, and performing heavy calculations in memory.

---

## 🚀 The Golden Rules of VBA Performance
1. **Never Select or Activate**: Direct object access is 10x faster (`ws.Range("A1").Value = 5` instead of `ws.Activate; Range("A1").Select; Selection.Value = 5`).
2. **Turn Off ScreenUpdating**: Stop the graphics card from rendering spreadsheet cell changes during loops.
3. **Turn Off Automatic Calculation**: Prevent recalculating 100,000 worksheet formulas every time one cell value changes.
4. **Use In-Memory Arrays**: Load data into RAM, transform, and dump back in a single operation.

---

## 🔗 Connected Notes
- [[Application]], [[Arrays]]
