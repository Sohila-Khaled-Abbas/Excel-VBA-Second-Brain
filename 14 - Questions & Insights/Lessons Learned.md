---
type: insights
topic: Architecture Principles
tags:
  - architecture
  - principles
created: 2026-10-03
updated: 2026-10-03
---

# 🎓 Core Engineering Lessons Learned

1. **Defensive Programming First**: Never assume a sheet, column, or file exists. Check first (`Dir()`, `SheetExists()`, `Intersect()`).
2. **Decouple Logic from UI**: Keep UserForm modules thin. UserForms should collect inputs, validate them, and call standard procedural modules to do the work.
3. **Always Restore Environment**: Any script that modifies `ScreenUpdating`, `Calculation`, or `EnableEvents` must guarantee their restoration via an exit trap.
