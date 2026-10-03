---
type: insights
topic: Breakthroughs
tags:
  - breakthroughs
  - mental-models
created: 2026-10-03
updated: 2026-10-03
---

# 💡 Things I Finally Understood (Aha! Moments)

1. **Why the Macro Recorder Code is Terrible**:
   - The recorder writes code for the *user interface* (recording every mouse click, scroll, and selection).
   - Real VBA talks directly to the *object model* in memory without ever needing to select anything. Eliminating `.Select` turns minutes of sluggish macro execution into fractions of a second.
2. **The `Set` Keyword**:
   - Primitive data types (`Long`, `Double`, `String`) hold actual values. They are assigned with `=`.
   - Objects (`Workbook`, `Worksheet`, `Range`) are complex memory structures. Assigning an object creates a memory pointer, which VBA explicitly requires the `Set` keyword to initialize.
3. **The Semicolon CSV Problem in Analytics**:
   - Different regional settings across international teams interpret commas as decimal separators. Writing custom file exports via `Print #` guarantees deterministic data exchange regardless of user regional settings.
