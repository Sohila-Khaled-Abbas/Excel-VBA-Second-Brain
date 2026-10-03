---
type: insights
topic: Error Log
tags:
  - mistakes
  - debugging
created: 2026-10-03
updated: 2026-10-03
---

# ⚠️ Mistakes & Anti-Patterns Log

| Anti-Pattern | Why It Fails | Correction |
| :--- | :--- | :--- |
| `Dim x, y, z As Long` | Only `z` is a `Long`; `x` and `y` silently default to `Variant` | `Dim x As Long, y As Long, z As Long` |
| `Dim r As Integer` | Throws `Run-time error '6': Overflow` once rows exceed 32,767 | Always use `Dim r As Long` |
| `ws.Range("A1:C10").Select` | Drastically slows execution and breaks in background tasks | `ws.Range("A1:C10").Font.Bold = True` directly |
| Forgetting `Application.ScreenUpdating = True` on error | Leaves Excel screen permanently frozen if an error aborts the macro | Always wrap macros in `On Error GoTo ErrorHandler` with a `CleanExit` restoration block |
