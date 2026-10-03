---
type: cheat-sheet
domain: VBA
topic: Loops
tags:
  - cheat-sheet
  - loops
  - control-flow
created: 2026-10-03
updated: 2026-10-03
---

# 🔄 Loops Cheat Sheet

| Loop Type | Best Used When | Syntax Template |
| :--- | :--- | :--- |
| **`For Each...Next`** | Iterating through all objects in a collection | `For Each ws In ThisWorkbook.Worksheets`<br>`Next ws` |
| **`For...Next` (Forward)** | Counting from fixed start to end | `For i = 1 To lastRow`<br>`Next i` |
| **`For...Next` (Reverse)** | **Deleting rows** or moving bottom-up | `For i = lastRow To 2 Step -1`<br>`Next i` |
| **`Do While...Loop`** | Looping while a condition remains `True` | `Do While Cells(r, 1).Value <> ""` <br>`    r = r + 1`<br>`Loop` |
| **`Do Until...Loop`** | Looping until a condition becomes `True` | `Do Until Cells(r, 1).Value = ""` <br>`    r = r + 1`<br>`Loop` |
