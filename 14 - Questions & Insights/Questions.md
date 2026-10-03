---
type: active-recall
topic: VBA Conceptual Understanding
tags:
  - active-recall
  - flashcards
  - review
created: 2026-10-03
updated: 2026-10-03
---

# ❓ Active Recall & Spaced Review Questions

> [!tip]
> **Learning Rule**: Attempt to answer each question out loud or in writing before expanding the collapsible details block.

---

### Q1: What is the fundamental architectural difference between `Range` and `Worksheet`?
<details>
<summary><b>Reveal Answer</b></summary>

A `Worksheet` is a 2D sheet container object that owns cells, tables, charts, and properties like tab color and visibility. A `Range` is a pointer to one or more contiguous or non-contiguous cells inside a worksheet. A Range cannot exist in a vacuum; it always belongs to a parent Worksheet.
</details>

---

### Q2: Why MUST you loop backwards (`Step -1`) when deleting rows in a `For Next` loop?
<details>
<summary><b>Reveal Answer</b></summary>

When Excel deletes a row (e.g., row 5), all subsequent rows immediately shift up by 1 (row 6 becomes row 5). If your counter increments forward from 5 to 6, the new row 5 is completely skipped without evaluation. Looping backwards from bottom to top ensures that shifting rows below does not affect rows that haven't been evaluated yet.
</details>

---

### Q3: What is the difference between a `Sub` and a `Function`?
<details>
<summary><b>Reveal Answer</b></summary>

A `Sub` performs actions (formatting, creating files, popping dialogs) and cannot return a value. A `Function` calculates and returns a value assigned to its own name. When called from a cell as a User-Defined Function (UDF), a `Function` is read-only and cannot alter worksheet formats or modify other cells.
</details>

---

### Q4: Why does copying a range into a `Variant` array execute 100x faster than cell-by-cell loops?
<details>
<summary><b>Reveal Answer</b></summary>

Every time VBA reads or writes directly to `Cells(r, c)`, it must cross the COM (Component Object Model) interface bridge between the VBA engine and Excel's C++ core. Crossing this bridge 50,000 times introduces massive latency. Transferring the entire range into a `Variant` array crosses the bridge exactly once into RAM, where loop operations run at hardware CPU speed.
</details>

---

### Q5: What is the critical danger of forgetting `Application.EnableEvents = False` in a `Worksheet_Change` event?
<details>
<summary><b>Reveal Answer</b></summary>

If the code inside `Worksheet_Change` modifies any cell value on the worksheet, that modification immediately triggers another `Worksheet_Change` event. This causes an infinite recursive loop, freezing Excel until the call stack overflows and Excel crashes.
</details>

---

### Q6: What is the difference between `ThisWorkbook` and `ActiveWorkbook`?
<details>
<summary><b>Reveal Answer</b></summary>

`ThisWorkbook` refers to the workbook where the VBA code itself is stored and running from. `ActiveWorkbook` refers to whichever workbook currently has window focus in the Excel application. If your macro opens a customer file, `ActiveWorkbook` changes to the customer file, while `ThisWorkbook` remains pointed to your automation tool.
</details>
