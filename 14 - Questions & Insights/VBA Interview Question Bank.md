---
title: VBA Interview Question Bank
type: reference
domain: questions_insights
status: evergreen
tags:
  - vba/interview
  - vba/review
  - vba/fundamentals
related:
  - "[[VBA Syntax Cheat Sheet]]"
  - "[[Debugging Cheat Sheet]]"
  - "[[Error Handling]]"
created: 2026-10-26
updated: 2026-10-26
---

# VBA Interview Question Bank

> **Active Recall Strategy**: Cover the answers below, read the question, formulate your answer mentally, then reveal. Review questions you mark 🔴 until they become 🟢.

---

## Category 1 — Core Syntax & Declarations

### Q1: What does `Option Explicit` do and why is it mandatory in professional VBA?
<details>
<summary>Show Answer</summary>

`Option Explicit` forces every variable to be declared with a `Dim`, `Public`, `Private`, or `Static` statement before use. Without it, VBA silently creates a new `Variant` variable for any misspelled name, causing deeply hidden bugs.

**Professional Rule**: Every module must start with `Option Explicit`. Period.
</details>

---

### Q2: What is the difference between `ByVal` and `ByRef` when passing arguments to a Sub or Function?
<details>
<summary>Show Answer</summary>

- **`ByRef`** (default): Passes a reference (memory pointer) to the original variable. Changes inside the procedure **modify the original variable** in the caller.
- **`ByVal`**: Passes a copy. Changes inside the procedure are **discarded on exit** and never affect the caller.

```vba
Sub Test()
    Dim x As Long
    x = 5
    ModifyByRef x  ' x becomes 10 after this
    ModifyByVal x  ' x stays 10 after this
End Sub

Sub ModifyByRef(ByRef n As Long)
    n = n * 2
End Sub

Sub ModifyByVal(ByVal n As Long)
    n = n * 100 ' No effect on x
End Sub
```

**Best Practice**: Always declare `ByVal` unless you *explicitly need* to return values via the argument. Using `ByRef` carelessly introduces side effects that are extremely hard to debug.
</details>

---

### Q3: When would you choose a `Function` procedure over a `Sub` procedure?
<details>
<summary>Show Answer</summary>

- **`Sub`**: Performs actions; returns nothing. Called with `Call Sub()` or just `Sub()`.
- **`Function`**: Performs actions AND returns a value via `Function = result`. Can be called from worksheet cells as a UDF, or from other VBA code where a return value is required.

**Choose `Function` when**: you need to compute a reusable value, process input and return output, or create a custom Excel formula.
</details>

---

## Category 2 — Excel Object Model

### Q4: What is the difference between `ActiveSheet`, `Sheets("Name")`, and `ThisWorkbook.Sheets("Name")`?
<details>
<summary>Show Answer</summary>

| Expression | Refers To | Risk |
| :--- | :--- | :--- |
| `ActiveSheet` | Whichever sheet is visible right now | **HIGH**: Changes if user clicks another tab during execution |
| `Sheets("Name")` | Sheet by name in **whichever workbook is active** | **MEDIUM**: Breaks if another workbook becomes active |
| `ThisWorkbook.Sheets("Name")` | Sheet by name in **the workbook containing the code** | ✅ **SAFE**: Explicit, unambiguous reference |

**Production Rule**: Never use `ActiveSheet` in production code. Always use `ThisWorkbook.Sheets(...)` or a well-scoped `Worksheet` variable.
</details>

---

### Q5: What are the 4 methods to find the last used row in a column, and when would you use each?
<details>
<summary>Show Answer</summary>

```vba
' Method 1: End(xlUp) — Best for most scenarios (reads last cell with data upward)
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Method 2: UsedRange — Unreliable if deleted rows leave phantom used range
lastRow = ws.UsedRange.Rows.Count

' Method 3: CurrentRegion — Good for contiguous blocks around a known anchor
lastRow = ws.Range("A1").CurrentRegion.Rows.Count

' Method 4: Find — Best when some cells in the column might be empty
Dim rng As Range
Set rng = ws.Columns("A").Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
If Not rng Is Nothing Then lastRow = rng.Row
```

**Preferred in production**: Method 1 (`End(xlUp)`) for standard use; Method 4 (`Find`) when column has gaps.
</details>

---

## Category 3 — Error Handling & Debugging

### Q6: What is the difference between `On Error Resume Next` and `On Error GoTo Label`?
<details>
<summary>Show Answer</summary>

- **`On Error Resume Next`**: Silently swallows the error and moves to the next line. Use **only** for targeted, expected errors (e.g., checking `If Err.Number <> 0` immediately after a risky single operation). Leaving it active globally is an anti-pattern.
- **`On Error GoTo Label`**: Jumps to a labeled error handler block. The professional approach for production routines that must log, notify, or clean up on failure.

**Anti-pattern warning**: `On Error Resume Next` at the top of a long procedure masks ALL errors, including catastrophic data corruption bugs. Always scope it tightly and reset with `On Error GoTo 0`.
</details>

---

## Category 4 — Events & UserForms

### Q7: What are the two critical lines every Worksheet_Change handler must include and why?
<details>
<summary>Show Answer</summary>

```vba
Application.EnableEvents = False  ' Before modifying anything on the sheet
' ... your logic ...
Application.EnableEvents = True   ' Always in a cleanup/error block
```

Without `EnableEvents = False`, any cell modification inside the handler triggers the event again → infinite recursion → Excel crash.
</details>

---

## Self-Assessment Tracker

| Question | Last Review | Status | Next Review |
| :------- | :---------: | :----: | :---------: |
| Q1 — Option Explicit | — | ⚪ | — |
| Q2 — ByVal vs ByRef | — | ⚪ | — |
| Q3 — Sub vs Function | — | ⚪ | — |
| Q4 — Sheet References | — | ⚪ | — |
| Q5 — Last Row Methods | — | ⚪ | — |
| Q6 — Error Handling | — | ⚪ | — |
| Q7 — Events Guard | — | ⚪ | — |

> Status Key: 🔴 Needs Work · 🟡 Learning · 🟢 Confident
