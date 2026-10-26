---
type: course-section
section_number: 6
section_title: Looping through Collections & Making Decisions
lecture_count: 10
status: not-started
tags:
  - course-section
  - vba
  - section-06
---

# Section 06: Looping through Collections & Making Decisions

> [!important]
> **Curriculum Pillar**: Section 6 contains **10 lectures**.  
> **Status**: `NOT-STARTED`  
> **Course Practice Files**: `SectionFiles/S6_Collections_With_IF.xlsm`, `S6_Collections_With_IF_Start.xlsx`

---

## 🎯 Section Objectives & Automation Capabilities
- Master collections as built-in object containers (`Worksheets`, `Workbooks`, `Cells`).
- Streamline repetitive object references and boost performance using `With...End With` blocks.
- Iterate over entire collections in a single step using `For Each...Next`.
- Direct program flow with conditional branching: single-line `If`, multi-line `If-Then-ElseIf-Else`, and nested conditions.
- Replace complex, unwieldy If-structures with clean `Select Case` evaluation blocks.
- Understand controlled jumps with `GoTo` (specifically for error exit pathways).

---

## 🧠 Architectural Takeaways & Mental Models
- **`With...End With` Construct**: Eliminates redundant object resolution. Each dot inside the block refers to the parent object. Not only cleaner to read, but faster because Excel resolves the object pointer once.
- **`For Each...Next` on Collections**: The most readable and safest iteration construct. No index tracking required:
  ```vba
  Dim ws As Worksheet
  For Each ws In ThisWorkbook.Worksheets
      ws.Protect Password:="Secret"
  Next ws
  ```
- **`Select Case`**: Far superior to cascading `ElseIf` ladders when testing a single variable against multiple outcomes, discrete lists, or numeric ranges (`Case 1 To 10`, `Case Is > 100`, `Case "A", "B"`).
- **GoTo Discipline**: Never use `GoTo` for standard program loops ("spaghetti code"). Reserve `GoTo` strictly for clean error handling exits (`On Error GoTo ErrorHandler`).

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- Forgetting the leading dot inside a `With` block (e.g. writing `Range("A1")` instead of `.Range("A1")`) breaks encapsulation and points back to the `ActiveSheet`!
- Modifying a collection (such as deleting sheets or rows) inside a `For Each` loop can cause skipped elements or runtime crashes. Use backwards `For i = count To 1 Step -1` instead.

---

## 📚 Lecture Index

- [ ] [[Lecture S06-01 - Overview Why Collections are a VBA Essential|Lecture 01: Overview: Why Collections are a VBA Essential]]
- [ ] [[Lecture S06-02 - With & End With for Easier Code Writing|Lecture 02: With & End With for Easier Code Writing]]
- [ ] [[Lecture S06-03 - For Each to Loop Through Collections (sheets, ranges etc.) in one go|Lecture 03: For Each to Loop Through Collections (sheets, ranges etc.) in one go]]
- [ ] [[Lecture S06-04 - IF Then (Else, ElseIF) for Conditional Outcomes|Lecture 04: IF Then (Else, ElseIF) for Conditional Outcomes]]
- [ ] [[Lecture S06-05 - Select Case as Alternative for Many IF Statements|Lecture 05: Select Case as Alternative for Many IF Statements]]
- [ ] [[Lecture S06-06 - Goto Statement to Change Program Flow|Lecture 06: Goto Statement to Change Program Flow]]
- [ ] [[Lecture S06-07 - Activity Get the Number of Formulas on the worksheet|Lecture 07: Activity: Get the Number of Formulas on the worksheet]]
- [ ] [[Lecture S06-08 - Project Get the Number of Formulas on the Worksheet|Lecture 08: Project: Get the Number of Formulas on the Worksheet]]
- [ ] [[Lecture S06-09 - Bonus Unhide All Sheets in One Go|Lecture 09: Bonus: Unhide All Sheets in One Go]]
- [ ] [[Lecture S06-10 - Key Takeaways for Collections & Decision Making in VBA|Lecture 10: Key Takeaways for Collections & Decision Making in VBA]]

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
