---
type: course-section
section_number: 10
section_title: Looping in VBA (controlling the flow of code)
lecture_count: 12
status: not-started
tags:
  - course-section
  - vba
  - section-10
---

# Section 10: Looping in VBA (controlling the flow of code)

> [!important]
> **Curriculum Pillar**: Section 10 contains **12 lectures**.  
> **Status**: `NOT-STARTED`  
> **Course Practice Files**: `SectionFiles/S10_Looping.xlsm`, `S10_Looping_Start.xlsx`

---

## 🎯 Section Objectives & Automation Capabilities
- Master counter-driven looping constructs: `For...Next`, `Step` increments, and reverse loops (`Step -1`).
- Solve the classic "Row Deletion Trap" using reverse loops.
- Implement conditional loops: `Do While...Loop` and `Do Until...Loop` (evaluating at start vs end of loop).
- Harness `Range.Find` and `FindNext` for lightning-fast search without iterating through millions of cells.
- Measure execution benchmarking using `Timer` and integrate speech notifications via `Application.Speech.Speak`.

---

## 🧠 Architectural Takeaways & Mental Models
- **The Reverse Loop Rule**: Whenever deleting rows or items based on a condition, **always loop backwards**:
  ```vba
  For i = lastRow To 2 Step -1
      If Cells(i, 1).Value = "DeleteMe" Then Rows(i).Delete
  Next i
  ```
  *(Looping forwards causes row indexes to shift upward, skipping adjacent candidate rows).*
- **`Do While` vs `Do Until`**:
  - `Do While [Condition]` runs *as long as* the condition is True.
  - `Do Until [Condition]` runs *until* the condition becomes True (stops when True).
- **`Range.Find` Power**: `Range.Find` is hundreds of times faster than looping through thousands of cells. Always specify `LookIn`, `LookAt:=xlWhole`, and record the first matched address to terminate `FindNext`.

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- Infinite loops in `Do While` / `Do Until`: Occur when the counter or loop termination condition is never modified inside the loop body. (Break with `Ctrl + Break` or `Esc`).
- Forgetting to save `firstAddress = foundCell.Address` when using `FindNext` will cause an infinite loop searching the same matches endlessly.

---

## 📚 Lecture Index

- [ ] [[Lecture S10-01 - Section Overview Looping in VBA|Lecture 01: Section Overview: Looping in VBA]]
- [ ] [[Lecture S10-02 - For Next Counter Loops (Simple Example)|Lecture 02: For Next Counter Loops (Simple Example)]]
- [ ] [[Lecture S10-03 - For Next Nested Loop (Loop through Text & Cells)|Lecture 03: For Next Nested Loop (Loop through Text & Cells)]]
- [ ] [[Lecture S10-04 - For Next Reverse Loop (Delete Filtered Hidden Rows)|Lecture 04: For Next Reverse Loop (Delete Filtered/Hidden Rows)]]
- [ ] [[Lecture S10-05 - Do Until & Do While Loop|Lecture 05: Do Until & Do While Loop]]
- [ ] [[Lecture S10-06 - Practical Example of Do Loop (with Inputbox)|Lecture 06: Practical Example of Do Loop (with Inputbox)]]
- [ ] [[Lecture S10-07 - Find Method for Quicker Lookup Results|Lecture 07: Find Method for Quicker Lookup Results]]
- [ ] [[Lecture S10-08 - Find Method with Do Loop - Multiple Matches|Lecture 08: Find Method with Do Loop - Multiple Matches]]
- [ ] [[Lecture S10-09 - Add a Timer to Test & Speech to Inform|Lecture 09: Add a Timer to Test & Speech to Inform]]
- [ ] [[Lecture S10-10 - Activity Document all Comments in Workbook in a Sheet|Lecture 10: Activity: Document all Comments in Workbook in a Sheet]]
- [ ] [[Lecture S10-11 - Project Document all Comments in Workbook in a Sheet|Lecture 11: Project: Document all Comments in Workbook in a Sheet]]
- [ ] [[Lecture S10-12 - Useful VBA Statements & Key takeaways for Looping in VBA|Lecture 12: Useful VBA Statements & Key takeaways for Looping in VBA]]

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
