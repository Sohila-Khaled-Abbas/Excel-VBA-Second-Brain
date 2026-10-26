---
type: course-section
section_number: 11
section_title: Working with Arrays
lecture_count: 9
status: not-started
tags:
  - course-section
  - vba
  - section-11
---

# Section 11: Working with Arrays

> [!important]
> **Curriculum Pillar**: Section 11 contains **9 lectures**.  
> **Status**: `NOT-STARTED`  
> **Course Practice Files**: `SectionFiles/S11_Arrays.xlsm`, `S11_Arrays_Start.xlsx`

---

## 🎯 Section Objectives & Automation Capabilities
- Understand In-Memory Arrays as the ultimate speed multiplier in Excel VBA.
- Declare, dimension, and populate Static 1D and 2D arrays.
- Create Dynamic Arrays that adjust size at runtime using `ReDim` and `ReDim Preserve`.
- Master the **Variant Memory Array Pattern**: Read millions of cells into RAM in one operation, process in microseconds, and write back in one blast.
- Build automated multi-sheet reports by categorizing and filtering array records in memory.

---

## 🧠 Architectural Takeaways & Mental Models
- **The Memory Multiplier**: Reading and writing cells on a worksheet involves significant COM overhead. Transferring data to a Variant array performs computations in RAM at microsecond speeds:
  ```vba
  Dim dataArr As Variant
  dataArr = Range("A2:D100000").Value ' Instant bulk read to 2D array
  ' ... process dataArr(r, c) entirely in RAM ...
  Range("F2:I100000").Value = dataArr ' Instant bulk write back to sheet
  ```
- **1D vs 2D Arrays**:
  - 1D array (`Dim arr(1 To 5)`): Vector of elements. Writes horizontally to a row by default.
  - 2D array (`Dim arr(1 To 10, 1 To 4)`): Grid of rows and columns.
- **`ReDim Preserve` Limitation**: You can only preserve existing values when resizing the **last dimension** of a multidimensional array (`ReDim Preserve arr(1 To 10, 1 To newColCount)`).

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- Excel ranges dumped into a Variant array are always **1-based, 2-dimensional**, even if only a single row or single column is copied (`dataArr(1 To n, 1 To 1)`).
- Attempting `ReDim Preserve arr(newRowCount, 4)` will trigger runtime error 9 ("Subscript out of range") because only the upper bound of the last dimension can be resized.

---

## 📚 Lecture Index

- [ ] [[Lecture S11-01 - Overview Working with Arrays in VBA|Lecture 01: Overview: Working with Arrays in VBA]]
- [ ] [[Lecture S11-02 - One Dimensional Arrays (and transferring back to sheet)|Lecture 02: One Dimensional Arrays (and transferring back to sheet)]]
- [ ] [[Lecture S11-03 - Dynamic Arrays (size adjusts depending on a condition)|Lecture 03: Dynamic Arrays (size adjusts depending on a condition)]]
- [ ] [[Lecture S11-04 - Preserving Dynamic Arrays (adding more elements to the existing array)|Lecture 04: Preserving Dynamic Arrays (adding more elements to the existing array)]]
- [ ] [[Lecture S11-05 - Two Dimensional Arrays|Lecture 05: Two Dimensional Arrays]]
- [ ] [[Lecture S11-06 - Variant Arrays Quick Method to Write to Cells|Lecture 06: Variant Arrays: Quick Method to Write to Cells]]
- [ ] [[Lecture S11-07 - Activity Create a New Worksheet per Company with Specific Information|Lecture 07: Activity: Create a New Worksheet per Company with Specific Information]]
- [ ] [[Lecture S11-08 - Project Create a New Worksheet per Company with Specific Information|Lecture 08: Project: Create a New Worksheet per Company with Specific Information]]
- [ ] [[Lecture S11-09 - Key Takeaways Arrays|Lecture 09: Key Takeaways Arrays]]

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
