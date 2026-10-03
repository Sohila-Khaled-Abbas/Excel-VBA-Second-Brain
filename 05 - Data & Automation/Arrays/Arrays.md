---
type: concept
domain: VBA
topic: Data Structures
difficulty: advanced
status: learning
source: Udemy
course_section: "[[Section 11 - Working with Arrays]]"
related_concepts:
  - "[[Data Types]]"
  - "[[Performance Optimization]]"
related_projects:
  - "[[Project 02 - Regional Sales Reporting Tool]]"
tags:
  - vba
  - arrays
  - performance
created: 2026-10-03
updated: 2026-10-03
---

# Arrays (In-Memory Data Structures)

> [!important]
> **Summary Statement**: An array is a contiguous block of indexed memory locations holding multiple values of the same or variant type, executing up to 100x faster than cell-by-cell sheet reading.

---

## 🚀 The 100x Speed Secret: Variant Arrays
Reading a 50,000-cell range into memory takes a single line of code and milliseconds of execution time:

```vba
Sub ProcessDataInMemory()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim rawData As Variant
    Dim r As Long, c As Long
    
    ' 1. SUCK DATA FROM SHEET INTO RAM IN 0.02 SECONDS
    rawData = ws.Range("A2:D50000").Value
    
    ' rawData is now a 2D 1-based array: rawData(row, col)
    For r = 1 To UBound(rawData, 1)
        ' Process in pure RAM
        rawData(r, 4) = rawData(r, 2) * rawData(r, 3) ' Quantity * Price
    Next r
    
    ' 2. DUMP PROCESSED DATA BACK TO SHEET IN 0.02 SECONDS
    ws.Range("A2:D50000").Value = rawData
End Sub
```

---

## ⚙️ Dynamic Arrays & `ReDim Preserve`
- `Dim arr() As String`: Dynamic array initialized without fixed size.
- `ReDim arr(1 To 100)`: Size allocated at runtime.
- `ReDim Preserve arr(1 To 200)`: Expands size while keeping existing values. *Limitation: In multi-dimensional arrays, you can only resize the LAST dimension when using Preserve.*

---

## 🔗 Connected Notes
- [[Performance Optimization]], [[Arrays Cheat Sheet]]
