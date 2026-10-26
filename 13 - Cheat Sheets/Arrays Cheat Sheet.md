---
type: cheat-sheet
domain: VBA
topic: In-Memory Arrays
tags:
  - cheat-sheet
  - arrays
  - memory
  - performance
created: 2026-10-03
updated: 2026-10-26
---

# ⚡ In-Memory Arrays Cheat Sheet

> [!important]
> **The 100x Speed Multiplier**: Reading and writing cells on a worksheet involves heavy COM overhead. Transferring data to a Variant memory array performs computations in RAM at microsecond speeds.

---

## 🚀 The 3-Step In-Memory Pattern

```vba
' 1. SUCK CELL BLOCK INTO MEMORY ARRAY (Instant bulk read)
' Note: Excel range-to-variant arrays are always 1-based, 2-dimensional!
Dim dataArr As Variant
dataArr = ws.Range("A2:D100000").Value

Dim r As Long, totalRows As Long, totalCols As Long
totalRows = UBound(dataArr, 1) ' Rows count
totalCols = UBound(dataArr, 2) ' Columns count

' 2. PROCESS ENTIRELY IN RAM
For r = 1 To totalRows
    ' Multiply Col 2 by Col 3 and store in Col 4
    dataArr(r, 4) = dataArr(r, 2) * dataArr(r, 3)
Next r

' 3. DUMP BACK TO WORKSHEET IN ONE SHOT (Instant bulk write)
ws.Range("A2:D100000").Value = dataArr
```

---

## 📐 Array Types & Dimensionality

```vba
' 1D Fixed Array (0-based by default, or 1 To N)
Dim arr1D(1 To 5) As String
arr1D(1) = "Apple": arr1D(2) = "Banana"

' Write 1D Array to Worksheet:
' Note: A 1D array writes HORIZONTALLY across columns by default!
Range("A1:E1").Value = arr1D

' To write a 1D array VERTICALLY down a column, use Application.Transpose:
Range("A1:A5").Value = Application.Transpose(arr1D)

' 2D Dynamic Array
Dim arr2D() As Variant
ReDim arr2D(1 To 100, 1 To 4)

' ReDim Preserve Rule:
' You can ONLY resize the LAST dimension when preserving contents!
ReDim Preserve arr2D(1 To 100, 1 To 8) ' Valid!
' ReDim Preserve arr2D(1 To 200, 1 To 4) ' RUNTIME ERROR 9!
```
