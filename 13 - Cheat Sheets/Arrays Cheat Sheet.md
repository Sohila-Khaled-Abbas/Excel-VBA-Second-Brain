---
type: cheat-sheet
domain: VBA
topic: Arrays
tags:
  - cheat-sheet
  - arrays
  - memory
created: 2026-10-03
updated: 2026-10-03
---

# ⚡ Arrays Cheat Sheet

```vba
' 1. SUCK CELL BLOCK INTO MEMORY ARRAY (1-based 2D array)
Dim arr As Variant
arr = Range("A1:D100").Value

' Dimensions:
Dim numRows As Long: numRows = UBound(arr, 1) ' 100
Dim numCols As Long: numCols = UBound(arr, 2) ' 4

' 2. READ / MODIFY IN RAM
Dim r As Long, c As Long
For r = 1 To numRows
    arr(r, 4) = arr(r, 2) * arr(r, 3)
Next r

' 3. DUMP BACK TO WORKSHEET IN ONE SHOT
Range("A1:D100").Value = arr

' 4. DYNAMIC ARRAY RESIZING
Dim dynArr() As String
ReDim dynArr(1 To 10)
ReDim Preserve dynArr(1 To 20) ' Preserves existing contents
```
