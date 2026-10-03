---
type: experiment
domain: VBA
topic: Performance
status: completed
related_concepts:
  - "[[Arrays]]"
  - "[[Performance Optimization]]"
tags:
  - experiment
  - benchmark
  - performance
created: 2026-10-03
updated: 2026-10-03
---

# Experiment 01: Cell Reading vs. Variant Array Benchmark

## 🔬 Hypothesis
Processing 50,000 rows in memory using a 2D `Variant` array will execute at least 50x faster than reading/writing cell-by-cell using `Cells(r, c)`.

---

## 💻 Benchmark Code

```vba
Sub BenchmarkTest()
    Dim tStart As Double
    Dim r As Long, total As Double
    
    ' TEST 1: Cell-by-Cell
    tStart = Timer
    For r = 1 To 50000
        total = total + Cells(r, 1).Value
    Next r
    Debug.Print "Cell-by-cell took: " & Round(Timer - tStart, 3) & "s"
    
    ' TEST 2: Variant Array in RAM
    Dim arrData As Variant
    tStart = Timer
    arrData = Range("A1:A50000").Value
    For r = 1 To 50000
        total = total + arrData(r, 1)
    Next r
    Debug.Print "Array in RAM took: " & Round(Timer - tStart, 3) & "s"
End Sub
```

---

## 📊 Observations
- **Cell-by-Cell**: ~3.85 seconds
- **Variant Array**: ~0.04 seconds
- **Speedup**: **~96x faster**!
