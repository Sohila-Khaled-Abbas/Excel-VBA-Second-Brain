---
type: cheat-sheet
domain: VBA
topic: Syntax & Fundamentals
tags:
  - cheat-sheet
  - syntax
  - vba
  - fundamentals
created: 2026-10-03
updated: 2026-10-26
---

# ⚡ VBA Syntax & Fundamentals Cheat Sheet

> [!important]
> **Core Rule**: Always place `Option Explicit` as line 1 of every module. This enforces explicit variable declaration and eliminates silent misspelled variable bugs.

---

## 💾 Data Types & Memory Footprint Matrix

| Data Type | Memory | Value Range / Purpose | Best Practice Use Case |
| :--- | :---: | :--- | :--- |
| **`Byte`** | 1 Byte | 0 to 255 | Character loops, small counters |
| **`Boolean`** | 2 Bytes | `True` or `False` | Flags, toggle conditions, binary gates |
| **`Integer`** | 2 Bytes | -32,768 to 32,767 | **Avoid for row numbers** (Excel has 1M+ rows) |
| **`Long`** | 4 Bytes | -2,147,483,648 to 2,147,483,647 | **Default for integers & row counters** |
| **`Single`** | 4 Bytes | Single-precision floating point | Scientific modeling with low precision |
| **`Double`** | 8 Bytes | High-precision floating point | **Default for decimals, calculations, percentages** |
| **`Currency`** | 8 Bytes | 4 decimal fixed point, no rounding errors | Financial amounts, exact currency arithmetic |
| **`Date`** | 8 Bytes | Jan 1, 100 to Dec 31, 9999 | Dates, timestamps, date math (`DateAdd`) |
| **`String`** | 10 B + len | Text strings (variable length) | File paths, names, messages, keys |
| **`Variant`** | 16 Bytes | Can hold any type, errors, or arrays | **In-memory arrays**, functions that may return errors |

---

## 🏷️ Variable Declaration & Assignment

```vba
Option Explicit ' Forces declaration

' Primitive / Value Variables (use = without Set)
Dim lastRow As Long
Dim grossRevenue As Double
Dim customerName As String
Dim isCompleted As Boolean
Dim reportDate As Date

' Object Variables (MUST use Set!)
Dim wb As Workbook:   Set wb = ThisWorkbook
Dim ws As Worksheet:  Set ws = wb.Sheets("Data")
Dim rng As Range:     Set rng = ws.Range("A2:C100")
Dim lo As ListObject: Set lo = ws.ListObjects("SalesTable")

' Constants
Public Const COMPANY_NAME As String = "Automation Corp"
Const TAX_RATE As Double = 0.19
```

---

## 🌐 Variable Scope Matrix

| Scope | Keyword | Declaration Location | Lifetime | Visibility |
| :--- | :--- | :--- | :--- | :--- |
| **Procedure** | `Dim` | Inside a Sub or Function | While procedure runs | Current procedure only |
| **Module** | `Dim` or `Private` | Top of module (Declarations) | While workbook open | All Subs in that module |
| **Global** | `Public` | Top of standard module | While workbook open | Entire VBA project |

---

## 🔀 Conditionals & Branching

### 1. If...Then...ElseIf...Else
```vba
If score >= 90 Then
    grade = "A"
ElseIf score >= 80 Then
    grade = "B"
ElseIf score >= 70 Then
    grade = "C"
Else
    grade = "F"
End If
```

### 2. Select Case (Clean Alternative to Cascading Ifs)
```vba
Select Case score
    Case 90 To 100: grade = "A"
    Case 80 To 89:  grade = "B"
    Case 70 To 79:  grade = "C"
    Case Is < 70:   grade = "F"
    Case Else:      grade = "Invalid"
End Select
```

---

## 📦 Procedures vs Functions

```vba
' Sub: Performs actions, modifies state, returns no value
Public Sub FormatReport(ByVal wsTarget As Worksheet)
    wsTarget.Columns.AutoFit
End Sub

' Function: Computes and returns a value to the caller or cell
Public Function CalculateTax(ByVal netAmount As Double, ByVal taxRate As Double) As Double
    CalculateTax = netAmount * taxRate ' Assign to function name to return!
End Function
```
