---
type: cheat-sheet
domain: VBA
topic: Debugging & Reliability
tags:
  - cheat-sheet
  - debugging
  - error-handling
  - vba
created: 2026-10-03
updated: 2026-10-26
---

# 🔍 Debugging & Reliability Cheat Sheet

---

## ⌨️ VBE Debugging Hotkeys

| Shortcut | Command | Action |
| :--- | :--- | :--- |
| **`F8`** | **Step Into** | Executes code line by line, entering called procedures |
| **`Shift + F8`** | **Step Over** | Executes code line by line; runs called subs without stepping in |
| **`Ctrl + Shift + F8`** | **Step Out** | Runs remaining lines of current Sub and pauses at caller |
| **`F9`** | **Toggle Breakpoint** | Sets / removes execution pause mark on current line |
| **`Ctrl + F8`** | **Run to Cursor** | Runs execution until paused at current cursor position |
| **`F5`** | **Run / Continue** | Runs macro to completion or next breakpoint |
| **`Ctrl + G`** | **Immediate Window** | Opens evaluation & scratchpad console |
| **`Ctrl + Break`** | **Pause Execution** | Halts infinite loops or frozen scripts |

---

## 🛠️ Defensive Error Handling Architecture

```vba
Public Sub ProductionProcedure()
    On Error GoTo ErrorHandler
    
    ' Disable environment rendering
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' --- CORE BUSINESS LOGIC ---
    
CleanExit:
    ' GUARANTEED RESTORATION: Always executed!
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Execution Failed"
    Resume CleanExit
End Sub
```

---

## 🚨 Common VBA Runtime Error Codes

| Err # | Description | Common Cause & Fix |
| :---: | :--- | :--- |
| **`9`** | **Subscript out of range** | Worksheet or workbook name misspelled; array index out of bounds. |
| **`13`** | **Type mismatch** | Assigning text to a numeric variable; cell contains `#N/A` or `#VALUE!`. |
| **`91`** | **Object variable not set** | Forgot the `Set` keyword when assigning a Workbook/Worksheet/Range. |
| **`1004`** | **Application-defined / Object-defined** | Unqualified range reference, protected worksheet, invalid range address. |
| **`424`** | **Object required** | User clicked Cancel on `Application.InputBox(Type:=8)`. |
| **`6`** | **Overflow** | Integer variable exceeded 32,767. (Fix: Change to `Long`). |
