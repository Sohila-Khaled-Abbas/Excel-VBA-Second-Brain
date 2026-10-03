---
type: snippet
language: VBA
category: optimization
difficulty: intermediate
use_case: Wrap heavy automation routines in turbo mode
objects:
  - Application
tags:
  - vba
  - snippet
  - performance
created: 2026-10-03
updated: 2026-10-03
---

# Snippet: Turbo Performance Execution Wrapper

## 🎯 Problem Statement
Long-running macros cause screen flickering, UI freezing, and sluggish execution by recalculating formulas and updating screen pixels on every row iteration.

---

## 💻 VBA Code Implementation

```vba
Option Explicit

Private mPrevScreenUpdating As Boolean
Private mPrevDisplayAlerts As Boolean
Private mPrevCalculation As XlCalculation
Private mPrevEnableEvents As Boolean

Public Sub FastModeOn()
    ' Cache original application state
    mPrevScreenUpdating = Application.ScreenUpdating
    mPrevDisplayAlerts = Application.DisplayAlerts
    mPrevCalculation = Application.Calculation
    mPrevEnableEvents = Application.EnableEvents
    
    ' Engage turbo execution
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Public Sub FastModeOff()
    ' Restore user application state
    Application.ScreenUpdating = mPrevScreenUpdating
    Application.DisplayAlerts = mPrevDisplayAlerts
    Application.Calculation = mPrevCalculation
    Application.EnableEvents = mPrevEnableEvents
End Sub

' Example Usage:
Public Sub RunHeavyConsolidation()
    On Error GoTo ErrorHandler
    FastModeOn
    
    ' Heavy ETL / Loop operations here...
    
CleanExit:
    FastModeOff
    Exit Sub
ErrorHandler:
    MsgBox "Failure: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

---

## 🔗 Related Notes
- [[Performance Optimization]], [[Application]]
