---
title: Universal Speed Booster Wrapper
type: snippet
domain: code_cookbook
status: evergreen
tags:
  - vba/performance
  - vba/optimization
  - vba/snippet
related:
  - "[[Performance Optimization]]"
  - "[[Defensive Programming]]"
created: 2026-10-26
updated: 2026-10-26
---

# Universal Speed Booster Wrapper

> **Purpose**: Eliminates screen flicker, stops event recursion, and boosts routine execution speed by up to 10x by safely toggling Excel application-level properties with bulletproof restoration on exit or error.

---

## Production Snippet: Speed Booster Class or Sub Pair

```vba
Option Explicit

' Global variables to store initial state
Private pCalcMode As XlCalculation
Private pScreenUpdating As Boolean
Private pEvents As Boolean
Private pStatusBar As Boolean

Public Sub ToggleSpeedBoost(ByVal Enable As Boolean)
    With Application
        If Enable Then
            ' 1. Capture current user settings
            pCalcMode = .Calculation
            pScreenUpdating = .ScreenUpdating
            pEvents = .EnableEvents
            pStatusBar = .DisplayStatusBar
            
            ' 2. Turn off expensive rendering & calculations
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayAlerts = False
            .Calculation = xlCalculationManual
            .Cursor = xlWait
        Else
            ' 3. Restore original settings cleanly
            .ScreenUpdating = pScreenUpdating
            .EnableEvents = pEvents
            .DisplayAlerts = True
            .Calculation = pCalcMode
            .DisplayStatusBar = pStatusBar
            .StatusBar = False
            .Cursor = xlDefault
        End If
    End With
End Sub
```

---

## Real-World Usage Example

```vba
Public Sub ProcessLargeDataset()
    On Error GoTo ErrorHandler
    
    ' Engage booster at beginning
    ToggleSpeedBoost True
    Application.StatusBar = "Processing records in memory..."
    
    ' -------------------------------------------------------------
    ' Critical Business Logic / In-Memory Transformation
    ' -------------------------------------------------------------
    Dim i As Long
    For i = 1 To 50000
        ' Operations execute at raw C++ native speed without UI paint
    Next i
    
CleanExit:
    ' GUARANTEED RESTORATION: Always restore state in exit block
    ToggleSpeedBoost False
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during batch processing: " & Err.Description, vbCritical, "Execution Failure"
    Resume CleanExit
End Sub
```

---

## Key Rules
1. **Never forget the Exit Sub / ErrorHandler restoration**: If your code errors out while `ScreenUpdating = False` and `EnableEvents = False`, Excel will appear permanently frozen to the user!
2. **Never leave Calculation on Manual**: Always restore the previous calculation mode (`pCalcMode`) instead of forcing `xlCalculationAutomatic`, respecting user preferences.
