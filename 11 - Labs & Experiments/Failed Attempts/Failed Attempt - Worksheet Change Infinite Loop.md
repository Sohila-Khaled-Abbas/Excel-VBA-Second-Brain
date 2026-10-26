---
title: Failed Attempt — Worksheet_Change Infinite Loop
type: post-mortem
domain: labs
status: resolved
tags:
  - vba/events
  - vba/debugging
  - vba/worksheet-events
  - vba/pitfalls
related:
  - "[[Worksheet Events]]"
  - "[[Defensive Programming]]"
  - "[[Error Handling]]"
created: 2026-10-26
updated: 2026-10-26
---

# Failed Attempt — Worksheet_Change Infinite Loop

> **Post-Mortem**: Documenting a subtle but catastrophic infinite event recursion that crashed Excel during UserForm data entry validation. Root cause identified, fix applied and generalized.

---

## What I Was Building
A `Worksheet_Change` event handler that auto-formats a data entry table: if a user enters a value in column B (Status column), the entire row's background color should change to match the status (Green = Complete, Yellow = In Progress, Red = Blocked).

---

## The Broken Code

```vba
' ❌ WRONG — Causes Infinite Loop!
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 2 Then ' Column B = Status
        Select Case LCase(Target.Value)
            Case "complete"
                Target.EntireRow.Interior.Color = RGB(198, 239, 206) ' Green
            Case "in progress"
                Target.EntireRow.Interior.Color = RGB(255, 235, 156) ' Yellow
            Case "blocked"
                Target.EntireRow.Interior.Color = RGB(255, 199, 206) ' Red
        End Select
    End If
End Sub
```

**What happened**: When `Target.EntireRow.Interior.Color = ...` executes, it modifies a cell property on the worksheet → this triggers `Worksheet_Change` again → which changes the color again → infinite recursion → **Excel freezes, crash/out-of-memory**.

---

## Root Cause Analysis

The `Worksheet_Change` event fires for **any property-level change** that Excel registers as a cell modification, including `.Interior.Color`. Setting a format property inside the event handler causes it to **call itself recursively**.

---

## The Fix: EnableEvents Guard

```vba
' ✅ CORRECT — Recursion Guard with EnableEvents
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Guard 1: Exit early if not in the Status column
    If Target.Column <> 2 Then Exit Sub
    
    ' Guard 2: Disable event recursion BEFORE modifying any sheet property
    Application.EnableEvents = False
    
    On Error GoTo RestoreEvents
    
    Select Case LCase(Target.Value)
        Case "complete"
            Target.EntireRow.Interior.Color = RGB(198, 239, 206)
        Case "in progress"
            Target.EntireRow.Interior.Color = RGB(255, 235, 156)
        Case "blocked"
            Target.EntireRow.Interior.Color = RGB(255, 199, 206)
        Case Else
            Target.EntireRow.Interior.ColorIndex = xlNone ' Clear color
    End Select

RestoreEvents:
    ' Guard 3: ALWAYS restore events, even on error
    Application.EnableEvents = True
End Sub
```

---

## Generalized Rule for All Event Handlers

```
ANY worksheet event handler that modifies the sheet MUST wrap with:
    Application.EnableEvents = False
    ... (modifications)
    Application.EnableEvents = True
```

| Event Type | Recursion Risk | Required Guard |
| :--- | :--- | :--- |
| `Worksheet_Change` | 🔴 **Critical** | `EnableEvents = False` |
| `Workbook_BeforeSave` | 🟡 **Medium** | `EnableEvents = False` if modifying cells |
| `Worksheet_SelectionChange` | 🟢 **Low** | Usually safe (selection ≠ data change) |
| `Worksheet_Calculate` | 🔴 **Critical** | Never recalculate inside `Calculate` event |

---

## Key Takeaway
The `EnableEvents` pattern is the **mandatory handshake** for all worksheet event handlers. Forgetting it isn't a code style issue — it's a runtime crash guarantee waiting to happen. Always pair `False` with `True` in an `On Error` block.
