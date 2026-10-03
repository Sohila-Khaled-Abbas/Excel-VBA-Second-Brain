---
type: concept
domain: VBA
topic: Events
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 16 - Worksheet and Workbook Events]]"
related_concepts:
  - "[[Workbook Events]]"
  - "[[Range]]"
tags:
  - vba
  - events
  - worksheet
created: 2026-10-03
updated: 2026-10-03
---

# Worksheet Events

> [!important]
> **Summary Statement**: Worksheet events are private event-handler procedures placed in a specific Sheet object module that fire automatically in response to user actions (e.g., changing a cell, clicking a cell, recalculating).

---

## 💻 Crucial Pattern: `Worksheet_Change` & `Application.EnableEvents`

```vba
' Placed inside specific Sheet module (e.g., Sheet1)
Private Sub Worksheet_Change(ByVal Target As Range)
    ' 1. Check if the changed cell is our target cell (B2)
    If Not Intersect(Target, Me.Range("B2")) Is Nothing Then
        
        ' 2. CRITICAL: Disable events to prevent infinite recursive loop!
        Application.EnableEvents = False
        On Error GoTo CleanExit
        
        ' 3. Perform action (e.g., force uppercase)
        Me.Range("B2").Value = UCase(Me.Range("B2").Value)
        
CleanExit:
        ' 4. Re-enable events
        Application.EnableEvents = True
    End If
End Sub
```

> [!warning]
> If your `Worksheet_Change` event modifies a cell value on the sheet, it will trigger another `Worksheet_Change` event!
> Without `Application.EnableEvents = False`, Excel enters an infinite recursive loop and crashes.

---

## 🔗 Connected Notes
- [[Workbook Events]], [[Range]]
