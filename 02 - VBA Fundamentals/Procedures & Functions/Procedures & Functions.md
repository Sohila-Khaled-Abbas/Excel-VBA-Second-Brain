---
type: concept
domain: VBA
topic: Fundamentals
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 19 - Creating Customized Functions Formulas]]"
related_concepts:
  - "[[VBA Syntax]]"
  - "[[Scope]]"
  - "[[Custom VBA Functions]]"
tags:
  - vba
  - procedures
  - functions
created: 2026-10-03
updated: 2026-10-03
---

# Procedures & Functions (Sub vs. Function)

> [!important]
> **Summary Statement**: A `Sub` (Subroutine) performs actions and alters states without returning a value; a `Function` computes and returns a value to the caller and can also be used as a custom formula in Excel worksheet cells.

---

## ⚖️ Sub vs. Function Comparison

| Feature | `Sub` Procedure | `Function` Procedure |
| :--- | :--- | :--- |
| **Returns a Value?** | ❌ No | ✅ Yes (assigned to function name) |
| **Callable from Excel Cells?** | ❌ No | ✅ Yes (User-Defined Function / UDF) |
| **Can Modify Worksheets?** | ✅ Yes (format, write, delete) | ⚠️ When called from a cell, can ONLY return a value; cannot modify other cells or formatting! |
| **Trigger Mechanism** | Button, Hotkey, Macro dialog, Event | Called by formula, Sub, or other Function |

---

## 💻 Syntax & Examples

```vba
' SUB PROCEDURE: Performs actions
Sub ArchiveCompletedRows()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A2:G2").Interior.Color = RGB(220, 220, 220)
End Sub

' FUNCTION PROCEDURE: Calculates and returns a value
Function CalculateNetMargin(revenue As Double, cost As Double) As Double
    If revenue = 0 Then
        CalculateNetMargin = 0
    Else
        CalculateNetMargin = (revenue - cost) / revenue
    End If
End Function
```

---

## 🔁 Argument Passing: `ByVal` vs. `ByRef`
- **`ByRef` (Default in VBA)**: Passes the memory pointer. If the called procedure modifies the parameter, the caller's variable also changes!
- **`ByVal`**: Passes a detached copy. Protects the caller's variable from unintentional mutation.

---

## 🔗 Connected Notes
- **Custom Formulas**: [[Custom VBA Functions]]
- **Scope**: [[Scope]]
