---
type: cheat-sheet
domain: VBA
topic: Syntax
tags:
  - cheat-sheet
  - syntax
  - vba
created: 2026-10-03
updated: 2026-10-03
---

# ⚡ VBA Syntax Cheat Sheet

## Variable Declaration
```vba
Option Explicit                   ' Always at line 1 of every module

Dim rowCount As Long              ' Default for integers / row numbers
Dim totalRev As Double            ' Default for decimals / currency
Dim clientName As String          ' Text strings
Dim isActive As Boolean           ' True / False
Dim reportDate As Date            ' Date / Time values
Dim rawData As Variant            ' Universal container (or in-memory arrays)

Dim ws As Worksheet               ' Object reference
Set ws = ActiveSheet              ' Requires 'Set' keyword!
```

## Conditionals
```vba
If x > 100 Then
    status = "High"
ElseIf x > 50 Then
    status = "Medium"
Else
    status = "Low"
End If

Select Case region
    Case "East", "North": rate = 0.05
    Case "West":          rate = 0.07
    Case Else:            rate = 0.00
End Select
```

## Procedures
```vba
' Sub: Executes actions
Public Sub DoSomething(ByVal customerID As Long)
    ' ...
End Sub

' Function: Calculates and returns value
Public Function CalculateTax(ByVal amount As Double, ByVal rate As Double) As Double
    CalculateTax = amount * rate
End Function
```
