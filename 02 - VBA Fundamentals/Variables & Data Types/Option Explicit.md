---
type: concept
domain: VBA
topic: Best Practices
difficulty: beginner
status: mastered
source: Udemy
course_section: "[[Section 05 - Working with Variables]]"
related_concepts:
  - "[[Variables]]"
  - "[[Dim Statement]]"
  - "[[Debugging]]"
tags:
  - vba
  - best-practices
  - reliability
created: 2026-10-03
updated: 2026-10-03
---

# Option Explicit

> [!important]
> **Summary Statement**: `Option Explicit` is a module-level compiler directive that forces the developer to explicitly declare every variable before use, catching typos at compile-time before code executes.

---

## 📖 Why It Is Mandatory
Without `Option Explicit`, VBA creates a brand-new `Variant` variable on the fly whenever it encounters an unrecognized identifier. A single spelling mistake can create silent, catastrophic bugs:

```vba
' WITHOUT Option Explicit:
Sub CalculateTax()
    totalRevenue = 100000
    taxRate = 0.2
    ' Typo in variable name below creates an empty Variant (0):
    finalTax = totelRevenue * taxRate 
    ' Result is 0 instead of 20000, and VBA reports NO error!
End Sub

' WITH Option Explicit:
Option Explicit

Sub CalculateTax()
    Dim totalRevenue As Double, taxRate As Double, finalTax As Double
    totalRevenue = 100000
    taxRate = 0.2
    finalTax = totelRevenue * taxRate ' COMPILE ERROR: Variable not defined!
End Sub
```

---

## 💡 How to Automate
In the Visual Basic Editor, go to **Tools → Options → Editor Tab → Check "Require Variable Declaration"**. Every newly created module will automatically have `Option Explicit` placed at line 1.

---

## 🔗 Connected Notes
- **Debugging & Prevention**: [[Defensive Programming]], [[Debugging]]
