---
type: concept
domain: VBA
topic: Object Model
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 13 - Excel Tables, Formulas & Pivot Tables]]"
related_concepts:
  - "[[Range]]"
  - "[[PivotTables]]"
tags:
  - vba
  - tables
  - listobjects
created: 2026-10-03
updated: 2026-10-03
---

# Excel Tables (ListObjects)

> [!important]
> **Summary Statement**: In VBA, formal Excel Tables are represented by the `ListObject` object, providing structured references, automatic range expansion, and isolated table manipulation.

---

## 💻 Interacting with ListObjects in VBA

```vba
Dim ws As Worksheet: Set ws = ActiveSheet
Dim tbl As ListObject: Set tbl = ws.ListObjects("SalesTable")

' Reference specific structured parts of the table
Dim dataBody As Range: Set dataBody = tbl.DataBodyRange
Dim headerRng As Range: Set headerRng = tbl.HeaderRowRange
Dim revenueCol As Range: Set revenueCol = tbl.ListColumns("Revenue").Range

' Add a new row dynamically
Dim newRow As ListRow: Set newRow = tbl.ListRows.Add
newRow.Range(1, 1).Value = "East Region"
newRow.Range(1, 2).Value = 450000
```

---

## 🔗 Connected Notes
- [[Range]], [[PivotTables]]
