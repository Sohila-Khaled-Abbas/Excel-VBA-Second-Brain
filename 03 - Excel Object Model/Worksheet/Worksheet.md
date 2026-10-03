---
type: concept
domain: VBA
topic: Object Model
difficulty: beginner
status: mastered
source: Udemy
course_section: "[[Section 04 - Referencing Ranges, Worksheets & Workbooks]]"
related_concepts:
  - "[[Workbook]]"
  - "[[Range]]"
related_projects:
  - "[[Project 01 - Table of Contents Generator]]"
tags:
  - vba
  - worksheet
  - sheet
created: 2026-10-03
updated: 2026-10-03
---

# Worksheet Object

> [!important]
> **Summary Statement**: The `Worksheet` object represents a single 2D grid tab inside a workbook, containing cells, formulas, tables, and drawing objects.

---

## 🏷️ 3 Ways to Reference a Worksheet in VBA

1. **Tab Name (User-Facing)**:
   `Worksheets("Sales Data")`
   *Fragile: If a user renames the tab to "Sales_Data", code crashes.*
2. **Index Position**:
   `Worksheets(1)`
   *Fragile: If a user drags the sheet tab to a different order, code points to the wrong sheet.*
3. **VBA CodeName (Best Practice for Internal Sheets)**:
   `wsSalesData.Range("A1")`
   *Indestructible: Configured in VBE Properties window; immune to tab renaming and tab reordering.*

---

## 💻 Adding and Manipulating Worksheets

```vba
Dim wsNew As Worksheet
Set wsNew = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
wsNew.Name = "Q4_Summary"
wsNew.Tab.Color = RGB(16, 124, 65)
```

---

## 🔗 Connected Notes
- [[Range]], [[Cells]], [[Workbook]]
