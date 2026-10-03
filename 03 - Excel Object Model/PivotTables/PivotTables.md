---
type: concept
domain: VBA
topic: Object Model
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 13 - Excel Tables, Formulas & Pivot Tables]]"
related_concepts:
  - "[[Excel Tables]]"
  - "[[Pivot Automation]]"
related_projects:
  - "[[Project 02 - Regional Sales Reporting Tool]]"
tags:
  - vba
  - pivottables
  - reporting
created: 2026-10-03
updated: 2026-10-03
---

# PivotTables

> [!important]
> **Summary Statement**: The `PivotTable` object in VBA allows automated creation, manipulation, source data updating, and refreshing of interactive aggregated summaries.

---

## 🏗️ Architecture: PivotCache vs. PivotTable

```mermaid
flowchart LR
    A[Raw Source Data Range] --> B["PivotCache (Memory Snapshot)"]
    B --> C["PivotTable 1 (Summary View)"]
    B --> D["PivotTable 2 (Detail View)"]
```

> [!tip]
> Multiple PivotTables built from the same data should share a single `PivotCache` to reduce file size and accelerate refresh times.

---

## 💻 Refreshing All PivotTables via Code
```vba
' Refresh all Pivot Caches across the entire workbook
Dim pc As PivotCache
For Each pc In ThisWorkbook.PivotCaches
    pc.Refresh
Next pc
```

---

## 🔗 Connected Notes
- [[Pivot Automation]], [[Project 02 - Regional Sales Reporting Tool]]
