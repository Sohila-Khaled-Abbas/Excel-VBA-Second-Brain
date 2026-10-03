---
type: concept
domain: VBA
topic: File I/O
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 12 - Working with Files, Folders & Text Files]]"
related_concepts:
  - "[[Text Files]]"
  - "[[Files]]"
related_projects:
  - "[[Project 02 - Regional Sales Reporting Tool]]"
tags:
  - vba
  - csv
  - etl
created: 2026-10-03
updated: 2026-10-03
---

# CSV (Comma/Semicolon Separated Values)

> [!important]
> **Summary Statement**: CSV export in VBA bridges Excel reports with downstream ERP systems (SAP, Oracle) and analytics warehouses.

---

## ⚠️ Regional Delimiter Warning (Comma vs. Semicolon)
> [!warning]
> Standard Excel `SaveAs FileFormat:=xlCSV` uses the Windows Regional Settings delimiter (commas in US/UK, semicolons in Europe).
> For robust data integration across international teams, **always write custom CSV export loops using low-level I/O (`Print #`)** to guarantee explicit delimiters.

---

## 🔗 Connected Notes
- [[Text Files]], [[Project 02 - Regional Sales Reporting Tool]]
