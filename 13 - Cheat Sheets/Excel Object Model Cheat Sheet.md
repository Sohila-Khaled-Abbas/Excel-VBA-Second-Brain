---
type: cheat-sheet
domain: VBA
topic: Excel Object Model
tags:
  - cheat-sheet
  - object-model
  - vba
  - range
created: 2026-10-03
updated: 2026-10-26
---

# 🏛️ Excel Object Model Cheat Sheet

> [!important]
> **Hierarchy Architecture**:  
> `Application` → `Workbooks("Book.xlsx")` → `Worksheets("Data")` → `Range("A1")` / `ListObjects("Table1")`

---

## 🎯 The 4 Last Row Methods Evaluated

| Method | Syntax | Pros | Cons / Gotchas |
| :--- | :--- | :--- | :--- |
| **1. End(xlUp)** *(Industry Standard)* | `Cells(Rows.Count, "A").End(xlUp).Row` | Mimics `Ctrl+Up`. Completely reliable if column A has values. | Ignores data in other columns if longer than column A. |
| **2. Find Method** | `Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row` | Finds true last row across **all** columns. | Throws error if sheet is completely blank. |
| **3. UsedRange** | `ActiveSheet.UsedRange.Rows.Count` | Quick one-liner. | Retains cached deleted cells; fails if data starts below row 1. |
| **4. End(xlDown)** | `Range("A1").End(xlDown).Row` | Mimics `Ctrl+Down`. | **Fails if column contains intermediate blank cells!** |

---

## 📑 Worksheet Referencing Tiers

```vba
' Tier 1: Tab Name (Fragile: Breaks if user renames sheet)
Worksheets("Sales Data").Range("A1")

' Tier 2: Index Number (Fragile: Breaks if user reorders tabs)
Worksheets(1).Range("A1")

' Tier 3: VBA CodeName (THE GOLD STANDARD: Immune to renaming & reordering)
shSales.Range("A1") 
```

---

## ⚡ Essential Range Operations

```vba
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("RawData")

' 1. Referencing
ws.Range("A1")              ' Static cell
ws.Cells(r, c)              ' Dynamic row & column index
ws.Range("A2:C10")          ' Multi-cell block
ws.Range("A" & r, "C" & r)  ' Dynamic row block

' 2. Resizing & Offsetting
ws.Range("A1").Offset(1, 0)                  ' Down 1 row
ws.Range("A1").Resize(10, 5)                 ' 10 rows by 5 columns
ws.Range("A1").CurrentRegion                 ' Entire contiguous table (Ctrl+A)
ws.Range("A1").CurrentRegion.Offset(1, 0) _
  .Resize(ws.Range("A1").CurrentRegion.Rows.Count - 1) ' Strip headers!

' 3. Fast Value Transfer (No Clipboard / Instant)
wsDest.Range("A1:C100").Value = wsSource.Range("A1:C100").Value

' 4. Clearing Data
ws.Range("A2:Z100").ClearContents ' Deletes values; preserves formats & validation
ws.Range("A2:Z100").Clear         ' Deletes values, formatting, notes, hyperlinks
```

---

## 📊 Excel Tables (`ListObjects`) Quick Reference

```vba
Dim lo As ListObject
Set lo = ws.ListObjects("InvoiceTable")

' Range components
Set rngHeaders = lo.HeaderRowRange
Set rngData    = lo.DataBodyRange
Set rngTotals  = lo.TotalsRowRange

' Specific Column
Set colAmount = lo.ListColumns("Amount").DataBodyRange

' Add New Row
Dim newRow As ListRow
Set newRow = lo.ListRows.Add(AlwaysInsert:=True)
newRow.Range(1, 1).Value = "Cust123"
```
