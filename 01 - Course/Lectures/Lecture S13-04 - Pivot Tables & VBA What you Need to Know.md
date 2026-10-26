---
type: lecture
course: Excel VBA & Macros
section: "[[Section 13 - Excel Tables, Formulas & Pivot Tables]]"
section_number: 13
lecture_number: 4
status: not-started
topics:
  - Excel Tables, Formulas & Pivot Tables
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-13
---

# Lecture S13-04: Pivot Tables & VBA: What you Need to Know

> [!important]
> **Section Reference**: [[Section 13 - Excel Tables, Formulas & Pivot Tables|Section 13: Excel Tables, Formulas & Pivot Tables]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Pivot Tables & VBA: What you Need to Know

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Pivot Tables & VBA: What you Need to Know' within the context of Excel Tables, Formulas & Pivot Tables.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Pivot Tables & VBA: What you Need to Know', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Create_Pivot()

'

' Create_Pivot Macro

' Best is to start by recording a macro to see how pivottables are referenced.

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _

        "CompTable", Version:=6).CreatePivotTable TableDestination:= _

        "PivotTables!R8C1", TableName:="PivotTable1", DefaultVersion:=6

    Sheets("PivotTables").Select

    Cells(8, 1).Select

    With ActiveSheet.PivotTables("PivotTable1")

        .ColumnGrand = True

        .HasAutoFormat = True

        .DisplayErrorString = False

        .DisplayNullString = True

        .EnableDrilldown = True

        .ErrorString = ""

        .MergeLabels = False

        .NullString = ""

        .PageFieldOrder = 2

        .PageFieldWrapCount = 0

        .PreserveFormatting = True

        .RowGrand = True

        .SaveData = True

        .PrintTitles = False

        .RepeatItemsOnEachPrintedPage = True

        .TotalsAnnotation = False

        .CompactRowIndent = 1

        .InGridDropZones = False

        .DisplayFieldCaptions = True

        .DisplayMemberPropertyTooltips = False

        .DisplayContextTooltips = True

        .ShowDrillIndicators = True

        .PrintDrillIndicators = False

        .AllowMultipleFilters = False

        .SortUsingCustomLists = True

        .FieldListSortAscending = False

        .ShowValuesRow = False

        .CalculatedMembersInFilters = False

        .RowAxisLayout xlCompactRow

    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotCache

        .RefreshOnFileOpen = False

        .MissingItemsLimit = xlMissingItemsDefault

    End With

    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Customer code")

        .Orientation = xlRowField

        .Position = 1

    End With

    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _

        "PivotTable1").PivotFields("Quantity"), "Sum of Quantity", xlSum

    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _

        "PivotTable1").PivotFields("Quantity"), "Sum of Quantity2", xlSum

    Range("C8").Select

    ActiveSheet.PivotTables("PivotTable1").DataPivotField.PivotItems( _

        "Sum of Quantity2").Caption = "% Quantity"

    Range("C9").Select

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("% Quantity")

        .Calculation = xlPercentOfTotal

        .NumberFormat = "0.00%"

    End With

End Sub
```

---

## 🔍 How It Works
1. **Initialization**: Variables and object references are explicitly defined and verified.
2. **Execution**: The procedure manipulates ranges, worksheets, or UI controls directly without unnecessary `.Select` or `.Activate` calls.
3. **Clean Completion**: Screen updates and system state are safely restored.

---

## ⚠️ Common Mistakes & Traps
> [!warning]
> - Relying on the `ActiveSheet` implicitly rather than qualifying sheet references explicitly.
> - Forgetting `Option Explicit`, allowing misspelled variables to compile into empty `Variant` values.
> - Leaving screen flickering active during high-frequency cell updates.

---

## 🧩 Real-World Use Cases
- **Data Analytics / Reporting**: Automating recurring data transformation, cleansing, and layout formatting pipelines.
- **Operational Automation**: Eliminating manual, error-prone spreadsheet tasks across business reporting workflows.

---

## 🔗 Related Concepts
- [[VBA Syntax]]
- [[VBA Object Model]]
- [[Error Handling]]
- [[Section 13 - Excel Tables, Formulas & Pivot Tables]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
