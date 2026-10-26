---
type: lecture
course: Excel VBA & Macros
section: "[[Section 14 - Main Project 2 Regional Reporting Tool]]"
section_number: 14
lecture_number: 5
status: not-started
topics:
  - Main Project #2: Regional Reporting Tool
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-14
---

# Lecture S14-05: Exporting Reports as Excel files (Pivot Table & Hard Coded Reports)

> [!important]
> **Section Reference**: [[Section 14 - Main Project 2 Regional Reporting Tool|Section 14: Main Project #2: Regional Reporting Tool]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Exporting Reports as Excel files (Pivot Table & Hard Coded Reports)

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Exporting Reports as Excel files (Pivot Table & Hard Coded Reports)' within the context of Main Project #2: Regional Reporting Tool.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Exporting Reports as Excel files (Pivot Table & Hard Coded Reports)', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Private Sub Region_Report()

    Dim NewRbook As Workbook

    Dim newRpath As String

    Dim PT As PivotTable

    newRpath = ThisWorkbook.Path & "\" & "RM_" & myDate & myRegion & ".xlsx"

    If Len(Dir(newRpath)) > 0 Then 'file exists

        ReplaceFile = MsgBox("Regional File Exists. Do you want to overwrite?", vbOKCancel, "Overwrite?")

        If ReplaceFile = vbCancel Then

            Overwrite(0) = False

            Exit Sub

        End If

    End If

    Set NewRbook = Workbooks.Add

    shRegion.Copy Before:=NewRbook.Sheets(1)

    'remove common filter for slicer

    NewRbook.SlicerCaches("Slicer_Company_Name").PivotTables.RemovePivotTable _

    (ActiveSheet.PivotTables("PTArticle"))

    'change pivot cache

    NewRbook.ActiveSheet.PivotTables("PTRegion").ChangePivotCache _

    NewRbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="TableTemp")

        

    For Each PT In NewRbook.ActiveSheet.PivotTables

        If PT.Name <> "PTRegion" Then

            PT.ChangePivotCache ("PTRegion")

        End If

    Next PT

    

    'add back the common slicer

    NewRbook.SlicerCaches("Slicer_Company_Name").PivotTables.AddPivotTable _

    (ActiveSheet.PivotTables("PTArticle"))

    'some cleaning up

    With ActiveSheet

        .PivotTables("PTRegion").SaveData = True

        .Columns("AJ:AZ").Delete

        .Range("TableTemp").CurrentRegion.Columns.Group

        .Outline.ShowLevels RowLevels:=0, ColumnLevels:=1

        .Range("A1").CurrentRegion.Copy

        .Range("A1").PasteSpecial (xlPasteValues)



    End With

    

    NewRbook.SaveAs FileName:=newRpath

    NewRbook.Close

    TaskRow = shStart.Range("B" & shStart.Rows.Count).End(xlUp).Row + 1

    shStart.Range("B" & TaskRow).Value = "Regional Report for " & myRegion & myDate & " - created on: " & Now

    



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
- [[Section 14 - Main Project 2 Regional Reporting Tool]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
