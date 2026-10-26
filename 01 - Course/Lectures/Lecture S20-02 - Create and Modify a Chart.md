---
type: lecture
course: Excel VBA & Macros
section: "[[Section 20 - Working with Charts]]"
section_number: 20
lecture_number: 2
status: not-started
topics:
  - Working with Charts
related_concepts:
  - '[[VBA Syntax]]'
  - '[[Excel Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-20
---

# Lecture S20-02: Create and Modify a Chart

> [!important]
> **Section Reference**: [[Section 20 - Working with Charts|Section 20: Working with Charts]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Create and Modify a Chart

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Create and Modify a Chart' within the context of Working with Charts.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Create and Modify a Chart', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub My_Column_Chart()
'
' My_Column_Chart Macro
'
    If Selection.Cells.Count > 1 Then
        Set myChart = ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Chart
        With myChart
            .Axes(xlValue).MajorGridlines.Delete
            .ChartGroups(1).GapWidth = 80
            .FullSeriesCollection(1).ApplyDataLabels
            .Axes(xlValue).Delete

            With .FullSeriesCollection(1).Format.Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.5
                .Transparency = 0
                .Solid
            End With
        End With
        ActiveSheet.Shapes(myChart.Parent.Name).Line.Visible = msoFalse
    Else
        MsgBox "Please select your source data first", vbInformation, "Missing Data"
    End If
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
- [[Excel Object Model]]
- [[Error Handling]]
- [[Section 20 - Working with Charts]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
