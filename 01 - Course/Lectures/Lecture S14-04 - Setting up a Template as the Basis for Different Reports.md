---
type: lecture
course: Excel VBA & Macros
section: "[[Section 14 - Main Project 2 Regional Reporting Tool]]"
section_number: 14
lecture_number: 4
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

# Lecture S14-04: Setting up a Template as the Basis for Different Reports

> [!important]
> **Section Reference**: [[Section 14 - Main Project 2 Regional Reporting Tool|Section 14: Main Project #2: Regional Reporting Tool]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Setting up a Template as the Basis for Different Reports

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Setting up a Template as the Basis for Different Reports' within the context of Main Project #2: Regional Reporting Tool.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Setting up a Template as the Basis for Different Reports', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Create_Reports()

    Dim tbTemp As ListObject

    Dim RowsTb As Long, LastRow As Long

    

    Call Entry_Point

    Overwrite(0) = True

    Overwrite(1) = True

    On Error GoTo Handle

    'create an object variable for the table inside Region Temp tab

    Set tbTemp = shRegion.ListObjects("TableTemp")

    RowsTb = tbTemp.DataBodyRange.Rows.Count

    'Remove existing data in regional table except for line 1 if there is more than 1 line of data

    If RowsTb > 1 Then tbTemp.DataBodyRange.Rows("2:" & RowsTb).Delete

    LastRow = shAll.Range("A" & Rows.Count).End(xlUp).Row

    myRegion = shStart.AxRegion.Value

    myDate = Format(shAll.Range("C2").Value, "YYYYMM")

    

    ''''''''''''version 1

'    Dim Start

'    Start = Timer

'    Dim i As Long, r As Long, LastComp As Long

'    Dim c As Byte

''

''            Loop to find match for region

'            LastComp = shMap.Range("MCompany").Rows.Count

'            r = 2

'            For i = 2 To LastRow

'

'                For c = 1 To LastComp

'

'                    If shAll.Range("A" & i).Value = shMap.Range("Mcompany").Cells(c, 1).Value Then

'                        If UCase(myRegion) = UCase(shMap.Range("Mcompany").Cells(c, 4).Value) Then

'                            shAll.Range("A" & i, "H" & i).Copy shRegion.Range("A" & r, "H" & r)

'                            r = r + 1

'                        End If

'

'                    End If

'                Next c

'            Next i

            

     ''''''''''''version 2

    'find the region for each company for the summary tab

    shAll.Range("I2").FormulaR1C1 = "=VLOOKUP(RC[-8],MCompany,4,FALSE)"

    shAll.Range("I2").AutoFill Destination:=shAll.Range("I2:I" & LastRow)

    Application.Calculate



    With shAll.Range("A1:I" & LastRow)

        .AutoFilter Field:=9, Criteria1:=myRegion

        .Offset(1).Resize(.Rows.Count - 1, .Columns.Count - 1).SpecialCells(xlCellTypeVisible).Copy

    End With



    'paste the values

    shRegion.Range("A2").PasteSpecial

     shAll.Range("A1").AutoFilter

     shAll.Columns("I").Delete

'    Debug.Print Timer - Start

    shRegion.Range("U1").Value = myDate

    Application.Calculate

    shRegion.PivotTables("PTRegion").PivotCache.Refresh

    Call Region_Report

    Call Manager_Report

    If Overwrite(0) = False And Overwrite(1) = False Then

        MsgBox "No filese were created.", , "Operation Cancelled"

    Else

        MsgBox "File(s) created in the same directory as this workbook.", , "Well Done"



    End If

    Call Exit_Point

    Exit Sub

    

Handle:

            If Err.Number = 1004 Then

            MsgBox "It seems the data for this region has not been uploaded yet! Please upload and try again."

            shAll.Range("A1").AutoFilter

            shAll.Columns("I").Delete

            Exit Sub

        Else

            MsgBox "Looks like no data has been imported yet.", vbInformation

            Exit Sub

        End If

    Call Exit_Point

    

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
