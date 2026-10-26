---
type: lecture
course: Excel VBA & Macros
section: "[[Section 14 - Main Project #2: Regional Reporting Tool]]"
section_number: 14
lecture_number: 3
status: not-started
topics:
  - Main Project #2: Regional Reporting Tool
related_concepts:
  - '[[VBA Syntax]]'
  - '[[Excel Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-14
---

# Lecture S14-03: Importing Data from Multiple Files using MultiSelect

> [!important]
> **Section Reference**: [[Section 14 - Main Project #2: Regional Reporting Tool|Section 14: Main Project #2: Regional Reporting Tool]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Importing Data from Multiple Files using MultiSelect

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Importing Data from Multiple Files using MultiSelect' within the context of Main Project #2: Regional Reporting Tool.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Importing Data from Multiple Files using MultiSelect', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Import_Sales()

    '1. First import data 1-to-1 to NewData tab
    '2. Use Find function to find the matching header from Summary tab in NewData tab
    '3. Copy and paste the required column from NewData in Summary tab

    Dim shSales As Worksheet
    Dim FileToOpen As Variant
    Dim FileCnt As Byte, c As Byte
    Dim LastRow As Long, LastTempRow As Long
    Dim SelectedBook As Workbook
    Dim GetHeader As Range
    Const StartRowTemp = 4  ' if start row on company templates change, update this line
    Const CompanyRng As String = "C2"
    Dim CompanyImported() As String
    Call Entry_Point
 
    On Error GoTo Handle:
    
    'Pick the files to import - allow multiselect
    FileToOpen = Application.GetOpenFilename(Filefilter:="Excel Files (*.xlsx), *.xlsx", Title:="Select Workbook to Import", MultiSelect:=True)
        
        If IsArray(FileToOpen) Then
            'how many files were selected
            
            For FileCnt = 1 To UBound(FileToOpen)
            
                'clean up the data on temporary sheet that were already imported
                shNewDat.Cells.Clear
                'locate last empty row in summary tab
                LastRow = shAll.Cells(Rows.Count, 1).End(xlUp).Row + 1
                Set SelectedBook = Workbooks.Open(FileName:=FileToOpen(FileCnt))
                
                'in case this sheet doesn't exist, error handle takes care of it by informing the user
                SelectedBook.Worksheets("Sales").Cells.Copy
                
                shNewDat.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
                 
                 SelectedBook.Close
                 
                'locate last row in the temp sheet
                LastTempRow = shNewDat.Cells(Rows.Count, 2).End(xlUp).Row
                
                'assign company name
                'calculate the number of rows that should be copied over and put the company code there
                shAll.Range("A" & LastRow, "A" & LastRow + (LastTempRow - StartRowTemp) - 1).Value = shNewDat.Range(CompanyRng).Value
              '  Debug.Print LastRow + LastTempRow - StartRowTemp - 1
                c = 2  'columns needed for the summary sheet
                'use find function to match headers, since the template could have a different order of columns than the summary tab
                Do While shAll.Cells(1, c) <> ""
                    Set GetHeader = shNewDat.Rows(StartRowTemp).Find(what:=shAll.Cells(1, c).Value, LookIn:=xlValues, MatchCase:=False, lookat:=xlWhole)
                    If Not GetHeader Is Nothing Then
                        shNewDat.Range(shNewDat.Cells(StartRowTemp + 1, GetHeader.Column), shNewDat.Cells(LastTempRow, GetHeader.Column)).Copy
                        shAll.Cells(LastRow, c).PasteSpecial
                    End If
                    c = c + 1 'check and copy next column
               Loop
                ReDim Preserve CompanyImported(1 To FileCnt)
                CompanyImported(FileCnt) = shNewDat.Range("C2").Value
            Next FileCnt
              
            MsgBox "Data Imported Sucessfully", vbInformation, "Good Job"
         
        End If

     shStart.Select
    Range("A1").Select
    TaskRow = shStart.Range("B" & shStart.Rows.Count).End(xlUp).Row + 1
    shStart.Range("B" & TaskRow).Value = "Data imported for: " & Join(CompanyImported, ",")
    
    
Call Exit_Point

Exit Sub

Handle:
If Err.Number = 9 Then
MsgBox "It looks like you have selected the wrong file."
Else
MsgBox "An error has occurred"
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
- [[Excel Object Model]]
- [[Error Handling]]
- [[Section 14 - Main Project #2: Regional Reporting Tool]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
