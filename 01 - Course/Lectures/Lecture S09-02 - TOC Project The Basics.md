---
type: lecture
course: Excel VBA & Macros
section: "[[Section 09 - Main Project 1 Create a Tool to Automate Table of Contents]]"
section_number: 9
lecture_number: 2
status: not-started
topics:
  - Main Project #1: Create a Tool to Automate Table of Contents
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-09
---

# Lecture S09-02: TOC Project: The Basics

> [!important]
> **Section Reference**: [[Section 09 - Main Project 1 Create a Tool to Automate Table of Contents|Section 09: Main Project #1: Create a Tool to Automate Table of Contents]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: TOC Project: The Basics

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'TOC Project: The Basics' within the context of Main Project #1: Create a Tool to Automate Table of Contents.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'TOC Project: The Basics', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Auto_Table_Contents()

'this procedure adds a table of contents to your active sheet. It creates a hyperlink to each sheet and takes the header from cell A1 of every sheet



    Dim Sh As Worksheet

    Dim StartCell As Range 'for inputbox

    Dim ShName As String

    Dim MsgConfirm As VBA.VbMsgBoxResult

    Dim EndCell As Range ' for msgbox information

    

    'prompt the user to select a cell

    On Error Resume Next

    Set StartCell = Excel.Application.InputBox("Where do you want to insert the Table of Contents?" & vbNewLine & "Please select the cell:" _

                                , "Insert Table of Contents", , , , , , 8)

    If Err.Number = 424 Then Exit Sub

    On Error GoTo Handle:

                         

    Set StartCell = StartCell.Cells(1, 1)

    Set EndCell = StartCell.Offset(Worksheets.Count - 2, 1)

    

    'allow the user to cancel otherwise their cells might be overwritten

    MsgConfirm = MsgBox("The values in cells: " & vbNewLine & _

    StartCell.Address & " to " & EndCell.Address & " could be overwritten." _

    & vbNewLine & "Would you like to continue?", vbOKCancel + vbDefaultButton2, "Confirmation Required")

    If MsgConfirm = vbCancel Then Exit Sub

    

    For Each Sh In Worksheets

        ShName = Sh.Name

         StartCell = ShName

        'include only if sheet is not activesheet and sheet is visible

        If ActiveSheet.Name <> ShName Then

            If Sh.Visible = xlSheetVisible Then

                ActiveSheet.Hyperlinks.Add Anchor:=StartCell, Address:="", SubAddress:= _

                 "'" & ShName & "'!A1", TextToDisplay:=ShName

                StartCell.Offset(0, 1).Value = Sh.Range("A1").Value

                Set StartCell = StartCell.Offset(1, 0)

            End If 'if sheet is visible

        End If ' if sheet <> activesheet

    Next Sh



    Exit Sub

Handle:

MsgBox "Unfortunately, an error has occured!"



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
- [[Section 09 - Main Project 1 Create a Tool to Automate Table of Contents]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
