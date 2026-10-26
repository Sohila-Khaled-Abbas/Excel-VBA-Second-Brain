---
type: project-doc
project: Project 01 - Table of Contents Generator
doc_type: code-notes
status: completed
created: 2026-10-03
updated: 2026-10-26
---

# Project 01: Core VBA Implementation Notes

> [!important]
> **Source Workbook**: `Milestones/TableOfContents/Project_Automatic_TableOfContent.xlsm`  
> **Module**: `AutomaticTableofContent`  
> **Macro Name**: `Auto_Table_Contents`

---

## 💻 Complete Production Code

```vba
Option Explicit

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

## 🔍 Code Breakdown & Engineering Patterns

### 1. Interactive Destination Cell Prompt (`Application.InputBox Type:=8`)
- Prompts the user to click or select any cell on the worksheet:
  ```vba
  Set StartCell = Excel.Application.InputBox( _
      "Where do you want to insert the Table of Contents?" & vbNewLine & "Please select the cell:", _
      "Insert Table of Contents", , , , , , 8)
  ```
- **Error Trapping on Cancel**: When the user clicks Cancel on a `Type:=8` InputBox, Excel generates **Error 424 (Object required)**. The code catches this via `On Error Resume Next` and immediately executes `If Err.Number = 424 Then Exit Sub`.

### 2. Overwrite Protection Warning
- Computes `EndCell` using `Worksheets.Count - 2` to predict the exact bounding box of cells that will be populated.
- Prompts with `vbOKCancel + vbDefaultButton2` so that pressing `Enter` defaults to **Cancel**, preventing accidental data loss.

### 3. Collection Iteration & Filter Gates
- Loops `For Each Sh In Worksheets`.
- **ActiveSheet Guard**: `If ActiveSheet.Name <> ShName Then` prevents creating a self-referential hyperlink to the TOC sheet itself.
- **Visibility Gate**: `If Sh.Visible = xlSheetVisible Then` ensures hidden and very hidden tabs (e.g. calculation models, raw tables) are excluded from public navigation.

### 4. Dynamic Two-Way Hyperlinks with Single-Quote Escaping
- Adds hyperlink anchor at `StartCell`:
  ```vba
  ActiveSheet.Hyperlinks.Add _
      Anchor:=StartCell, _
      Address:="", _
      SubAddress:="'" & ShName & "'!A1", _
      TextToDisplay:=ShName
  ```
- **Critical Single Quotes**: Wrapping `'" & ShName & "'!A1"` ensures sheet names containing spaces, dashes, or special characters do not break with "Reference is not valid" runtime errors.
- **Header Capture**: Extracts the title / description from `Sh.Range("A1").Value` into the adjacent column (`StartCell.Offset(0, 1)`).
- Advances destination pointer: `Set StartCell = StartCell.Offset(1, 0)`.

---

## 🚀 Deployment via Personal Macro Workbook (`PERSONAL.XLSB`)
To make this TOC generator available in **any** Excel workbook on your machine:
1. Record a dummy macro and select store in *"Personal Macro Workbook"*.
2. Open VBE (`Alt + F11`), locate `VBAProject (PERSONAL.XLSB)`.
3. Create a new module and paste `Auto_Table_Contents()`.
4. Add a button to the **Quick Access Toolbar (QAT)** or custom Ribbon tab targeting `PERSONAL.XLSB!Auto_Table_Contents`.
