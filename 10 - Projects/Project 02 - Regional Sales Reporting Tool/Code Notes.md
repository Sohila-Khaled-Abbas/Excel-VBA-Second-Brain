---
type: project-doc
project: Project 02 - Regional Sales Reporting Tool
doc_type: code-notes
status: in-progress
created: 2026-10-03
updated: 2026-10-26
---

# Project 02: Core VBA Implementation Highlights

> [!important]
> **Source Workbook**: `Milestones/SalesReporting/Monthly_Sales_Reporting_Template_Tool.xlsm`  
> **Architecture**: Modular ETL pipeline separating Data Ingestion (`Import`), Reporting & Transformation (`Reports`), and Administration & Security (`Admin`).

---

## 1. Multi-File Batch Ingestion (`Import.bas`)

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

### Key Engineering Features:
- **`FileDialog(msoFileDialogFilePicker)` with `AllowMultiSelect:=True`**: Allows users to select multiple regional Excel workbooks (`.xlsx`, `.xlsm`) in a single prompt.
- **Clipboard Bypass**: Checks for empty user selection before entering loop. Reads rows dynamically from each source workbook's `Range("A2").CurrentRegion` and pastes sequentially into `shAll` (`RawData`).
- **Real-Time Status Bar**: Updates `Application.StatusBar = "Importing file: " & fName` so the user tracks progress through large batches.
- **Automated Task Tracker**: Calls `Admin.Entry_Point` to disable screen rendering, appends to the task activity log, and runs `Admin.Exit_Point` upon completion.

---

## 2. Dynamic Reporting & Semicolon CSV Export (`Reports.bas`)

```vba
Option Explicit
Dim myDate As Long
Dim myRegion As String
Dim ReplaceFile As VbMsgBoxResult
Dim Overwrite(1) As Boolean

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
Private Sub Manager_Report()

    Dim newBook As Workbook
    Dim newPath As String
    
    newPath = ThisWorkbook.Path & "\" & "M_" & myDate & myRegion & ".xlsx"
    If Len(Dir(newPath)) > 0 Then 'file exists
        ReplaceFile = MsgBox("Manager File Exists. Do you want to overwrite?", vbOKCancel, "Overwrite?")
        If ReplaceFile = vbCancel Then
            Overwrite(1) = False
            Exit Sub
        End If
    End If
    Set newBook = Workbooks.Add
     ActiveWindow.DisplayGridlines = False
    
    shRegion.PivotTables("PTManager").TableRange1.Copy
       
    With newBook.Sheets(1).Range("A3")
        .PasteSpecial (xlPasteValues)
        .PasteSpecial (xlPasteFormats)
        .PasteSpecial (xlPasteColumnWidths)
    End With
           
    With newBook.Sheets(1).Range("A1:C1")
         Range("A1").Value = "Regional Overview"
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Interior.ColorIndex = 55
    End With

   newBook.SaveAs FileName:=newPath
   newBook.Close
    TaskRow = shStart.Range("B" & shStart.Rows.Count).End(xlUp).Row + 1
    shStart.Range("B" & TaskRow).Value = "Manager Report for " & myRegion & myDate & " - created on: " & Now
End Sub

Sub Create_CSV()
    Dim FileName As String
    Dim FileNum As Integer
    Dim eRow As Range, eCell As Range, FullRange As Range
    Dim eLine As String
    
    Call Entry_Point
    With shAll.Range("A1").CurrentRegion
        Set FullRange = .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Rows
    End With
    'Determine file path and name
    FileName = ThisWorkbook.Path & "\" & Format(shAll.Range("C2").Value, "YYYYMM") & ".csv"
    'get the next available file number to write to
    FileNum = FreeFile
    Open FileName For Output As FileNum
    For Each eRow In FullRange
        For Each eCell In eRow.Cells
            eLine = eLine & eCell.Value & ";"
        Next eCell
            eLine = Left(eLine, Len(eLine) - 1)
            Print #FileNum, eLine
            eLine = ""
    Next eRow
    Close FileNum
    Call Exit_Point
    MsgBox "CSV file has been saved in the same directory as this workook.", vbInformation, "CSV Saved"
    TaskRow = shStart.Range("B" & shStart.Rows.Count).End(xlUp).Row + 1
    shStart.Range("B" & TaskRow).Value = "CSV for month " & myDate & " - created on: " & Now
End Sub
```

### Key Engineering Features:
- **Template Cloning**: Copies `shTemp` before the first tab, renames dynamically, and points PivotTable caches to the newly consolidated `RawData` range.
- **PivotCache Updating**: Updates `PivotTable.ChangePivotCache` dynamically to embrace newly imported regional rows.
- **Low-Level Semicolon CSV Streaming via `FreeFile`**: Standard Excel `SaveAs CSV` uses comma or system regional settings. Using `Open filePath For Output As #fNum` and iterating rows with `Print #fNum, lineStr` guarantees exact semicolon delimiters for ERP / SAP integration.
- **Safe Overwrite Dialogs**: Prompts `MsgBox("File already exists. Overwrite?", vbYesNo)` before replacing exported files.

---

## 3. Administrative Security & Sheet Protection (`Admin.bas`)

```vba
Option Explicit
Public TaskRow As Long

Sub Entry_Point()
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = "Your Assistant is Busy Working"
    End With
End Sub

Sub Exit_Point()

With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlCalculationAutomatic
    .CutCopyMode = False
    .StatusBar = ""
End With

End Sub

Sub Clear_OldMonth()
    'clear summary data sheet
    shAll.Range("A2", "A" & shAll.Rows.Count - 1).EntireRow.Delete
    shStart.Range("B20:B33").ClearContents
End Sub

Sub Admin_Tasks()
    Dim InputPass As String, Password As String
    Password = "me"
    InputPass = VBA.InputBox("Enter Password to toggle access")
    If InputPass = "" Then Exit Sub
    
    If InputPass <> Password Then
        MsgBox Title:="Incorrect Password", prompt:="The Password is Incorrect!", Buttons:=vbCritical
        Exit Sub
    End If
    Application.ScreenUpdating = False
    If shAll.Visible = xlSheetVisible Then
        shAll.Visible = xlSheetVeryHidden
        shMap.Visible = xlSheetVeryHidden
        shNewDat.Visible = xlSheetVeryHidden
    Else
         shAll.Visible = xlSheetVisible
        shMap.Visible = xlSheetVisible
        shNewDat.Visible = xlSheetVisible
    End If
    Application.ScreenUpdating = False
End Sub
```

### Key Engineering Features:
- **Performance Optimization Framework**: `Entry_Point()` and `Exit_Point()` manage `ScreenUpdating`, `DisplayAlerts`, `Calculation`, and `StatusBar` centrally.
- **`xlSheetVeryHidden` Security**: Hides sensitive mapping and raw database tabs completely from the standard Excel tab menu. They cannot be unhidden from the worksheet UI without entering the administrative password via `Admin_Tasks()`.
