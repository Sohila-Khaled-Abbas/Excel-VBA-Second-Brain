---
type: project-doc
project: Project 03 - Invoice Generation Tool
doc_type: code-notes
status: in-progress
created: 2026-10-03
updated: 2026-10-26
---

# Project 03: Key Code Patterns & Snippets

> [!important]
> **Source Workbook**: `Milestones/InvoiceGeneration/Invoice_Generator.xlsm`  
> **Form Suite**: `FmGetCustomer` (dynamic search), `FmInputMaster` (validation & add), `FmViewMaster` (browse & edit), `FmGetInvoice` (MultiPage PDF & Outlook email dispatch).

---

## 1. Worksheet Event Trigger (`shInvoice`)

```vba
Option Explicit


Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Columns.Count > 25 Then Exit Sub
    'check if customer is entered
    If Not Intersect(Target, Me.ListObjects("InvoiceTable").ListColumns(1).Range) Is Nothing Then
        Set mySelection = Target
        'send the customer range and customer value to the sub function
        Call Find_Cust_Comp(shMaster.Range("mycustomer"), mySelection.Value)
        FmGetCustomer.tbCustomer.Value = mySelection.Value
        Application.EnableEvents = False
        FmGetCustomer.Show
        Application.EnableEvents = True
        'check if company is entered
    ElseIf Not Intersect(Target, Me.ListObjects("InvoiceTable").ListColumns(2).Range) Is Nothing Then
        Set mySelection = Target
        'send company range and company value to the sub function
        Call Find_Cust_Comp(shMaster.Range("mycompany"), mySelection.Value)
        FmGetCustomer.tbCompany.Value = mySelection.Value
        Application.EnableEvents = False
         FmGetCustomer.Show
        Application.EnableEvents = True
    End If
End Sub


Private Sub AXcbCollapse_Click()
    'collapse or show grouped columns
    If Me.AXcbCollapse.Value = True Then
        Me.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    Else
        Me.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    End If
End Sub

Private Sub AXcbHidePaid_Click()
    'filer or unfilter the paid invoices
    If Me.AXcbHidePaid.Value = True Then
      Me.ListObjects("InvoiceTable").Range.AutoFilter Field:=24, Criteria1:="<>*paid*"
    Else
      Me.ListObjects("InvoiceTable").Range.AutoFilter
    End If
End Sub
```

### Key Engineering Features:
- **`Worksheet_Change` Event**: Intercepts edits inside the structured `InvoiceTable`.
- **Target Intersection**: Uses `Intersect(Target, Me.ListObjects("InvoiceTable").ListColumns(1).Range)` to detect when a customer name or company is typed.
- **Event Suppression**: Disables `Application.EnableEvents = False` before showing `FmGetCustomer` to prevent cascading event triggers.

---

## 2. Dynamic Real-Time Searchable ListBox (`FmGetCustomer`)

```vba
Option Explicit

Private Sub btCancel_Click()
    Unload Me
End Sub

Private Sub btSelect_Click()
    Dim FindCust As Range
    If Me.LbCustomer.ListIndex = -1 Then
        MsgBox "Please select a a name from the list"
        Exit Sub
    Else
        Set FindCust = shMaster.Range("mycustomer").Find(What:=Me.LbCustomer.Value, lookat:=xlWhole)
        If Not FindCust Is Nothing Then
            shMaster.Range("A" & FindCust.Row, "G" & FindCust.Row).Copy
            shInvoice.Range("B" & mySelection.Row).PasteSpecial (xlPasteValues)
         'add instruction text for last line
        
        shInvoice.Range("B" & mySelection.Row + 1).Value = "Type & Enter to select customer or company"
        shInvoice.Range("Y" & mySelection.Row).Copy
        shInvoice.Range("Y" & mySelection.Row + 1).PasteSpecial xlPasteValidation
        Application.CutCopyMode = False
        Range("A" & mySelection.Row + 1).Select
           
        End If
    End If
    Unload Me
End Sub

Private Sub tbCustomer_AfterUpdate()
    Me.LbCustomer.Clear
    Call Find_Cust_Comp(shMaster.Range("mycustomer"), Me.tbCustomer.Value)
End Sub

Private Sub tbCompany_AfterUpdate()
    Me.LbCustomer.Clear
    Call Find_Cust_Comp(shMaster.Range("mycompany"), Me.tbCompany.Value)
End Sub
```

### Key Engineering Features:
- **`InStr` Filtering**: Searches the master customer column for character matches as the user types, populating matching candidates dynamically into `LbCustomer`.
- **One-Click Data Injection**: Double-clicking or selecting a customer copies all metadata (address, VAT number, email) into the active invoice row in a single shot.

---

## 3. Dynamic Input Validation (`FmInputMaster`)

```vba
Option Explicit

Private Sub btAdd_Click()
    Dim ctr As Control
    Dim InpReq As Boolean
    Dim rMaster As Long
    
    On Error GoTo ErrorHandle
    '1. Check if required fields are not empty & VAT is numeric
    InpReq = True
    For Each ctr In Me.Controls
        If TypeName(ctr) = "TextBox" Then
            If ctr.Tag = "required" Then
                If ctr.Text = "" Then
                    InpReq = False
                End If
            End If 'tag
        End If 'typename
    Next ctr
    If InpReq = False Or IsNumeric(Me.tbVAT) = False Then
        MsgBox "Please input the required fields which are marked with ""*""" & vbNewLine & _
    "Please ensure VAT is a Number Value", vbCritical, "Input incomplete"
        Me.tbCustName.SetFocus
        Exit Sub
    End If
    '2. Add data to Master Data
    With shMaster
        If Me.btAdd.Caption = "UPDATE" Then
            rMaster = EditNum
        Else
        rMaster = .ListObjects("CustomerTable").Range.Rows.Count + 1
        End If
        .Range("A" & rMaster) = Me.tbCustName
        .Range("B" & rMaster) = Me.tbCompName
        .Range("C" & rMaster) = Me.tbAddress1
        .Range("D" & rMaster) = Me.tbAddress2
        .Range("E" & rMaster) = Me.tbAddress3
        .Range("F" & rMaster) = Me.tbUID
        .Range("G" & rMaster) = Me.tbVAT / 100
        .Range("H" & rMaster) = Me.tbEmail

    
    End With
    Unload Me
    MsgBox "Master data has been updated", vbInformation, "Updated!"
    Exit Sub
    
ErrorHandle:
    MsgBox "It looks like an error has occurred! Please make sure you have followed the instructions.", vbCritical
    Unload Me

End Sub

Private Sub btCancel_Click()
    Unload Me
End Sub
```

### Key Engineering Features:
- **`.Tag = "required"` Dynamic Check**: Loops through all form controls; any `TextBox` tagged as `"required"` is validated for non-empty text.
- **Numeric Validation**: Validates VAT with `IsNumeric(Me.tbVAT)` before committing to the Excel Table.
- **Dual Mode (ADD vs UPDATE)**: Adapts button caption and logic (`rMaster = EditNum` vs `rMaster = lastRow + 1`) depending on whether the user is inserting a new client or editing an existing record.

---

## 4. MultiPage Invoice Generator & Outlook Email Automation (`FmGetInvoice`)

```vba
' Key Excerpt: PDF Export & Outlook Mail Generation
Private Sub btEmail_Click()
    Dim EmailAddress As String
    Dim FindCust As Range
    Dim myOutlook As Outlook.Application
    Dim newEmail As Outlook.MailItem
    
    On Error GoTo ErrorHandle
    Set FindCust = shMaster.Range("mycustomer").Find(What:=shTemp.Range("B13").Value, lookat:=xlWhole)
    If Not FindCust Is Nothing Then
        EmailAddress = shMaster.Range("H" & FindCust.Row).Value
        If EmailAddress = "" Then
            MsgBox "No email address found for this customer.", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Generate PDF
    PDFPath = ThisWorkbook.Path & "\Invoices\" & FileName & ".pdf"
    shTemp.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PDFPath, Quality:=xlQualityStandard
    
    ' Create Outlook Email
    Set myOutlook = New Outlook.Application
    Set newEmail = myOutlook.CreateItem(olMailItem)
    With newEmail
        .To = EmailAddress
        .Subject = "Invoice " & FileName & " from Company"
        .Body = "Dear Customer," & vbNewLine & vbNewLine & "Please find attached your invoice."
        .Attachments.Add PDFPath
        .Display
    End With
CleanExit:
    Exit Sub
ErrorHandle:
    MsgBox "Error generating invoice email: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

### Key Engineering Features:
- **`ExportAsFixedFormat`**: Exports the populated invoice sheet directly to a polished PDF document.
- **Outlook Integration**: Attaches the PDF and populates recipient address from master data.
- **Early/Late Binding Flexibility**: Works with Early Binding (`New Outlook.Application`) or Late Binding (`CreateObject("Outlook.Application")`) for zero-reference portability.
