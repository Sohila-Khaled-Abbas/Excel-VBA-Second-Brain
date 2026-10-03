---
type: project-doc
project: Project 03 - Invoice Generation Tool
doc_type: code-notes
status: in-progress
created: 2026-10-03
updated: 2026-10-03
---

# Project 03: Key Code Patterns & Snippets

### 1. Dynamic Searchable ListBox Filtering
```vba
Private Sub txtSearch_Change()
    Dim i As Long
    Dim searchStr As String
    
    searchStr = LCase(Me.txtSearch.Text)
    Me.lstCustomers.Clear
    
    For i = 1 To UBound(mMasterCustomers, 1)
        If InStr(1, LCase(mMasterCustomers(i, 2)), searchStr) > 0 Then
            With Me.lstCustomers
                .AddItem mMasterCustomers(i, 1) ' ID
                .List(.ListCount - 1, 1) = mMasterCustomers(i, 2) ' Company Name
                .List(.ListCount - 1, 2) = mMasterCustomers(i, 3) ' Contact Name
                .List(.ListCount - 1, 3) = mMasterCustomers(i, 4) ' Email
            End With
        End If
    Next i
End Sub
```

### 2. PDF & Outlook Email Generation
```vba
Public Sub GenerateAndEmailInvoice(wsInvoice As Worksheet, custEmail As String, invNumber As String)
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Invoices\Invoice_" & invNumber & ".pdf"
    
    ' Export PDF
    wsInvoice.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard
        
    ' Late-binding Outlook email
    Dim outApp As Object, outMail As Object
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)
    
    With outMail
        .To = custEmail
        .Subject = "Invoice #" & invNumber & " from Automation Corp"
        .Body = "Dear Customer," & vbCrLf & vbCrLf & _
                "Please find attached Invoice #" & invNumber & " for services rendered." & vbCrLf & _
                "Thank you for your business!"
        .Attachments.Add pdfPath
        .Display ' Show to user for validation
    End With
    
    Set outMail = Nothing
    Set outApp = Nothing
End Sub
```
