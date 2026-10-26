---
type: lecture
course: Excel VBA & Macros
section: "[[Section 18 - Main Project #3: Invoice Generation Tool]]"
section_number: 18
lecture_number: 8
status: not-started
topics:
  - Main Project #3: Invoice Generation Tool
related_concepts:
  - '[[VBA Syntax]]'
  - '[[Excel Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-18
---

# Lecture S18-08: Export PDF & Excel Workbook for Each Invoice based on ListBox Selection

> [!important]
> **Section Reference**: [[Section 18 - Main Project #3: Invoice Generation Tool|Section 18: Main Project #3: Invoice Generation Tool]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Export PDF & Excel Workbook for Each Invoice based on ListBox Selection

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Export PDF & Excel Workbook for Each Invoice based on ListBox Selection' within the context of Main Project #3: Invoice Generation Tool.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Export PDF & Excel Workbook for Each Invoice based on ListBox Selection', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Private Sub btInvoice_Click()

    Dim ExcelPath As String
    Dim r As Long
    Dim PosTextRow As Byte
    Dim InvTable As ListObject
    Dim NextInvoice As Variant
    Dim NewBook As Workbook
    
    On Error GoTo ErrorHandle
    
    '1. Check if Folders exist
    ExcelPath = ThisWorkbook.Path & "\" & Range("Excel_Folder").Value
    PDFPath = ThisWorkbook.Path & "\" & Range("PDF_Folder").Value
    
    If Dir(ExcelPath, vbDirectory) = VBA.Constants.vbNullString Or Dir(PDFPath, vbDirectory) = VBA.Constants.vbNullString Then
        MsgBox "The Excel & PDF folders do not exist. Please create these in the same folder as this workbook first.", vbInformation, "Folder does not exist"
        Exit Sub
    End If
    '2. Figure out what has been selected in the List Box
    If Me.lbOpenInvoice.ListIndex = -1 Then
        MsgBox "Please select a an invoice from the list"
        Exit Sub
    Else
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        PosTextRow = InStr(1, Me.lbOpenInvoice.Value, " - Table Row: ", vbTextCompare) + Len(" - Table Row: ")
        r = CLng(Mid(Me.lbOpenInvoice.Value, PosTextRow))
        
    '3. Update Invoice Table by assigning a new Invoice number if missing
        Set InvTable = shInvoice.ListObjects("InvoiceTable")
        'check if there is at least a number in amount 1 before creating the invoice

        If InvTable.ListColumns("Amount Line 1").DataBodyRange(r).Value = "" Then
            MsgBox "There is no value to invoice yet."
            Unload Me
            Exit Sub
        End If

        NextInvoice = InvTable.ListColumns("Invoice Number").DataBodyRange(r).Value
        If NextInvoice = "" Then NextInvoice = Application.WorksheetFunction.Max(InvTable.ListColumns("Invoice Number").DataBodyRange) + 1
        InvTable.ListColumns("Invoice Number").DataBodyRange(r).Value = NextInvoice
        
    '4. Copy fields from Invoice table to template tab
        shTemp.Range("D7").Value = NextInvoice
        shTemp.Range("D8").Value = Date
        shTemp.Range("B12").Value = InvTable.ListColumns("Company").DataBodyRange(r).Value
        shTemp.Range("B13").Value = InvTable.ListColumns("Customer").DataBodyRange(r).Value
        shTemp.Range("B15").Value = InvTable.ListColumns("Address Line 1").DataBodyRange(r).Value
        shTemp.Range("B16").Value = InvTable.ListColumns("Address Line 2").DataBodyRange(r).Value
        shTemp.Range("B17").Value = InvTable.ListColumns("Address Line 3").DataBodyRange(r).Value
        shTemp.Range("B21").Value = InvTable.ListColumns("Services Line 1").DataBodyRange(r).Value
        shTemp.Range("C21").Value = InvTable.ListColumns("Date Line 1").DataBodyRange(r).Value
        shTemp.Range("D21").Value = InvTable.ListColumns("Amount Line 1").DataBodyRange(r).Value
        shTemp.Range("B22").Value = InvTable.ListColumns("Services Line 2").DataBodyRange(r).Value
        shTemp.Range("C22").Value = InvTable.ListColumns("Date Line 2").DataBodyRange(r).Value
        shTemp.Range("D22").Value = InvTable.ListColumns("Amount Line 2").DataBodyRange(r).Value
        shTemp.Range("B23").Value = InvTable.ListColumns("Services Line 3").DataBodyRange(r).Value
        shTemp.Range("C23").Value = InvTable.ListColumns("Date Line 3").DataBodyRange(r).Value
        shTemp.Range("D23").Value = InvTable.ListColumns("Amount Line 3").DataBodyRange(r).Value
        'VAT
        shTemp.Range("E21").Value = InvTable.ListColumns("VAT").DataBodyRange(r).Value


        'make sure service text is fully visible
         shTemp.Rows("21:23").EntireRow.AutoFit
    
    '5. Create a new Excel Workbook for the invoice selected
        Set NewBook = Workbooks.Add
        shTemp.Copy before:=NewBook.Sheets(1)
        NewBook.ActiveSheet.Name = CStr(NextInvoice)
        
    '6. Create a PDF Version of the Invoice
        FileName = "\" & NextInvoice & "_" & shTemp.Range("B12").Value
        NewBook.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=PDFPath & FileName
        NewBook.SaveAs FileName:=ExcelPath & FileName
        NewBook.Close
        
    '7. Add Date and update status to "created" in Invoice Table
        InvTable.ListColumns("Invoice Date").DataBodyRange(r).Value = Date
        InvTable.ListColumns("Status").DataBodyRange(r).Value = "Created"
    
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        MsgBox "Invoice for " & Me.lbOpenInvoice.Value & " has been created!" & vbNewLine _
        & vbNewLine & "Excel file is saved here: " & ExcelPath & FileName & vbNewLine _
        & vbNewLine & "PDF file is saved here: " & PDFPath & FileName, vbInformation, "Invoice Created!"

    '8. Move on to the second page of Multipage control
        Me.MPInvoice.Value = 1
    End If
        Exit Sub
ErrorHandle:
    MsgBox "It looks like an error has occurred! Please make sure you have followed the instructions.", vbCritical
    Unload Me
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
- [[Section 18 - Main Project #3: Invoice Generation Tool]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
