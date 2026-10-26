---
type: lecture
course: Excel VBA & Macros
section: "[[Section 18 - Main Project 3 Invoice Generation Tool]]"
section_number: 18
lecture_number: 9
status: not-started
topics:
  - Main Project #3: Invoice Generation Tool
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-18
---

# Lecture S18-09: Create Email in Outlook & Attach PDF Invoice

> [!important]
> **Section Reference**: [[Section 18 - Main Project 3 Invoice Generation Tool|Section 18: Main Project #3: Invoice Generation Tool]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Create Email in Outlook & Attach PDF Invoice

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Create Email in Outlook & Attach PDF Invoice' within the context of Main Project #3: Invoice Generation Tool.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Create Email in Outlook & Attach PDF Invoice', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Private Sub btEmail_Click()



    Dim EmailAddress As String

    Dim FindCust As Range

    Dim myOutlook As Outlook.Application

    Dim newEmail As Outlook.MailItem

    

    On Error GoTo ErrorHandle

    'look for customer in master data to be able to find email address

    Set FindCust = shMaster.Range("mycustomer").Find(What:=shTemp.Range("B13").Value, lookat:=xlWhole)

    If Not FindCust Is Nothing Then

        EmailAddress = shMaster.Range("H" & FindCust.Row).Value

        If EmailAddress = "" Then

            MsgBox "There is no email address for this customer. Please update the master data. For now create the email yourself."

            Unload Me

            Exit Sub

        End If

        

        'create new Outlook session

        Set myOutlook = New Outlook.Application

        Set newEmail = myOutlook.CreateItem(olMailItem)

        With newEmail

            .To = EmailAddress

            .Subject = "Invoice"

            .Body = "Please find attached my invoice. It was a pleasure working with you."

            .Attachments.Add PDFPath & FileName & ".pdf"

            .Save

            

            '.send

        End With

        'free everything up

        myOutlook.Quit

        Set myOutlook = Nothing

        Set newEmail = Nothing

        MsgBox "Your Email is now saved in your Outlook Draft folder. PDF has been attached.", vbInformation

    Else

        MsgBox "Customer is not in the Masterdata. Sorry, no email can be created", vbInformation, "Master Data Missing"

    End If

    Unload Me

    

    

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
- [[VBA Object Model]]
- [[Error Handling]]
- [[Section 18 - Main Project 3 Invoice Generation Tool]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
