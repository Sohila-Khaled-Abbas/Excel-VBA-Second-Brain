---
type: lecture
course: Excel VBA & Macros
section: "[[Section 18 - Main Project 3 Invoice Generation Tool]]"
section_number: 18
lecture_number: 2
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

# Lecture S18-02: Setting up an Input UserForm to Collect Master Data

> [!important]
> **Section Reference**: [[Section 18 - Main Project 3 Invoice Generation Tool|Section 18: Main Project #3: Invoice Generation Tool]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Setting up an Input UserForm to Collect Master Data

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Setting up an Input UserForm to Collect Master Data' within the context of Main Project #3: Invoice Generation Tool.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Setting up an Input UserForm to Collect Master Data', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
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
