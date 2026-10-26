---
type: lecture
course: Excel VBA & Macros
section: "[[Section 11 - Working with Arrays]]"
section_number: 11
lecture_number: 2
status: not-started
topics:
  - Working with Arrays
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-11
---

# Lecture S11-02: One Dimensional Arrays (and transferring back to sheet)

> [!important]
> **Section Reference**: [[Section 11 - Working with Arrays|Section 11: Working with Arrays]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: One Dimensional Arrays (and transferring back to sheet)

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'One Dimensional Arrays (and transferring back to sheet)' within the context of Working with Arrays.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'One Dimensional Arrays (and transferring back to sheet)', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Array_OneDimension()



    Dim MonthArray(1 To 12) As String

    Dim i As Byte

    Dim r As Long

    

    'Fill up the month array

    For i = 1 To 12

        'MonthArray(i) = Range("A" & 4 + i).Value

        'my months is a named range

        MonthArray(i) = Range("mymonths").Cells(i, 1).Value

    Next i

        'example of writing the array back to Excel

        Range("C5:N5").Value = MonthArray

        'example of writing the array back vertically - it needs to be transposed

        Range("C5:C16").Value = Application.WorksheetFunction.Transpose(MonthArray)

        

       ' MsgBox "Array element for position 3 is: " & MonthArray(3)



        'to create a report that leaves out 4 rows between each month

        For i = LBound(MonthArray) To UBound(MonthArray)

            Range("E7").Offset(r).Value = MonthArray(i)

            r = r + 5

        Next i

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
- [[Section 11 - Working with Arrays]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
