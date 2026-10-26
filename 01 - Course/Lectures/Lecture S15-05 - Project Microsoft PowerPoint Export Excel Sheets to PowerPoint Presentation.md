---
type: lecture
course: Excel VBA & Macros
section: "[[Section 15 - Interacting with Other Applications (Word, PowerPoint etc.)]]"
section_number: 15
lecture_number: 5
status: not-started
topics:
  - Interacting with Other Applications (Word, PowerPoint etc.)
related_concepts:
  - '[[VBA Syntax]]'
  - '[[VBA Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-15
---

# Lecture S15-05: Project Microsoft PowerPoint: Export Excel Sheets to PowerPoint Presentation

> [!important]
> **Section Reference**: [[Section 15 - Interacting with Other Applications (Word, PowerPoint etc.)|Section 15: Interacting with Other Applications (Word, PowerPoint etc.)]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: Project Microsoft PowerPoint: Export Excel Sheets to PowerPoint Presentation

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'Project Microsoft PowerPoint: Export Excel Sheets to PowerPoint Presentation' within the context of Interacting with Other Applications (Word, PowerPoint etc.).

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'Project Microsoft PowerPoint: Export Excel Sheets to PowerPoint Presentation', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Create_PowerPoint_Slides()

    Dim PApp As PowerPoint.Application

    Dim PPres As PowerPoint.Presentation

    Dim PSlide As PowerPoint.Slide

    Dim PShape As PowerPoint.Shape

    Dim PLayout As PowerPoint.CustomLayout

    Dim toCopy As Range

    Dim Sh As Worksheet

    Dim i As Long 'slide counter

    

    Application.ScreenUpdating = False

    On Error Resume Next

    Set PApp = GetObject(, "Powerpoint.Application")

    On Error GoTo 0

    'open a new powerpoint session if not open

    If PApp Is Nothing Then

        Set PApp = New PowerPoint.Application

    End If

    ' get presentation

    If PApp.Presentations.Count = 0 Then

        Set PPres = PApp.Presentations.Add

        i = 0

    Else

        Set PPres = PApp.ActivePresentation

        i = PPres.Slides.Count

    End If

    Set PLayout = PPres.SlideMaster.CustomLayouts(7)

    

    For Each Sh In ThisWorkbook.Worksheets

        If Sh.Index > 7 Then

            Set toCopy = Sh.UsedRange

            Set PSlide = PPres.Slides.AddSlide(i + 1, PLayout)

            DoEvents

            toCopy.CopyPicture xlScreen, xlPicture

            PSlide.Shapes.Paste

            DoEvents

            Set PShape = PSlide.Shapes(1)

            With PShape

                .LockAspectRatio = msoTrue

                .Left = 40

                .Top = 40

                .Width = 850

            End With

            i = i + 1

        End If

    Next Sh

    MsgBox "Your Slides have been created."

   ' PApp.Visible = True

     Application.ScreenUpdating = True

    Set PApp = Nothing

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
- [[Section 15 - Interacting with Other Applications (Word, PowerPoint etc.)]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
