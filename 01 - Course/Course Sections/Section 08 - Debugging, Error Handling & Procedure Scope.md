---
type: course-section
section_number: 8
section_title: Debugging, Error Handling & Procedure Scope
lecture_count: 9
status: not-started
tags:
  - course-section
  - vba
  - section-08
---

# Section 08: Debugging, Error Handling & Procedure Scope

> [!important]
> **Curriculum Pillar**: Section 8 contains **9 lectures**.  
> **Status**: `NOT-STARTED`  
> **Course Practice Files**: `SectionFiles/S8_ErrorHandling.xlsm`, `S8_ErrorHandling_Start.xlsm`

---

## 🎯 Section Objectives & Automation Capabilities
- Master professional VBE debugging tools: Breakpoints (`F9`), Step Into (`F8`), Step Over (`Shift+F8`), and Run to Cursor (`Ctrl+F8`).
- Harness the **Immediate Window** (`Ctrl+G`), **Locals Window**, and **Watch Window** for state inspection.
- Implement defensive Error Handling architecture using `On Error GoTo`, `On Error Resume Next`, and `Err` object inspection.
- Supercharge macro execution speed by disabling screen rendering, automatic calculation, and alerts.
- Structure modular code: Procedure Scope (`Public` vs `Private`), calling Subs, and passing arguments `ByVal` vs `ByRef`.

---

## 🧠 Architectural Takeaways & Mental Models
- **Defensive Error Handling Framework**:
  ```vba
  Sub RobustProcedure()
      On Error GoTo ErrorHandler
      ' ... execution code ...
  CleanExit:
      ' Restore application settings here!
      Application.ScreenUpdating = True
      Exit Sub
  ErrorHandler:
      MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
      Resume CleanExit
  End Sub
  ```
- **The Performance Trinity**:
  ```vba
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  Application.Calculation = xlCalculationManual
  ```
  *(Always restore all three in the `CleanExit` block!)*
- **`ByVal` vs `ByRef`**: VBA passes arguments `ByRef` by default (giving the called Sub access to modify the original variable). Use `ByVal` when you want to pass a safe, isolated copy of the value.

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- Leaving `On Error Resume Next` active across an entire procedure hides fatal errors, corrupting data silently. Scope it tightly and reset immediately with `On Error GoTo 0` or `On Error GoTo Handler`.
- Forgetting to place an `Exit Sub` right before the `ErrorHandler:` label causes the error handler to execute on every single successful run!

---

## 📚 Lecture Index

- [ ] [[Lecture S08-01 - Overview Debugging, Handling Errors & Procedure Scope|Lecture 01: Overview: Debugging, Handling Errors & Procedure Scope]]
- [ ] [[Lecture S08-02 - Debugging Options Watch, Locals, Immediate Windows & More|Lecture 02: Debugging Options: Watch, Locals, Immediate Windows & More]]
- [ ] [[Lecture S08-03 - Error Handling Different Methods|Lecture 03: Error Handling: Different Methods]]
- [ ] [[Lecture S08-04 - Faster Code by Suppressing Pop-ups & Flickering Screen|Lecture 04: Faster Code by Suppressing Pop-ups & Flickering Screen]]
- [ ] [[Lecture S08-05 - Procedure Scope & Running one Sub from Another|Lecture 05: Procedure Scope & Running one Sub from Another]]
- [ ] [[Lecture S08-06 - Passing Arguments to Sub Procedures (By Ref, By Val)|Lecture 06: Passing Arguments to Sub Procedures (By Ref, By Val)]]
- [ ] [[Lecture S08-07 - Activity Get the Total Number of Formulas (or Comments) Used in a Workbook|Lecture 07: Activity: Get the Total Number of Formulas (or Comments) Used in a Workbook]]
- [ ] [[Lecture S08-08 - Project Get the Total Number of Formulas (or Comments) Used in a Workbook|Lecture 08: Project: Get the Total Number of Formulas (or Comments) Used in a Workbook]]
- [ ] [[Lecture S08-09 - Key Takeaways Debugging, Error Handling & Course Map Status|Lecture 09: Key Takeaways Debugging, Error Handling & Course Map Status]]

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
