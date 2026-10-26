---
type: course-section
section_number: 19
section_title: Creating Customized Functions / Formulas
lecture_count: 8
status: not-started
tags:
  - course-section
  - vba
  - section-19
---

# Section 19: Creating Customized Functions / Formulas

> [!important]
> **Curriculum Pillar**: Section 19 contains **8 lectures**.  
> **Status**: `NOT-STARTED`  
> **Course Practice Files**: `SectionFiles/S19_FunctionProcedures.xlsm`, `S19_FunctionProcedures_Start.xlsx`

---

## 🎯 Section Objectives & Automation Capabilities
- Understand the purpose and mechanics of **User Defined Functions (UDFs)**.
- Differentiate between `Sub` procedures (actions) and `Function` procedures (returning a calculated value).
- Create custom functions that can be used directly in normal worksheet formulas (`=MyFunction()`).
- Implement optional arguments (`Optional Txt As String = "Default"`) and handle data types flexibly.
- Build practical business UDFs: exact Age calculation, splitting text from numbers, and `SumByColor`.

---

## 🧠 Architectural Takeaways & Mental Models
- **The Golden Rule of UDFs**: A Function called from a worksheet formula **can only return a value to its host cell**. It **CANNOT** modify other cells, change formatting, add sheets, delete rows, or alter Excel environment settings!
- **Function Syntax & Return Value**: The return value is assigned directly to the name of the function:
  ```vba
  Function CalcCommission(sales As Double) As Double
      If sales > 10000 Then
          CalcCommission = sales * 0.1
      Else
          CalcCommission = sales * 0.05
      End If
  End Function
  ```
- **Volatile Functions**: By default, UDFs only recalculate when their direct parameter cells change. Adding `Application.Volatile True` forces recalculation on every sheet calculation (use sparingly to avoid performance drag!).

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- Attempting to format cells (e.g. `Target.Interior.Color = ...`) inside a worksheet-called UDF will fail silently.
- Overusing `Application.Volatile` across hundreds of formulas will severely degrade workbook calculation performance.

---

## 📚 Lecture Index

- [ ] [[Lecture S19-01 - Section Overview Function Procedures|Lecture 01: Section Overview: Function Procedures]]
- [ ] [[Lecture S19-02 - Setting up a Simple Function (including Optional Arguments)|Lecture 02: Setting up a Simple Function (including Optional Arguments)]]
- [ ] [[Lecture S19-03 - Function to Calculate Age (Result in Years & Months)|Lecture 03: Function to Calculate Age (Result in Years & Months)]]
- [ ] [[Lecture S19-04 - Functions that Split Text and Numbers from Cell Value|Lecture 04: Functions that Split Text and Numbers from Cell Value]]
- [ ] [[Lecture S19-05 - Using Functions in Sub Procedures|Lecture 05: Using Functions in Sub Procedures]]
- [ ] [[Lecture S19-06 - Activity Function That Sums based on Cell Color|Lecture 06: Activity: Function That Sums based on Cell Color]]
- [ ] [[Lecture S19-07 - Project Function That Sums based on Cell Color|Lecture 07: Project: Function That Sums based on Cell Color]]
- [ ] [[Lecture S19-08 - Key Takeaways Function Procedures in VBA & Course Map Status|Lecture 08: Key Takeaways Function Procedures in VBA & Course Map Status]]

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
