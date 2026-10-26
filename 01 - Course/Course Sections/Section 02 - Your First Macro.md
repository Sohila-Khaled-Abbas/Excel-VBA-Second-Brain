---
type: course-section
section_number: 2
section_title: Your First Macro
lecture_count: 10
status: completed
tags:
  - course-section
  - vba
  - section-02
---

# Section 02: Your First Macro

> [!important]
> **Curriculum Pillar**: Section 2 contains **10 lectures**.  
> **Status**: `COMPLETED`  
> **Course Practice Files**: `SectionFiles/S2_RecordMacros.xlsm`, `S2_RecordMacros_Start.xlsm`, `S2_RecordMacros_Start.xlsx`

---

## 🎯 Section Objectives & Automation Capabilities
- Master the Excel Macro Recorder: know precisely when to use it, what it records, and its architectural limitations.
- Navigate the Visual Basic Editor (VBE), Project Explorer, Properties Window, and Code Windows.
- Understand Absolute vs Relative macro recording modes (`ActiveCell.Offset`).
- Implement Go To Special (`SpecialCells`) for dynamic selection patterns that adapt to variable data shapes.
- Build modern, interactive execution triggers (Buttons, Shapes, Quick Access Toolbar, Ribbon, Keyboard Shortcuts).

---

## 🧠 Architectural Takeaways & Mental Models
- **Recorder as a Translator**: The Macro Recorder is an exploratory reconnaissance tool. It captures user actions verbatim but generates bloated, inefficient code riddled with `.Select` and `.Activate`.
- **Absolute vs. Relative Recording**: Absolute recording hardcodes exact cell references (e.g., `Range("B4").Select`). Relative recording (`ActiveCell.Offset(r, c)`) enables scalable row-by-row consolidation.
- **Go To Special (`xlCellTypeConstants`, `xlCellTypeFormulas`)**: Extremely powerful for non-contiguous data manipulation, isolating blank cells, or deleting notes in one shot without looping.
- **Clean Naming Conventions**: Never begin macro names with numbers or include spaces or special characters. Use camelCase or Pascal_Case.

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- Never leave recorded `.Select` / `.Activate` statements in production code; they degrade execution speed by up to 90% and trigger annoying screen flicker.
- In Office 365, "Comments" are now threaded conversations, while traditional yellow boxes are "Notes". VBA syntax uses `Range.Comment` / `ClearComments` for notes.

---

## 📚 Lecture Index

- [x] [[Lecture S02-01 - Overview Macros & Visual Basic Editor (VBE)|Lecture 01: Overview: Macros & Visual Basic Editor (VBE)]] — *Kicking Start*
- [x] [[Lecture S02-02 - Steps to Your First Macro Recording|Lecture 02: Steps to Your First Macro Recording]] — *Easy -- Don't use Number in beginning -- Don't use special Characters -- Don't use Space -- Do Test.*
- [x] [[Lecture S02-03 - Visual Basic Editor What You Need to Know|Lecture 03: Visual Basic Editor: What You Need to Know]] — *Alt + F11, Developer Tab / Visual Basic, View / Macros / View Macros, then edit, Right Mouse click on sheet / View Code. "Sub". View/Toolbars/Edit*
- [x] [[Lecture S02-04 - Project Dynamic Selection (Go to Special) - Perfect for Flexible Macros|Lecture 04: Project: Dynamic Selection (Go to Special) - Perfect for Flexible Macros]] — *Logical Thinking*
- [x] [[Lecture S02-05 - Project Copy (Consolidate) Data - Absolute versus Relative Macro Recording|Lecture 05: Project: Copy (Consolidate) Data - Absolute versus Relative Macro Recording]] — *Done -- Took Notes Before*
- [x] [[Lecture S02-06 - 7 Ways to Run Macros VBA code (incl. creative & modern buttons)|Lecture 06: 7 Ways to Run Macros / VBA code (incl. creative & modern buttons)]] — *Done -- Took Notes Before*
- [x] [[Lecture S02-07 - OFFICE 365 Update Change from Comments to Notes|Lecture 07: OFFICE 365 Update: Change from Comments to Notes]] — *Done -- Took Notes Before*
- [x] [[Lecture S02-08 - Activity Record a Macro that Deletes all Comments|Lecture 08: Activity: Record a Macro that Deletes all Comments]] — *Done -- Took Notes Before*
- [x] [[Lecture S02-09 - What's the Difference Between Macros and VBA|Lecture 09: What's the Difference Between Macros and VBA?]] — *Done -- Took Notes Before*
- [x] [[Lecture S02-10 - Quiz Test Your Knowledge on Macros|Lecture 10: Quiz: Test Your Knowledge on Macros]] — *Full Mark*

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
