---
type: resources
topic: Course Materials & Workbooks Guide
tags:
  - resources
  - workbooks
  - course-files
  - reference
created: 2026-10-26
updated: 2026-10-26
---

# 📦 Course Materials & Workbooks Guide

> [!important]
> **Source Directory**: `D:\courses\Data Science\Data Engineering\Udemy Courses\VBA & Marcos Course`  
> **Course**: *Master Excel VBA and Macros from Basics to Advanced with Three Complete Real-World Automation Projects*  
> **Instructor**: Leila Gharani (Microsoft MVP)

This document provides a complete inventory and mapping of all official practice workbooks, slide decks, milestone project tools, and starter templates accompanying the course.

---

## 🗂️ Master Directory Overview

```text
D:\courses\Data Science\Data Engineering\Udemy Courses\VBA & Marcos Course\
├── ALL_VBA_Section_Milestone_Files_In_One_Place/
│   ├── Milestones/                          # 3 Production-Style Enterprise Automation Projects
│   │   ├── InvoiceGeneration/               # Project 03: Complete UserForm-driven invoicing system
│   │   ├── SalesReporting/                  # Project 02: Multi-file consolidation & ETL pipeline
│   │   └── TableOfContents/                 # Project 01: Dynamic TOC navigation generator
│   └── SectionFiles/                        # 15+ Practice Workbooks (.xlsm) with starter & solution code
├── Detailed_VBA_Course_Outline.xlsx         # Authoritative course syllabus (21 sections, 174 lectures, notes)
├── Excel+VBA_Course+Slides.pdf              # Comprehensive 130-page lecture slide deck & syntax reference
├── How_to_Ask_Questions.pdf                 # Student support & debugging guidelines
├── IconSets.xlsx                            # Visual UI icons for UserForms and modern dashboard buttons
└── Excel-VBA-Second-Brain/                  # This interconnected Obsidian Second Brain knowledge vault
```

---

## 🏆 Milestone Project Workbooks

| Project # | Project Title | Path & Files | Primary Techniques & Architecture | Second Brain Hub |
| :---: | :--- | :--- | :--- | :--- |
| **01** | **Table of Contents Tool** *(Section 09)* | `Milestones/TableOfContents/`<br>• `Project_Automatic_TableOfContent.xlsm` | • `Sheets.Add` positioning<br>• Dynamic sheet iteration & two-way hyperlinks<br>• `On Error GoTo` safe sheet reset<br>• Personal Macro Workbook (`PERSONAL.XLSB`) | [[Project 01 - Table of Contents Generator]] |
| **02** | **Regional Sales Reporting Tool** *(Section 14)* | `Milestones/SalesReporting/`<br>• `Monthly_Sales_Reporting_Template_Tool.xlsm`<br>• `Sales_Document_BereKleid.xlsx`<br>• `Sales_Document_LucasBasics.xlsx`<br>• `Sales_Document_MetaCreations.xlsx`<br>• `Sales_Document_UrbanRight.xlsx` | • `Application.FileDialog(msoFileDialogFilePicker)` with `AllowMultiSelect`<br>• Batch file opening, data copy, and close<br>• Automated PivotTable creation & cache refreshing<br>• Custom semicolon-delimited CSV export via `FreeFile`<br>• Password-protected sheet visibility toggle | [[Project 02 - Regional Sales Reporting Tool]] |
| **03** | **Invoice Generation Tool** *(Section 18)* | `Milestones/InvoiceGeneration/`<br>• `Invoice_Generator.xlsm`<br>• `Invoice_Generator_Start.xlsm` | • UserForms (`FmInputMaster`, `FmViewMaster`, `FmGetCustomer`, `FmGetInvoice`)<br>• Multi-column `ListBox` with real-time `InStr` text search<br>• Excel Table (`ListObjects`) CRUD operations<br>• `ExportAsFixedFormat xlTypePDF`<br>• Outlook Late-Binding automation with email attachment | [[Project 03 - Invoice Generation Tool]] |

---

## 💻 Section Practice Workbooks (`SectionFiles/`)

Every section file includes starter templates (`_Start.xlsx` / `_Start.xlsm`) for deliberate practice alongside completed solution workbooks (`.xlsm`):

| Section | Solution File (`.xlsm`) | Starter File(s) | Key Modules & Topics Covered |
| :---: | :--- | :--- | :--- |
| **02** | `S2_RecordMacros.xlsm` | `S2_RecordMacros_Start.xlsm`<br>`S2_RecordMacros_Start.xlsx` | `A_FormatSheet`, `B_SpecialCells`, `C_AbsoluteRelative`, `ProjectActivity` (Macro recorder, VBE, relative references, Go To Special). |
| **04** | `S4_ReferencingRanges.xlsm` | `S4_ReferencingRanges_Start.xlsm`<br>`S4_ReferencingRanges_Start.xlsx`<br>`Destbook.xlsx` | `LessonsRanges`, `ProjectActivity` (Range vs Cells, Offset, 4 Last Row methods, CurrentRegion, Worksheet codenames, hardcoded workbook copy). |
| **05** | `S5_Variables.xlsm` | `S5_Variables_Start.xlsx` | `LessonExample`, `LessonMain`, `ProjectActivity` (Dim, Option Explicit, Integer vs Long, Double, String, Date, Object variables, Set statement, Scope). |
| **06** | `S6_Collections_With_IF.xlsm` | `S6_Collections_With_IF_Start.xlsx` | `LessonsCollections`, `ProjectActivity` (With...End With, For Each in Sheets/Ranges, If-Then-Else, Select Case, GoTo, Unhide All Sheets). |
| **07** | `S7_BuiltinFunctions.xlsm` | `S7_BuiltinFunctions_Start.xlsx` | `LessonsFunctions`, `ProjectActivity` (VBA vs Worksheet functions, InStr, DateAdd, MsgBox with vbYesNo, InputBox vs Application.InputBox Type 8). |
| **08** | `S8_ErrorHandling.xlsm` | `S8_ErrorHandling_Start.xlsm` | `ErrorLessons`, `ProcedureLessons`, `ProjectActivity` (Watch/Locals/Immediate, On Error GoTo, Resume Next, Err.Clear, ScreenUpdating, ByRef vs ByVal). |
| **10** | `S10_Looping.xlsm` | `S10_Looping_Start.xlsx` | `Lesson_a_For_Next`, `Lesson_b_DoLoop`, `Lesson_c_Find`, `ProjectActivity` (For Next, Reverse Step -1 loops for row deletion, Do While/Until, Range.Find, Timer). |
| **11** | `S11_Arrays.xlsm` | `S11_Arrays_Start.xlsx` | `LessonsArrays`, `ProjectActivity` (1D and 2D arrays, ReDim Preserve, Variant memory arrays for instant bulk read/write operations). |
| **12** | `S12_Folders_Files.xlsm` | `S12_Folders_Files_Start.xlsx`<br>`ProjectActivity.csv`<br>`TestTextCSV.csv` | `LessonsFilesFolders`, `ProjectActivity` (Dir() function, GetOpenFilename, FileDialog folder loops, FreeFile text export, semicolon CSV). |
| **13** | `S13_Formulas_Tables_Pivot.xlsm` | `S13_Formulas_Tables_Pivot_Start.xlsx` | `LessonFormulaTables`, `LessonPivotTables`, `ProjectActivity` (FormulaR1C1 injection, ListObjects syntax, PivotTable caches, IFERROR wrapping). |
| **15** | `S15_Other_Applications.xlsm` | `S15_Other_Applications_Start.xlsx`<br>`Letter_Word.docx`<br>`Regional Report PowerPoint.pptx` | `LessonsApplications` (PDF export, Outlook email dispatch, Word bookmarks automation, PowerPoint presentation generation). |
| **16** | `S16_Events.xlsm` | `S16_Events_Start.xlsx` | `LessonsEvents`, `ThisWorkbook`, `Sheet4-6` (Workbook Open/BeforeClose, Worksheet SelectionChange, Worksheet_Change, Application.Intersect, EnableEvents). |
| **17** | `S17_UserForms_ActiveX.xlsm` | `S17_UserForms_ActiveX_Start.xlsx`<br>`S17_PictureClip.wmf`<br>`IconSets.xlsx` | `LessonsForms`, `FrmWSInput`, `FrmWBCheck` (ActiveX Controls, UserForm initialization, ComboBox, Multi-column ListBox, Form controls vs ActiveX). |
| **19** | `S19_FunctionProcedures.xlsm` | `S19_FunctionProcedures_Start.xlsx` | `FunctionSubLessons`, `ProjectActivity` (Custom UDFs, Optional arguments, Age calculation, Split text and numbers, SumByColor function). |
| **20** | `S20_Charts.xlsm` | `S20_Charts_Start.xlsx` | `ChartLessons`, `FrmChart` (ChartObjects hierarchy, programmatic chart styling, chart animation with DoEvents/Timer, embedding charts in UserForms). |

---

## 📑 Core Reference Documents

- **`Detailed_VBA_Course_Outline.xlsx`**:
  Tracks all 21 sections, 174 lectures, user completion checkboxes, and student study notes. Synchronized directly with [[Course Roadmap]] and [[Progress Dashboard]].
- **`Excel+VBA_Course+Slides.pdf`**:
  A 130-slide visual curriculum containing essential syntax cheat sheets, object diagrams, memory comparison tables, and key takeaways for every single module.
- **`How_to_Ask_Questions.pdf`**:
  Instructor framework for effective problem description, isolating bugs with minimal code examples, and writing clean debugging queries. Synthesized in [[VBA Q&A and Debugging Strategy]].
- **`IconSets.xlsx`**:
  A collection of clean vector icons and UI graphics for building professional Excel dashboards and modern UserForm navigation bars.

---

## 🔒 Confidentiality & Git Policy

In accordance with the repository security guidelines set out in `_private/README.md`:
- All binary `.xlsm`, `.xlsx`, `.pptx`, `.docx`, and `.pdf` files remain stored on local disk and are **strictly excluded** from public git commits via `.gitignore`.
- This vault stores the **synthesized knowledge, structural diagrams, evergreen concepts, verified code snippets, and architectural workflows** derived from the course material.
