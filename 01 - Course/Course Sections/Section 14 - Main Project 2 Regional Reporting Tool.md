---
type: course-section
section_number: 14
section_title: Main Project #2: Regional Reporting Tool
lecture_count: 11
status: not-started
tags:
  - course-section
  - vba
  - section-14
---

# Section 14: Main Project #2: Regional Reporting Tool

> [!important]
> **Curriculum Pillar**: Section 14 contains **11 lectures**.  
> **Status**: `NOT-STARTED`  
> **Course Practice Files**: `Milestones/SalesReporting/Monthly_Sales_Reporting_Template_Tool.xlsm`, 4 regional sales data workbooks

---

## 🎯 Section Objectives & Automation Capabilities
- Build **Main Project #2: Regional Sales Reporting Tool** (Complete Multi-File Consolidation & Reporting Pipeline).
- Implement multi-file workbook batch ingestion using `FileDialog` with `AllowMultiSelect:=True`.
- Dynamically transfer sales data into a centralized raw repository while updating status bars.
- Generate formatted, presentation-ready regional and managerial reports from a master template.
- Export clean semicolon-delimited CSV text files for enterprise ERP compatibility.
- Implement administrative security: password-protected tab hiding (`xlSheetVeryHidden`) and automated audit tracking.

---

## 🧠 Architectural Takeaways & Mental Models
- **Milestone Project #2**: Represents a complete end-to-end Data Engineering & BI pipeline built natively in Excel VBA.
- **Modular Pipeline Architecture**:
  1. `Import_Sales`: Batch file picker, clipboard bypass, sequential append to `RawData`.
  2. `Create_Reports`: Refresh pivot caches, apply template formatting, generate static value exports.
  3. `Create_CSV`: Custom `FreeFile` streaming with guaranteed semicolon delimiters.
  4. `Admin_Tasks`: Password prompt protecting sensitive mapping and configuration tabs.
- **`xlSheetVeryHidden`**: Makes tabs completely invisible in the standard Excel UI (unhide option is disabled); only accessible via VBA code with the correct administrative password.

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- Opening multiple workbooks without setting `Application.ScreenUpdating = False` causes severe screen flashing and multiplies execution time by 5-10x.
- Always check if the user clicked "Cancel" on the `FileDialog` before attempting to loop through `SelectedItems`.

---

## 📚 Lecture Index

- [ ] [[Lecture S14-01 - What You'll Learn (& Should be Aware of..)|Lecture 01: What You'll Learn (& Should be Aware of..)]]
- [ ] [[Lecture S14-02 - Project Overview Regional Sales Reporting Tool|Lecture 02: Project Overview: Regional Sales Reporting Tool]]
- [ ] [[Lecture S14-03 - Importing Data from Multiple Files using MultiSelect|Lecture 03: Importing Data from Multiple Files using MultiSelect]]
- [ ] [[Lecture S14-04 - Setting up a Template as the Basis for Different Reports|Lecture 04: Setting up a Template as the Basis for Different Reports]]
- [ ] [[Lecture S14-05 - Exporting Reports as Excel files (Pivot Table & Hard Coded Reports)|Lecture 05: Exporting Reports as Excel files (Pivot Table & Hard Coded Reports)]]
- [ ] [[Lecture S14-06 - Fine-tuning Report Export Procedure (Double-check if user wants to overwrite)|Lecture 06: Fine-tuning Report Export Procedure (Double-check if user wants to overwrite)]]
- [ ] [[Lecture S14-07 - Setting up Automatic Export of Text File (Fix delimiter in CSV to Semicolon)|Lecture 07: Setting up Automatic Export of Text File (Fix delimiter in CSV to Semicolon)]]
- [ ] [[Lecture S14-08 - Track the List of Tasks Completed with VBA|Lecture 08: Track the List of Tasks Completed with VBA]]
- [ ] [[Lecture S14-09 - Final Touches & Creating Interface to Hide Specific Tabs with a Password|Lecture 09: Final Touches & Creating Interface to Hide Specific Tabs with a Password]]
- [ ] [[Lecture S14-10 - Key Takeaways & Tips for VBA based Excel Tools|Lecture 10: Key Takeaways & Tips for VBA based Excel Tools]]
- [ ] [[Lecture S14-11 - Second Milestone Completed!|Lecture 11: Second Milestone Completed!]]

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
