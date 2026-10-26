---
type: course-section
section_number: 12
section_title: Working with Files, Folders & Text Files
lecture_count: 9
status: not-started
tags:
  - course-section
  - vba
  - section-12
---

# Section 12: Working with Files, Folders & Text Files

> [!important]
> **Curriculum Pillar**: Section 12 contains **9 lectures**.  
> **Status**: `NOT-STARTED`  
> **Course Practice Files**: `SectionFiles/S12_Folders_Files.xlsm`, `S12_Folders_Files_Start.xlsx`, `ProjectActivity.csv`, `TestTextCSV.csv`

---

## 🎯 Section Objectives & Automation Capabilities
- Check file and directory existence programmatically using the `Dir()` function.
- Provide interactive file selection using `Application.GetOpenFilename` (filters, titles, multi-select).
- Enable folder browsing and recursive directory loops using `Application.FileDialog(msoFileDialogFolderPicker)`.
- Export Excel data directly to clean CSV files without saving or converting whole workbooks.
- Master low-level text file streaming using `FreeFile`, `Open...For Output / Input`, and `Print #` / `Line Input #`.

---

## 🧠 Architectural Takeaways & Mental Models
- **`Dir()` Function Mechanics**: `Dir(path)` returns the filename if it exists, or `""` if not. Calling `Dir()` with no arguments continues iterating through matching files in the same directory.
- **Interactive File Selection**:
  - `GetOpenFilename`: Prompts user for file(s) and returns the full path as a String (or False if canceled). Does not open the file automatically.
  - `FileDialog(msoFileDialogFolderPicker)`: Clean UI for selecting a folder target.
- **The FreeFile Semicolon CSV Export**: Standard Excel `SaveAs CSV` uses the Windows regional list separator (comma in US, semicolon in Europe). Low-level text streaming via `Print #` guarantees exact delimiters regardless of user machine settings:
  ```vba
  fNum = FreeFile
  Open filePath For Output As #fNum
  Print #fNum, lineString
  Close #fNum
  ```

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- Calling `Dir()` a second time while an existing `Dir()` loop is running resets the directory pointer and crashes the iteration.
- Forgetting to `Close #fNum` locks the text file on the operating system, preventing subsequent read/write operations until Excel restarts.

---

## 📚 Lecture Index

- [ ] [[Lecture S12-01 - Overview Working with Folders & Files|Lecture 01: Overview: Working with Folders & Files]]
- [ ] [[Lecture S12-02 - DIR Function Check if Files or Folders Exist|Lecture 02: DIR Function: Check if Files or Folders Exist]]
- [ ] [[Lecture S12-03 - GetOpenFileName Method User Selects a File|Lecture 03: GetOpenFileName Method: User Selects a File]]
- [ ] [[Lecture S12-04 - FileDialog Property Loop Inside a Folder|Lecture 04: FileDialog Property: Loop Inside a Folder]]
- [ ] [[Lecture S12-05 - Create CSV File from Data in Excel|Lecture 05: Create CSV File from Data in Excel]]
- [ ] [[Lecture S12-06 - Printing & Writing to Text Files for More Control|Lecture 06: Printing & Writing to Text Files for More Control]]
- [ ] [[Lecture S12-07 - Activity Export Data to Text file (fix Delimiter to Semicolon)|Lecture 07: Activity: Export Data to Text file (fix Delimiter to Semicolon)]]
- [ ] [[Lecture S12-08 - Project Export Data to Text file (fix Delimiter to Semicolon)|Lecture 08: Project: Export Data to Text file (fix Delimiter to Semicolon)]]
- [ ] [[Lecture S12-09 - Key Takeaways Files & Folders|Lecture 09: Key Takeaways Files & Folders]]

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
