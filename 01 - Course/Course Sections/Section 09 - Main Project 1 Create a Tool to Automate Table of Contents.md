---
type: course-section
section_number: 9
section_title: Main Project #1: Create a Tool to Automate Table of Contents
lecture_count: 7
status: not-started
tags:
  - course-section
  - vba
  - section-09
---

# Section 09: Main Project #1: Create a Tool to Automate Table of Contents

> [!important]
> **Curriculum Pillar**: Section 9 contains **7 lectures**.  
> **Status**: `NOT-STARTED`  
> **Course Practice Files**: `Milestones/TableOfContents/Project_Automatic_TableOfContent.xlsm`

---

## 🎯 Section Objectives & Automation Capabilities
- Build **Main Project #1**: An automated, self-healing Table of Contents (TOC) generator with dynamic two-way hyperlinks.
- Solicit destination coordinates interactively using `Application.InputBox(Type:=8)`.
- Implement overwriting warnings with default cancel safety (`vbDefaultButton2`).
- Dynamically extract worksheet titles from cell `A1` and construct valid hyperlink SubAddresses.
- Deploy the tool globally to all workbooks via the **Personal Macro Workbook** (`PERSONAL.XLSB`).

---

## 🧠 Architectural Takeaways & Mental Models
- **Milestone Project #1**: Bridges object referencing, collection looping, user dialogs, and error handling into a production utility.
- **Handling Sheet Names with Spaces**: Hyperlinks to sheet names containing spaces or punctuation must be enclosed in single quotes in the SubAddress:
  `SubAddress:="'" & ShName & "'!A1"`.
- **Filtering Hidden Tabs**: Check `If Sh.Visible = xlSheetVisible Then` to prevent creating dead hyperlinks to hidden or very hidden calculation sheets.
- **Personal Macro Workbook (`PERSONAL.XLSB`)**: Lives in the `XLSTART` directory. Opens hidden in every Excel session, making your macros available in every workbook on your machine. Assign to a custom Ribbon icon for one-click access.

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- Failing to quote worksheet names with spaces in hyperlink sub-addresses results in "Reference is not valid" errors when users click the link.
- Not verifying if a Table of Contents already exists can lead to recursive self-referencing links.

---

## 📚 Lecture Index

- [ ] [[Lecture S09-01 - Project overview Table of Contents (TOC) with one click|Lecture 01: Project overview: Table of Contents (TOC) with one click]]
- [ ] [[Lecture S09-02 - TOC Project The Basics|Lecture 02: TOC Project: The Basics]]
- [ ] [[Lecture S09-03 - TOC Project Adding Hyperlinks|Lecture 03: TOC Project: Adding Hyperlinks]]
- [ ] [[Lecture S09-04 - TOC Project Testing the VBA Code|Lecture 04: TOC Project: Testing the VBA Code]]
- [ ] [[Lecture S09-05 - TOC Project Adding Error Handling & More Checks|Lecture 05: TOC Project: Adding Error Handling & More Checks]]
- [ ] [[Lecture S09-06 - TOC Project Adding to Personal Macro Workbook (Make it Available to all Files)|Lecture 06: TOC Project: Adding to Personal Macro Workbook (Make it Available to all Files)]]
- [ ] [[Lecture S09-07 - First Milestone Completed!|Lecture 07: First Milestone Completed!]]

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
