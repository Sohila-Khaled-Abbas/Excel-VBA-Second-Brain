---
type: course-section
section_number: 18
section_title: Main Project #3: Invoice Generation Tool
lecture_count: 11
status: not-started
tags:
  - course-section
  - vba
  - section-18
---

# Section 18: Main Project #3: Invoice Generation Tool

> [!important]
> **Curriculum Pillar**: Section 18 contains **11 lectures**.  
> **Status**: `NOT-STARTED`  
> **Course Practice Files**: `Milestones/InvoiceGeneration/Invoice_Generator.xlsm`, `Invoice_Generator_Start.xlsm`

---

## 🎯 Section Objectives & Automation Capabilities
- Build **Main Project #3: Invoice Generation Tool** (Enterprise UserForm Automation System).
- Design and code a full UserForm suite: `FmInputMaster`, `FmViewMaster`, `FmGetCustomer`, `FmGetInvoice`.
- Build a real-time, searchable multi-column ListBox that filters customer records as the user types.
- Perform master data CRUD operations (Create, Read, Update, Delete) on structured Excel Tables.
- Automate single and batch PDF invoice generation, file naming, directory creation, and Outlook email dispatch with attachments.

---

## 🧠 Architectural Takeaways & Mental Models
- **Milestone Project #3 Architecture**:
  - Form Suite:
    1. `FmInputMaster`: Create & edit customer/product records with dynamic `.Tag = "required"` validation.
    2. `FmViewMaster`: Browse master datasets, view multi-column records, and pass selected IDs to editing forms.
    3. `FmGetCustomer`: Real-time text search box leveraging `InStr` to filter ListBox items instantaneously.
    4. `FmGetInvoice`: MultiPage form managing invoice line items, calculating totals, exporting PDF invoices, and composing Outlook emails.
- **Dynamic InStr Filtering**:
  As the user types in `tbCustomer`, the `Change` event clears the ListBox and repopulates only rows where `InStr(1, cellValue, searchStr, vbTextCompare) > 0`.

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- In UserForms with multiple inter-communicating dialogs, manage state using public variables or pass IDs via hidden labels/properties to avoid out-of-sync edits.
- Ensure VAT and price inputs are strictly validated with `IsNumeric()` before writing to worksheet tables.

---

## 📚 Lecture Index

- [ ] [[Lecture S18-01 - Tool Overview What You Will Learn|Lecture 01: Tool Overview: What You Will Learn]]
- [ ] [[Lecture S18-02 - Setting up an Input UserForm to Collect Master Data|Lecture 02: Setting up an Input UserForm to Collect Master Data]]
- [ ] [[Lecture S18-03 - Working with Multi-Column ListBoxes|Lecture 03: Working with Multi-Column ListBoxes]]
- [ ] [[Lecture S18-04 - Edit Master Data with UserForm|Lecture 04: Edit Master Data with UserForm]]
- [ ] [[Lecture S18-05 - Searchable ListBox (with Selection Change event)|Lecture 05: Searchable ListBox (with Selection Change event)]]
- [ ] [[Lecture S18-06 - Enhanced Searchable ListBox Write Selection back to Excel Table|Lecture 06: Enhanced Searchable ListBox: Write Selection back to Excel Table]]
- [ ] [[Lecture S18-07 - MultiPage UserForm to Create Invoices & Email (Setup)|Lecture 07: MultiPage UserForm to Create Invoices & Email (Setup)]]
- [ ] [[Lecture S18-08 - Export PDF & Excel Workbook for Each Invoice based on ListBox Selection|Lecture 08: Export PDF & Excel Workbook for Each Invoice based on ListBox Selection]]
- [ ] [[Lecture S18-09 - Create Email in Outlook & Attach PDF Invoice|Lecture 09: Create Email in Outlook & Attach PDF Invoice]]
- [ ] [[Lecture S18-10 - Final Updates & Debugging|Lecture 10: Final Updates & Debugging]]
- [ ] [[Lecture S18-11 - Last Milestone Completed!|Lecture 11: Last Milestone Completed!]]

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
