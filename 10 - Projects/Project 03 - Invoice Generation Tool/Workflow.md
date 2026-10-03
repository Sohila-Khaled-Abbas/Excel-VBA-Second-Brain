---
type: project-doc
project: Project 03 - Invoice Generation Tool
doc_type: workflow
status: in-progress
created: 2026-10-03
updated: 2026-10-03
---

# Project 03: Implementation Workflow

1. **UserForm Setup & Controls**:
   - `MultiPage` control dividing workflow into logical phases.
   - `TextBoxSearch`: Triggers `TextBoxSearch_Change` event.
   - `ListBoxCustomers`: Multi-column (ID, Name, Contact, Email, Terms).
   - `ListBoxProducts`: Multi-column (SKU, Description, Unit Price, Stock).
2. **Searchable ListBox Logic**:
   - Filter memory array based on `InStr(1, customerName, txtSearch, vbTextCompare) > 0`.
   - Repopulate ListBox dynamically without flickering.
3. **Master Data CRUD**:
   - Provide "Add New Customer" modal writing directly back to the `CustomerMaster` Excel Table (`ListObject`).
4. **Invoice Template Population**:
   - Clear existing line items on `Invoice_Template` tab.
   - Write header metadata (Customer details, Invoice Number, Date, Due Date).
   - Loop through selected line items, injecting SKU, Description, Qty, Unit Price.
5. **Document Rendering & Email Dispatch**:
   - Call `.ExportAsFixedFormat` targeting `ThisWorkbook.Path & "\Invoices"`.
   - Call late-bound Outlook procedure to create draft email and attach the PDF.
