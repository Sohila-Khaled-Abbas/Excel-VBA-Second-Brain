---
type: project-doc
project: Project 03 - Invoice Generation Tool
doc_type: requirements
status: in-progress
created: 2026-10-03
updated: 2026-10-03
---

# Project 03: Requirements & Problem Statement

## 🏢 Business Problem
Operational billing in small-to-mid-sized enterprises is often plagued by manual copy-pasting between customer contact sheets, product catalogs, and invoice templates. Billing errors, mismatched tax calculations, forgotten attachments, and inconsistent email formatting lead to delayed receivables and unprofessional client communication.

## 🎯 Solution Architecture Requirements
Build an interactive, UserForm-driven billing application that:
1. Provides an intuitive graphical interface for selecting customer master data and line items.
2. Implements a real-time **searchable multi-column ListBox** that filters customers dynamically as the user types.
3. Allows editing, adding, and validating master data directly from the UserForm.
4. Uses a MultiPage control to step cleanly through Customer Selection → Product Selection → Invoice Generation.
5. Populates a professional invoice template worksheet with line items, formulas, and auto-calculated totals.
6. Exports the completed invoice to a locked **PDF file** with structured naming (`Invoice_INV1001_AcmeCorp.pdf`).
7. Creates a personalized **Outlook email** pre-populating recipient address, professional body copy, and the attached PDF invoice.
