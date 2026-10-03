---
type: project-doc
project: Project 03 - Invoice Generation Tool
doc_type: testing
status: in-progress
created: 2026-10-03
updated: 2026-10-03
---

# Project 03: Testing Matrix & Production Considerations

| Scenario | Test Condition | Expected Result | Pass/Fail |
| :--- | :--- | :--- | :---: |
| **Invalid Email Address** | User enters text without `@` or domain | Validation flags error; highlights text box in red | ✅ Pass |
| **No Line Items Selected** | User clicks generate invoice with 0 items | Prompt: "Please add at least one line item" | ✅ Pass |
| **Invoices Directory Missing** | `\Invoices\` folder does not exist | Code creates folder using `MkDir` before PDF export | ✅ Pass |
| **Outlook Not Running** | Outlook application closed | `CreateObject("Outlook.Application")` spawns background instance | ✅ Pass |
