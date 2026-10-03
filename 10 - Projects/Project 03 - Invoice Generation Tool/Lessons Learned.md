---
type: project-doc
project: Project 03 - Invoice Generation Tool
doc_type: lessons-learned
status: in-progress
created: 2026-10-03
updated: 2026-10-03
---

# Project 03: Architecture & UI Takeaways

1. **Separation of Concerns**: Keep UserForm code focused purely on UI event handling, input validation, and reading/writing control states. Delegate business calculations and PDF/Outlook operations to dedicated standard modules.
2. **Defensive Late Binding**: Always use late binding (`CreateObject`) for external Office applications in tools distributed to multiple team members. Early binding creates broken reference errors if users have different versions of Office (e.g., Office 2016 vs. M365).
