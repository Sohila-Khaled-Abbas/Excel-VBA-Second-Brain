---
type: project-doc
project: Project 01 - Table of Contents Generator
doc_type: lessons-learned
status: completed
created: 2026-10-03
updated: 2026-10-03
---

# Project 01: Key Lessons Learned

1. **SubAddress Quoting Trap**: If a worksheet name contains spaces or hyphens, Excel requires the target address to be enclosed in single quotes (`'Sheet Name'!A1`). Forgetting this breaks hyperlinks.
2. **`PERSONAL.XLSB` Deployment**: Placing this code in standard modules of `PERSONAL.XLSB` and mapping it to a custom Quick Access Toolbar (QAT) button turns a one-off macro into an invaluable daily productivity tool.
