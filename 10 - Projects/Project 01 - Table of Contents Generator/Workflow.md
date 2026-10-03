---
type: project-doc
project: Project 01 - Table of Contents Generator
doc_type: workflow
status: completed
created: 2026-10-03
updated: 2026-10-03
---

# Project 01: Implementation Workflow

1. **Pre-flight Checks**:
   - Verify workbook is not protected.
   - Cache `Application.ScreenUpdating` state.
2. **Sheet Creation / Reset**:
   - Check if sheet `"Table of Contents"` or `"TOC"` already exists using an `On Error Resume Next` lookup function.
   - If exists, clear contents or delete and re-add.
3. **Iterative Hyperlink Generation**:
   - Use `For Each ws In ActiveWorkbook.Worksheets`.
   - Build subaddress hyperlink using `ws.Name & "!A1"`.
   - Handle worksheet names containing spaces by wrapping in single quotes (`"'" & ws.Name & "'!A1"`).
4. **Return Navigation Insertion**:
   - Insert `=HYPERLINK("#'Table of Contents'!A1", "« Back to TOC")` in cell `A1` of each target worksheet.
5. **Final Aesthetics**:
   - Apply corporate palette (Microsoft Green header, light borders, gridlines visible).
