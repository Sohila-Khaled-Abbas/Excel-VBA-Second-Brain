---
type: troubleshooting-matrix
domain: VBA
topic: Diagnostics
tags:
  - troubleshooting
  - debugging
  - error-codes
created: 2026-10-03
updated: 2026-10-03
---

# 🛠️ VBA Troubleshooting Matrix

| Error Code | Error Message | Common Cause | How to Diagnose | Exact Fix | Prevention Best Practice |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Error 9** | *Subscript out of range* | Referencing a worksheet, workbook, or array index that does not exist | Check spelling of sheet tab name or array bounds in Locals Window | Verify sheet name or use dynamic `UBound(arr)` | Use `SheetExists()` helper function before referencing |
| **Error 13** | *Type mismatch* | Assigning text to a numeric variable, or mathematical operations on `#N/A` | Inspect variable types in Locals Window using `F8` | Ensure numeric conversion with `CDbl()`, `CLng()` or check `IsNumeric()` | Validate cell content before arithmetic |
| **Error 91** | *Object variable or With block variable not set* | Missing `Set` keyword on object variable, or object lookup failed (`Find` returned `Nothing`) | Check line where object is assigned; look for missing `Set` | Add `Set ws = ...`; wrap with `If Not obj Is Nothing Then` | Always verify object references before calling their properties |
| **Error 6** | *Overflow* | Number exceeds data type bounds (e.g., using `Integer` past 32,767) | Check row counter variable value in Watch Window | Change `Dim r As Integer` to `Dim r As Long` | **Never use `Integer` in Excel VBA; always use `Long`** |
| **Error 1004** | *Application-defined or object-defined error* | Invalid range address, writing to protected sheet, deleting non-existent file | Print target range address in Immediate Window (`?rng.Address`) | Unprotect sheet, check valid range boundaries | Check `.ProtectContents` property before writing |
| **Error 5** | *Invalid procedure call or argument* | String function parameter out of bounds (e.g., `Mid(str, -1)`) | Check argument values in Immediate Window | Validate index parameters with `Len()` | Guard string slice arguments |
