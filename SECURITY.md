# Security Policy

## Supported Versions

This repository contains educational documentation, VBA code patterns, and automation scripts. The following versions receive active maintenance:

| Version / Release | Supported |
| :--- | :--- |
| Latest (`main` branch) | ✅ Active |
| Tagged releases (v2.x) | ✅ Active |
| Tagged releases (v1.x) | ⚠️ Security fixes only |

---

## ⚠️ Important: VBA Macro Security Notice

Excel VBA macros can pose security risks when sourced from unknown or untrusted workbooks. Before using any VBA code from this repository in a live Excel environment:

### For All Users
1. **Review code before running**: Never execute VBA code you haven't read and understood. All macros in this vault's documentation should be inspected in the Visual Basic Editor (Alt+F11) before running.
2. **Use the `_private/` directory boundary**: The local `.gitignore`-protected `_private/` folder is the only place practice workbooks (`.xlsm`) should reside. Never open untrusted `.xlsm` files downloaded from the public internet with macros enabled.
3. **Excel Trust Center Settings**: Set Excel's macro security to **"Disable all macros with notification"** (`Tools → Macro → Security → Medium`). This prompts you to review each workbook before enabling its macros.

### For Contributors Adding Code Snippets
4. **No file system destructive operations**: Do not contribute VBA patterns that perform `Kill`, `RmDir`, or recursive deletion on hard-coded system paths without explicit user confirmation dialogs.
5. **No network calls without disclosure**: Any snippet using `XMLHTTP`, `WinHttp`, or `Microsoft.XMLHTTP` to make external network requests must clearly document its purpose, endpoint, and data transmitted in the note's header.
6. **No registry modifications**: Do not contribute snippets that write to the Windows Registry using `CreateObject("WScript.Shell").RegWrite`.
7. **No shell execution without warning**: Snippets using `WScript.Shell.Run` to execute external binaries must include a clear ⚠️ warning callout and explain exactly what the command executes.

---

## Reporting a Vulnerability

If you discover a security issue in the automation scripts (`scripts/` directory) — for example, a path traversal vulnerability, insecure shell execution pattern, or hardcoded credentials — please **do not open a public issue**.

Instead:
1. Navigate to the **[Security Advisory](../../security/advisories/new)** section of this GitHub repository.
2. Submit a private security advisory describing:
   - The affected file(s) and line numbers
   - A description of the vulnerability and its potential impact
   - Steps to reproduce
   - Your suggested fix (if any)

You will receive an acknowledgment within **72 hours** and a resolution timeline within **7 days** for confirmed issues.

---

## Known Safe Patterns

The following libraries used in this vault's code snippets are standard, trusted Windows COM automation objects:

| Object String | Purpose | Risk Level |
| :--- | :--- | :--- |
| `Outlook.Application` | Outlook email automation | 🟡 Low (requires Outlook installed) |
| `Word.Application` | Word document generation | 🟡 Low (requires Word installed) |
| `PowerPoint.Application` | Slide deck export | 🟡 Low (requires PPT installed) |
| `WScript.Shell` | Run Python/system scripts | 🔴 **Review each use case** |
| `VBScript.RegExp` | Regex text parsing | 🟢 Safe (no I/O, no network) |
| `Scripting.FileSystemObject` | File/folder operations | 🟡 Low (scoped to project dirs) |
