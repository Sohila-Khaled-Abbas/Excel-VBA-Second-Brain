---
type: concept
domain: VBA
topic: File I/O
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 12 - Working with Files, Folders & Text Files]]"
related_concepts:
  - "[[Files]]"
tags:
  - vba
  - folders
created: 2026-10-03
updated: 2026-10-03
---

# Folder Automation (`FileDialog(msoFileDialogFolderPicker)`)

> [!important]
> **Summary Statement**: Uses Office's built-in folder picker dialog to allow users to select target directories dynamically.

---

## 💻 Folder Picker Implementation
```vba
Dim fDialog As FileDialog
Dim selectedFolder As String

Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
fDialog.Title = "Select Directory Containing Regional Sales Files"

If fDialog.Show = -1 Then
    selectedFolder = fDialog.SelectedItems(1) & ""
Else
    Exit Sub ' User cancelled
End If
```

---

## 🔗 Connected Notes
- [[Files]], [[File Automation]]
