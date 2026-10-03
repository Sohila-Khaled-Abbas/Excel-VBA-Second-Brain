---
type: concept
domain: VBA
topic: File I/O
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 12 - Working with Files, Folders & Text Files]]"
related_concepts:
  - "[[Folders]]"
  - "[[CSV]]"
related_projects:
  - "[[Project 02 - Regional Sales Reporting Tool]]"
tags:
  - vba
  - files
  - io
created: 2026-10-03
updated: 2026-10-03
---

# File Operations (`Dir`, `GetOpenFilename`, `FileDialog`)

> [!important]
> **Summary Statement**: VBA provides native tools to scan directories, prompt users for file selections, check file existence, and batch process external workbooks.

---

## 💻 3 Critical File Techniques

### 1. `Dir()`: Check if File Exists
```vba
If Dir("C:\Reports\Sales.xlsx") <> "" Then
    ' File exists
End If
```

### 2. `Dir()`: Loop All `.xlsx` Files in a Folder
```vba
Dim folderPath As String, fileName As String
folderPath = "C:\DataImports"
fileName = Dir(folderPath & "*.xlsx")

Do While fileName <> ""
    Debug.Print "Found: " & fileName
    ' Process file...
    fileName = Dir() ' Calls next matching file in folder
Loop
```

### 3. `Application.GetOpenFilename`: Let User Pick Files
```vba
Dim pickedFile As Variant
pickedFile = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm),*.xlsx;*.xlsm")
If pickedFile <> False Then
    Workbooks.Open pickedFile
End If
```

---

## 🔗 Connected Notes
- [[Folders]], [[CSV]], [[Project 02 - Regional Sales Reporting Tool]]
