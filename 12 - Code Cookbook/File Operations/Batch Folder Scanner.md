---
type: snippet
language: VBA
category: files
difficulty: intermediate
use_case: Iterate through all Excel files in a target directory
objects:
  - Workbook
  - FileDialog
tags:
  - vba
  - snippet
  - file-io
created: 2026-10-03
updated: 2026-10-03
---

# Snippet: Batch Folder Scanner via `Dir()`

## 🎯 Problem Statement
Automates the ingestion of all `.xlsx` or `.csv` files stored in a designated local or network folder.

---

## 💻 VBA Code Implementation

```vba
Option Explicit

Public Sub ProcessAllWorkbooksInFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim wbCurrent As Workbook
    Dim processedCount As Long
    
    folderPath = "C:\DataImports\"
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    fileName = Dir(folderPath & "*.xlsx")
    
    Application.ScreenUpdating = False
    
    Do While Len(fileName) > 0
        ' Open in read-only background mode
        Set wbCurrent = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
        
        ' ----------------------------------------------------
        ' Perform custom data extraction / ETL here
        ' ----------------------------------------------------
        
        wbCurrent.Close SaveChanges:=False
        processedCount = processedCount + 1
        
        ' Fetch next matching file
        fileName = Dir()
    Loop
    
    Application.ScreenUpdating = True
    MsgBox "Successfully processed " & processedCount & " workbooks.", vbInformation, "Complete"
End Sub
```

---

## 🔗 Related Notes
- [[Files]], [[Folders]], [[Project 02 - Regional Sales Reporting Tool]]
