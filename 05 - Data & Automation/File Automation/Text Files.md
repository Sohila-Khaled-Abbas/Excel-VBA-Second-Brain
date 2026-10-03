---
type: concept
domain: VBA
topic: File I/O
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 12 - Working with Files, Folders & Text Files]]"
related_concepts:
  - "[[CSV]]"
  - "[[Files]]"
tags:
  - vba
  - text-files
  - io
created: 2026-10-03
updated: 2026-10-03
---

# Text Files & Low-Level I/O (`Print #`, `Write #`, `FreeFile`)

> [!important]
> **Summary Statement**: Low-level I/O streams text directly to and from `.txt` and `.csv` files using file channel numbers, giving complete control over custom delimiters and line endings.

---

## 💻 Writing Delimited Text Directly to Disk
```vba
Sub ExportCustomTextFile()
    Dim fNum As Integer: fNum = FreeFile
    Dim outPath As String: outPath = "C:\Reports\Export.txt"
    Dim r As Long, lastRow As Long
    
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Open outPath For Output As #fNum
    
    For r = 1 To lastRow
        Print #fNum, ActiveSheet.Cells(r, 1).Value & ";" & ActiveSheet.Cells(r, 2).Value
    Next r
    
    Close #fNum
End Sub
```

---

## 🔗 Connected Notes
- [[CSV]], [[Files]]
