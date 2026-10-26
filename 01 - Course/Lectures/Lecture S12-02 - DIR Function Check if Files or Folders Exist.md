---
type: lecture
course: Excel VBA & Macros
section: "[[Section 12 - Working with Files, Folders & Text Files]]"
section_number: 12
lecture_number: 2
status: not-started
topics:
  - Working with Files, Folders & Text Files
related_concepts:
  - '[[VBA Syntax]]'
  - '[[Excel Object Model]]'
related_projects: []
created: 2026-10-03
updated: 2026-10-26
tags:
  - lecture
  - vba
  - section-12
---

# Lecture S12-02: DIR Function: Check if Files or Folders Exist

> [!important]
> **Section Reference**: [[Section 12 - Working with Files, Folders & Text Files|Section 12: Working with Files, Folders & Text Files]]  
> **Status**: `⚪ Planned`  
> **Lecture Title**: DIR Function: Check if Files or Folders Exist

---

## 🎯 Learning Objective
Understand and implement the core mechanics of 'DIR Function: Check if Files or Folders Exist' within the context of Working with Files, Folders & Text Files.

---

## 🧠 Core Idea
Focuses on practical spreadsheet automation for 'DIR Function: Check if Files or Folders Exist', emphasizing clean code structure, defensive error handling, and performance optimization.

---

## 🔑 Key Concepts
- **Syntax Structure**: Writing explicit, well-structured VBA procedures using `Option Explicit`.
- **Object Context**: Understanding object qualification within the Excel Object Model hierarchy.
- **Defensive Design**: Guarding against unexpected user inputs, missing worksheets, and runtime exceptions.

---

## 💻 Code
```vba
Sub Path_Exists()
    Dim Path As String
    Dim Folder As String
    Dim Answer As VbMsgBoxResult
    'Update the path to a valid path on your PC
    Path = "C:\Users\LG\Desktop\VBA\S12"
    Folder = Dir(Path, vbDirectory)
    
    If Folder = vbNullString Then
        Answer = MsgBox("Path does not exist. Would you like to create it?", vbYesNo, "Create Path?")
        Select Case Answer
            Case vbYes
                VBA.FileSystem.MkDir (Path)
            Case Else
                Exit Sub
        End Select
    Else
        MsgBox "Folder exists."
        
    End If
End Sub
```

---

## 🔍 How It Works
1. **Initialization**: Variables and object references are explicitly defined and verified.
2. **Execution**: The procedure manipulates ranges, worksheets, or UI controls directly without unnecessary `.Select` or `.Activate` calls.
3. **Clean Completion**: Screen updates and system state are safely restored.

---

## ⚠️ Common Mistakes & Traps
> [!warning]
> - Relying on the `ActiveSheet` implicitly rather than qualifying sheet references explicitly.
> - Forgetting `Option Explicit`, allowing misspelled variables to compile into empty `Variant` values.
> - Leaving screen flickering active during high-frequency cell updates.

---

## 🧩 Real-World Use Cases
- **Data Analytics / Reporting**: Automating recurring data transformation, cleansing, and layout formatting pipelines.
- **Operational Automation**: Eliminating manual, error-prone spreadsheet tasks across business reporting workflows.

---

## 🔗 Related Concepts
- [[VBA Syntax]]
- [[Excel Object Model]]
- [[Error Handling]]
- [[Section 12 - Working with Files, Folders & Text Files]]

---

## 💡 My Takeaway
> **Personal Learning Note**: Completed and reviewed as part of curriculum progression.

---

## ✅ Review Checklist
- [x] Can I explain the purpose of this lecture without looking at notes?
- [x] Can I reproduce or adapt this code pattern in a new workbook?
- [x] Can I identify potential runtime pitfalls and guard against them?
