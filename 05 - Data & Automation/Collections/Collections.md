---
type: concept
domain: VBA
topic: Data Structures
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 06 - Looping through Collections & Making Decisions]]"
related_concepts:
  - "[[For Each]]"
  - "[[Arrays]]"
tags:
  - vba
  - collections
created: 2026-10-03
updated: 2026-10-03
---

# Collections

> [!important]
> **Summary Statement**: A `Collection` is an ordered set of items that can be referred to as a unit, supporting dynamic additions, removals, and keyed lookups (acting as an easy deduplication tool).

---

## 💻 Deduplicating Lists with Collections
```vba
Sub GetUniqueValues()
    Dim uniqueList As New Collection
    Dim cell As Range
    
    On Error Resume Next ' Duplicate keys throw error 457; ignore to deduplicate!
    For Each cell In ActiveSheet.Range("A2:A1000")
        If cell.Value <> "" Then
            uniqueList.Add cell.Value, CStr(cell.Value)
        End If
    Next cell
    On Error GoTo 0
    
    Debug.Print "Found " & uniqueList.Count & " unique items."
End Sub
```

---

## 🔗 Connected Notes
- [[Arrays]], [[For Each]]
