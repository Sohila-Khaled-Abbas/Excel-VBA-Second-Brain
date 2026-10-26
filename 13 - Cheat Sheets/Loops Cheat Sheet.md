---
type: cheat-sheet
domain: VBA
topic: Loops & Iteration
tags:
  - cheat-sheet
  - loops
  - control-flow
created: 2026-10-03
updated: 2026-10-26
---

# 🔄 Loops & Iteration Cheat Sheet

---

## 📊 Loop Constructs Comparison

| Loop Type | When to Use | Syntax Template | Performance Notes |
| :--- | :--- | :--- | :--- |
| **`For Each...Next`** | Iterating over all objects in a collection | `For Each ws In Worksheets`<br>`Next ws` | **Fastest & cleanest** for objects; avoids index lookups |
| **`For...Next` (Forward)** | Counting from fixed start to end | `For i = 2 To lastRow`<br>`Next i` | Standard forward traversal; clean index access |
| **`For...Next` (Reverse)** | **Deleting rows** or bottom-up operations | `For i = lastRow To 2 Step -1`<br>`Next i` | **Mandatory for deletions** to prevent index shifts |
| **`Do While...Loop`** | Looping as long as condition remains `True` | `Do While Cells(r, 1) <> ""` <br>`    r = r + 1`<br>`Loop` | Pre-checks condition before entering |
| **`Do Until...Loop`** | Looping until condition becomes `True` | `Do Until Cells(r, 1) = ""` <br>`    r = r + 1`<br>`Loop` | Runs until termination condition is met |

---

## ⚠️ The Row Deletion Trap & Reverse Loop Solution

> [!warning]
> If you delete rows moving forward (`For i = 2 To lastRow`), deleting row 5 causes row 6 to become row 5. The loop moves to index 6 on the next turn, skipping the newly shifted row entirely!

```vba
' CORRECT: Loop backwards with Step -1
Dim lastRow As Long, i As Long
lastRow = Cells(Rows.Count, "A").End(xlUp).Row

For i = lastRow To 2 Step -1
    If Cells(i, "D").Value = "Delete" Or Cells(i, "D").Value = 0 Then
        Rows(i).Delete
    End If
Next i
```

---

## 🔍 Range.Find + Do Loop Pattern (Lightning Fast)

```vba
Dim foundCell As Range, firstAddress As String
Set foundCell = ws.Range("A:A").Find(What:="Target", LookIn:=xlValues, LookAt:=xlWhole)

If Not foundCell Is Nothing Then
    firstAddress = foundCell.Address
    Do
        ' Process matching cell
        foundCell.Interior.Color = vbYellow
        
        ' Find next occurrence
        Set foundCell = ws.Range("A:A").FindNext(foundCell)
    Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
End If
```
