---
type: concept
domain: VBA
topic: Events
difficulty: intermediate
status: learning
source: Udemy
course_section: "[[Section 16 - Worksheet and Workbook Events]]"
related_concepts:
  - "[[Worksheet Events]]"
tags:
  - vba
  - events
  - workbook
created: 2026-10-03
updated: 2026-10-03
---

# Workbook Events

> [!important]
> **Summary Statement**: Workbook events are procedures residing in the `ThisWorkbook` module that trigger when workbook-level lifecycle events occur (e.g., `Workbook_Open`, `Workbook_BeforeClose`, `Workbook_BeforeSave`).

---

## 💻 Auto-Configuring on Workbook Open
```vba
Private Sub Workbook_Open()
    ' Runs automatically whenever this workbook is opened
    MsgBox "Welcome to the Financial Model. Refreshing data connections...", vbInformation
    ThisWorkbook.RefreshAll
End Sub
```

---

## 🔗 Connected Notes
- [[Worksheet Events]], [[Workbook]]
