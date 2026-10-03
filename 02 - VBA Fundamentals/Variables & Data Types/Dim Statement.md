---
type: concept
domain: VBA
topic: Syntax
difficulty: beginner
status: mastered
source: Udemy
course_section: "[[Section 05 - Working with Variables]]"
related_concepts:
  - "[[Variables]]"
  - "[[Data Types]]"
  - "[[Option Explicit]]"
tags:
  - vba
  - syntax
  - declaration
created: 2026-10-03
updated: 2026-10-03
---

# Dim Statement

> [!important]
> **Summary Statement**: `Dim` (Dimension) is the VBA keyword used to declare variables and allocate storage space in memory.

---

## 📖 Syntax & Structure
```vba
Dim variableName As DataType
```

- If `As DataType` is omitted, VBA silently declares the variable as `Variant`.
- Multiple declarations on one line must each specify their type:
  ```vba
  Dim i As Long, j As Long, k As Long
  ```

## 🧠 Personal Understanding Layer
- **Historical Note**: `Dim` stands for "Dimension", originally used in early BASIC to dimension the size of arrays.
- **Best Practice**: Declare variables at the top of the procedure or immediately before their first logical block of use.

---

## 🔗 Connected Notes
- **Rule Enforcer**: [[Option Explicit]]
- **Scope Levels**: [[Scope]]
