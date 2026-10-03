---
type: concept
domain: VBA
topic: Fundamentals
difficulty: beginner
status: mastered
source: Udemy
course_section: "[[Section 05 - Working with Variables]]"
related_concepts:
  - "[[Variables]]"
  - "[[Dim Statement]]"
  - "[[Performance Optimization]]"
tags:
  - vba
  - data-types
  - memory
created: 2026-10-03
updated: 2026-10-03
---

# Data Types

> [!important]
> **Summary Statement**: Data types specify the kind of value a variable can store, determining its memory allocation, numeric boundaries, and execution efficiency.

---

## 📊 Comprehensive VBA Data Type Reference

| Data Type | Memory | Value Range | Recommended Usage |
| :--- | :---: | :--- | :--- |
| **`Long`** | 4 bytes | -2,147,483,648 to 2,147,483,647 | **Default for all row counts, integers, and counters** |
| **`Integer`** | 2 bytes | -32,768 to 32,767 | ⚠️ Avoid for row counts! (Excel has 1,048,576 rows; causes Overflow error 6) |
| **`Double`** | 8 bytes | ~1.79E+308 (15 digits precision) | **Default for financial amounts, decimals, scientific calculations** |
| **`Single`** | 4 bytes | ~3.40E+38 (7 digits precision) | Floating point when memory is constrained |
| **`Currency`**| 8 bytes | Fixed 4 decimal places | Highly accurate monetary math (prevents floating point rounding errors) |
| **`String`** | 10 bytes + length | Variable up to ~2 billion characters | Text, codes, names, paths |
| **`Boolean`** | 2 bytes | `True` or `False` | Flags, toggle states, conditional logic |
| **`Date`** | 8 bytes | Jan 1, 100 to Dec 31, 9999 | Dates, timestamps, elapsed timers |
| **`Variant`** | 16-22 bytes | Any data type or array | Flexible; required for copying entire cell ranges to memory arrays |
| **`Object`** | 4 bytes (ptr) | Reference to any COM object | Generic object pointer (`Workbook`, `Worksheet`, `Range`) |

---

## 🧠 Personal Understanding Layer
1. **Official / Technical Explanation**: VBA is strongly typed when declared explicitly; otherwise, it defaults to the universal `Variant` wrapper.
2. **My Understanding**: Always use `Long` for whole numbers and `Double` for decimals. Never use `Integer` because Excel sheets exceed 32,767 rows and will crash with `Overflow 6`.
3. **Simple Analogy**: Choosing data types is like packing boxes for moving. Using `Variant` for everything is like putting a single toothbrush into an enormous refrigerator box. It works, but wastes space and effort.
4. **Real Example**: Reading an entire 50,000-row table into a `Variant` array in 0.05 seconds, processing in memory, and dumping back in one shot.
5. **Where I Can Use It**: High-volume financial report aggregation and sales reconciliation.

---

## ⚠️ The Integer Overflow Trap
```vba
' DANGEROUS:
Dim r As Integer
For r = 1 To 40000 ' CRASHES: Run-time error '6': Overflow
Next r

' SAFE & PROFESSIONAL:
Dim r As Long
For r = 1 To 40000 ' Executes instantly and cleanly
Next r
```

---

## 🔗 Connected Notes
- **Parent Concept**: [[Variables]]
- **Array Mechanics**: [[Arrays]]
- **Optimization Strategy**: [[Performance Optimization]]
