---
type: post-mortem
topic: Control Flow
status: documented
tags:
  - failed-attempt
  - post-mortem
created: 2026-10-03
updated: 2026-10-03
---

# Failed Attempt 01: Top-Down Row Deletion Trap

## What Happened
Attempted to delete empty rows using a standard forward loop:
```vba
For r = 1 To lastRow
    If Cells(r, 1).Value = "" Then Rows(r).Delete
Next r
```

## Root Cause
When row 5 was deleted, row 6 shifted to become the new row 5. The loop incremented to `r = 6`, completely skipping the row that had just shifted up!

## Permanent Fix
Always loop backwards:
```vba
For r = lastRow To 1 Step -1
    If Cells(r, 1).Value = "" Then Rows(r).Delete
Next r
```
