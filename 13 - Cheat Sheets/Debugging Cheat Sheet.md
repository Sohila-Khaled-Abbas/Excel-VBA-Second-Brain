---
type: cheat-sheet
domain: VBA
topic: Debugging
tags:
  - cheat-sheet
  - debugging
  - vba
created: 2026-10-03
updated: 2026-10-03
---

# 🔍 Debugging & Troubleshooting Cheat Sheet

## Essential Hotkeys
- `F8`: Step Into (Executes line-by-line).
- `Shift + F8`: Step Over (Runs procedure call without entering it).
- `Ctrl + Shift + F8`: Step Out (Runs remainder of current procedure).
- `F9`: Toggle Breakpoint on current line.
- `F5`: Run / Continue execution.
- `Ctrl + G`: Open Immediate Window.
- `Ctrl + Break`: Force break an infinite loop.

## Immediate Window Commands (`Ctrl + G`)
- `?ActiveCell.Value`: Print active cell value.
- `?Sheets.Count`: Print number of sheets.
- `Application.ScreenUpdating = True`: Reset screen when stuck frozen.
- `Application.EnableEvents = True`: Reset events if disabled by error.
