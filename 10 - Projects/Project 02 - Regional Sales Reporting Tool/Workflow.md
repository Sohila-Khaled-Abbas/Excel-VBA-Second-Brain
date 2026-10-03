---
type: project-doc
project: Project 02 - Regional Sales Reporting Tool
doc_type: workflow
status: in-progress
created: 2026-10-03
updated: 2026-10-03
---

# Project 02: Detailed Pipeline Workflow

1. **Initialization & Optimization**:
   - Turn off `ScreenUpdating`, `DisplayAlerts`, and switch `Calculation` to manual.
2. **File Selection**:
   - Open FileDialog with filters for `*.xlsx; *.xlsm`.
   - Validate that the user selected at least one file.
3. **Batch Import & Consolidation**:
   - Clear existing rows in `Consolidated_Data` tab (retaining headers).
   - Loop through array of selected paths.
   - For each workbook, read the `CurrentRegion` and copy values to next available row in master sheet.
4. **Pivot Cache Invalidation & Refresh**:
   - Dynamically update the `PivotTable.SourceData` range to cover the new consolidated row count.
   - Call `PivotCache.Refresh`.
5. **Dual Report Export**:
   - **Interactive Excel File**: Creates new workbook with active PivotTable for internal analysts.
   - **Hardcoded Excel File**: Converts all formulas and pivots to static values for external auditors.
6. **Semicolon CSV Generation**:
   - Loop through consolidated records using `FreeFile` and `Print #`.
   - Explicitly format date as `YYYY-MM-DD` and concatenate with `;`.
7. **Task Tracking & State Restoration**:
   - Append execution metadata (timestamp, user, file count, total revenue) to `Activity_Log`.
   - Restore application environment to automatic.
