---
type: project-doc
project: Project 02 - Regional Sales Reporting Tool
doc_type: lessons-learned
status: in-progress
created: 2026-10-03
updated: 2026-10-03
---

# Project 02: Data Engineering & Analytics Insights

1. **Excel as an ETL Engine**: While Power Query is often used for ad-hoc ingestion, VBA provides superior control when the workflow requires creating dynamic outputs across multiple formats (PivotTable workbook, hardcoded report, and custom CSV) simultaneously.
2. **The Semicolon Delimiter Requirement**: European ERP systems often reject standard comma-separated files because commas represent decimal points in European regional settings. Controlling the delimiter with low-level VBA I/O eliminates external data transformation dependencies.
