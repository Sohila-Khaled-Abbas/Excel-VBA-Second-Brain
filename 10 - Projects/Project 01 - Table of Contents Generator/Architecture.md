---
type: project-doc
project: Project 01 - Table of Contents Generator
doc_type: architecture
status: completed
created: 2026-10-03
updated: 2026-10-03
---

# Project 01: Architecture & Data Flow

```mermaid
flowchart TD
    A[User clicks 'Generate TOC' Macro] --> B[Disable ScreenUpdating & Alerts]
    B --> C{Does 'Table of Contents' tab exist?}
    C -- Yes --> D[Delete old TOC sheet / Clear range]
    C -- No --> E[Create new Worksheet at Index 1]
    D --> E
    E --> F[Apply Header Styling & Formatting]
    F --> G[Loop through each Worksheet in ActiveWorkbook]
    G --> H{Is sheet the TOC sheet?}
    H -- Yes --> I[Skip]
    H -- No --> J{Is sheet Visible?}
    J -- Visible --> K[Write Sheet Name & Add Hyperlink]
    K --> L[Add 'Back to TOC' link on Sheet Cell A1]
    J -- Hidden --> M[Optionally list with '(Hidden)' flag or skip]
    L --> N[Increment TOC row counter]
    N --> G
    G --> O[AutoFit TOC Columns & Style Table]
    O --> P[Restore ScreenUpdating & Alerts]
    P --> Q[Select Cell A1 on TOC Sheet]
```
