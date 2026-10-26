---
type: course-section
section_number: 20
section_title: Working with Charts
lecture_count: 5
status: not-started
tags:
  - course-section
  - vba
  - section-20
---

# Section 20: Working with Charts

> [!important]
> **Curriculum Pillar**: Section 20 contains **5 lectures**.  
> **Status**: `NOT-STARTED`  
> **Course Practice Files**: `SectionFiles/S20_Charts.xlsm`, `S20_Charts_Start.xlsx`

---

## 🎯 Section Objectives & Automation Capabilities
- Navigate the Chart Object Hierarchy: `Application -> Workbook -> Worksheet -> ChartObjects -> Chart`.
- Programmatically create, configure, source, and style charts using VBA.
- Create animated charts that step through monthly forecasts or budget comparisons using loops and `DoEvents`.
- Export chart objects as images (`Chart.Export`) and embed interactive charts inside UserForms.

---

## 🧠 Architectural Takeaways & Mental Models
- **ChartObjects vs Chart**:
  - `ChartObject`: The shape container embedded on a worksheet (controls dimensions, position: `.Left`, `.Top`, `.Width`, `.Height`).
  - `Chart`: The actual data visualization inside the container (controls `.ChartType`, `.SeriesCollection`, `.Axes`, `.ChartTitle`).
- **Programmatic Chart Creation**:
  ```vba
  Dim cho As ChartObject
  Set cho = ws.ChartObjects.Add(Left:=100, Top:=50, Width:=400, Height:=250)
  With cho.Chart
      .SetSourceData Source:=ws.Range("A1:B10")
      .ChartType = xlColumnClustered
  End With
  ```
- **Embedding Charts in UserForms**: Excel UserForms do not have a native chart control. The standard pattern exports the chart as a temporary image file (`cho.Chart.Export Filename:=tempPath, FilterName:="GIF"`) and loads it into an `Image` control (`Me.Image1.Picture = LoadPicture(tempPath)`).

---

## ⚠️ Common Pitfalls & Gotchas
> [!warning]
- Omitting `DoEvents` during animated chart loops will freeze the screen rendering, causing the animation to jump straight from start to finish without showing steps.
- Forgetting to delete temporary chart images exported to disk leaves clutter in user temp directories.

---

## 📚 Lecture Index

- [ ] [[Lecture S20-01 - The Chart Object Library & Section Overview|Lecture 01: The Chart Object Library & Section Overview]]
- [ ] [[Lecture S20-02 - Create and Modify a Chart|Lecture 02: Create and Modify a Chart]]
- [ ] [[Lecture S20-03 - Animated Charts Different Methods|Lecture 03: Animated Charts: Different Methods]]
- [ ] [[Lecture S20-04 - Show Charts inside UserForms|Lecture 04: Show Charts inside UserForms]]
- [ ] [[Lecture S20-05 - Key Takeaways Charts|Lecture 05: Key Takeaways: Charts]]

---

## 🔗 Section Knowledge Links
- **Parent Roadmap**: [[Course Roadmap]]
- **Course Overview**: [[Course Overview]]
- **Progress Tracking**: [[Progress Dashboard]]
- **Course Materials Guide**: [[Course Materials & Workbooks Guide]]
