WBS Roadmap Builder v3.0 – Official User & Technical Guide 
Release Date: March 2026 
Software Version: 3.0 
Author: Doug Brann (GitHub: @TheDougMiester)
License: MIT License
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Copyright (c) 2026 Doug Brann
________________________________________
PART 1: USER OPERATING GUIDE
1. Introduction
The WBS Roadmap Builder is an integrated Excel-based project management tool. It synchronizes a structured Work Breakdown Structure (WBS) with dynamic Roadmap and Gantt chart visualization engines to provide a unified, automated view of your project’s health.
2. The WBS Data Entry Tab
The WBS Tab is your command center. Access to this tab is semi-protected to ensure formula integrity while allowing data flexibility.

Core Configuration
•	Project Name: Located in Column F. This title will appear at the top of all generated charts.
•	Project Start: All tasks must occur on or after this date.
•	Current Date: Identifies the "Today" line on your charts. If set earlier than the Project Start, it defaults to the Start Date.
The "Golden Rules" of Data Entry
•	Level Management (Column B): Use the dropdown to set a task's level (1 to 6). The spreadsheet automatically calculates the WBS number (e.g., 1.2.1) and handles text indentation.
•	Event Types (Column H): Categorize milestones using the dropdown. These map to specific visual symbols:
o	△ / ▲: Milestone (Planned / Accomplished)
o	◇ / ◆: Decision Point (Planned / Accomplished)
o	▽ / ▼: Deliverable (Planned / Accomplished)
The Automated Calculation Engine (Columns J–N)
These columns are Read-Only. The software recalculates these dates based on:
1.	Dependencies: If Task B depends on Task A, its start date will "slip" to match Task A’s finish.
2.	Parent-Child Rollups: Top-level tasks (Level 1) automatically adopt the earliest start and latest finish dates of their nested sub-tasks.
3. Generating Your Charts
Click the "Build Roadmap" button to trigger the Plot Builder interface.

 
•	WBS Level: Choose the depth of the chart. Level 1 shows major phases; Level 3 provides granular detail.
•	Time Units: Toggle between Monthly or Quarterly timeline headers.
•	Show “Loopback” Dependencies: This is something that rarely (if ever) happens. When ticked, any item depending on something in its future will get a backwards arrow pointing from the later to the earlier event. With any luck, I’ve removed this feature, but I kept the tick box just in case the bug isn’t fully removed.
•	Show Start/Stop Column: Adds two columns with the earliest and latest dates of each WBS item
•	Show Slipped Markers: Displays red icons alongside scheduled markers to highlight delays.
•	Don't Plot Hidden Rows: If you have collapsed rows in Excel using the grouping (+/-) feature, the builder will ignore them.
•	Decide if you want the plot to use colored cells or shapes.

________________________________________
PART 2: ADVANCED LOGIC & CUSTOMIZATION
1. Dependency Management
The "Dependencies (WBS#s)" column (Column G) drives the scheduling engine.
•	Multiple Links: Use comma-separated WBS numbers (e.g., 2.1, 2.2).
•	Constraint: Only the lowest-level "leaf" tasks should have manual dependencies.
•	Circular References: If Task A depends on B and B depends on A, the cell will highlight Red and the link will be ignored.
2. Visual Coding
•	Slipped Schedule: Enter values in Slipped Start (Column K) and Slipped Duration (Column L). This creates a "Ripple Effect," pushing all dependent tasks forward and creating a red "impact" overlay on the charts.
•	Percent Done (Column N): Enter 0% to 100%. The cell heat-maps (Red → Yellow → Green) for an instant status check.
•	Start/End Colors (Columns O-P): Defines the color and pattern for bars occurring before and after the "Current Date."
________________________________________
PART 3: TECHNICAL MAINTENANCE (FOR VBA CODERS)
1. System Architecture
The tool utilizes a "Clean Slate" strategy. When a plot is generated, the existing sheet is deleted and a new one is scaffolded from scratch to prevent graphical artifacts.
•	Axis Engine (WbsTimeLineAxis): Converts calendar dates into Excel X-coordinates.
•	Header Engine (TimelineHeaderBuilder): Dynamically merges cells to create Fiscal Year (Oct-Sep) and Quarter tiers.
•	Validation (RegExTools): Uses Regular Expressions to parse WBS strings and dependencies.
2. VBA Reference Requirements
To maintain functionality, the following libraries must be enabled in Tools > References:
1.	Visual Basic for Applications
2.	Microsoft Excel 16.0 Object Library
3.	Microsoft Scripting Runtime (For Dictionaries)
4.	Microsoft VBScript Regular Expressions 5.5
3. Update Checklist
•	Versioning: Update VersionNumber.bas before distribution. This stamps the footer of every generated chart.
•	Performance: To speed up generation for large projects, the UnifiedPlotBuilder suppresses Application.ScreenUpdating and DisplayAlerts during the draw cycle.
•	Protection: Ensure the WBS tab is protected before deployment to hide calculation logic in Columns Q+.

