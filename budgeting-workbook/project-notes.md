# Budgeting Spreadsheet Builder — Project Notes

## Current Architecture
This project consolidates what used to be 100+ sheets (e.g. BUILDING 1 to BUILDING 100) into a single worksheet that supports up to 150 buildings. Each building has multiple levels with inputs for perimeter, area, and cost systems.

The workbook is `.xlsm` with multiple active VBA modules and a custom reset utility.

## Goals Achieved
- ✅ Replaced separate building sheets with a scalable row-based input system
- ✅ Created dynamic building/level visibility using VBA (UpdateLevels)
- ✅ Built a `ResetModule` to clear values while preserving formulas, layout, and headers
- ✅ Removed unused features like roof area and streamlined logic for performance
- ✅ Reduced redundancy across VBA modules and improved naming conventions

## Key VBA Modules
- **ResetModule.bas** – Clears data fields without touching formulas
- **UpdateLevels.bas** – Shows/hides level rows based on number of buildings selected
- **ConditionalFormatting** – Highlights invalid or missing values (e.g. blank costs, 0 area)
- **MasterReset** – Ties everything together with a reset button on the main sheet

## Outstanding To-Do
- Add custom export features (e.g. to CSV or PDF)
- Build form interface for batch data entry (optional UX)
- Expand conditional formatting to flag error states in cost logic
- Refactor any remaining hard-coded cell references to named ranges

## Notes for Codex
- All inputs are on one main sheet
- VBA must never delete formulas or overwrite header rows
- Reset actions should always preserve calculated fields and formatting
- Modules should be reusable and readable by engineers with moderate Excel knowledge
