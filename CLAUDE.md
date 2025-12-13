# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Modern Excel PMS (Project Management System) - A pure Excel-based project management tool that aims to provide Jira/Asana-like UX using only Excel formulas, conditional formatting, and minimal VBA. The system manages a 3-tier hierarchy: Cases (projects) -> Measures (initiatives) -> WBS tasks.

## Build Commands

```bash
# Generate the macro-enabled workbook (ModernExcelPMS.xlsm)
python3 tools/build_workbook.py

# Generate with multiple project sheets and sample data
python3 tools/build_workbook.py --projects 3 --sample-first --output ModernExcelPMS.xlsm

# Generate with report output
python3 tools/build_workbook.py --projects 3 --sample-first \
  --report-output workbook_report.md --pdf-output workbook_report.pdf

# Generate with custom sheet protection password
PMS_SHEET_PASSWORD="custom-password" python3 tools/build_workbook.py
```

No external Python dependencies required - the build script uses only the standard library with direct OpenXML generation.

## Architecture

### Hub & Spoke Data Model
- **Hub sheets**: `Case_Master` (portfolio view), `Measure_Master` (project list)
- **Spoke sheets**: `PRJ_xxx` (detailed WBS cloned from `Template`)
- Data flows upward via `INDIRECT` references; each WBS exposes progress in cell `J2`

### Sheet Structure
| Sheet | Purpose |
|-------|---------|
| `Config` | Holidays, member master, status dropdown values |
| `Template` | WBS template (hidden in production); copied to create new `PRJ_xxx` |
| `PRJ_xxx` | Individual WBS with Gantt, auto-calculated status, progress tracking |
| `Case_Master` | Case portfolio with drill-down filter to linked measures |
| `Measure_Master` | Links cases to WBS sheets; pulls progress via `INDIRECT` |
| `Kanban_View` | Dynamic kanban board using `FILTER`/`MAP`; double-click updates status |

### Key Formulas
- **Progress aggregation**: `SUMPRODUCT(effort, progress) / SUM(effort)` for weighted average
- **End date**: `WORKDAY(start, effort-1, holidays)` accounting for holidays
- **Status auto-detection**: `IFS` checking completion %, due date, start date
- **Kanban cards**: `FILTER` + `MAP` + `HSTACK` spilling task cards by status

### VBA Modules (in `docs/vba/`)
- `modWbsCommands.bas`: Row swap (Up/Down), template duplication, status update from Kanban
- `modProtection.bas`: Sheet protection utilities
- `Kanban_View.bas`: Double-click event handler
- `ThisWorkbook.bas`: `PRJ_xxx` sheet name numbering

## Build Script Details

`tools/build_workbook.py` generates `.xlsm` directly by writing OpenXML ZIP structure:
- Creates all sheets with formulas, data validations, conditional formatting
- Embeds VBA from `docs/vba/*.bas` files
- Applies sheet protection with password hash
- Outputs progress report in Markdown and/or PDF format

Modify sheet content by editing the `*_sheet()` functions and `template_cells()` in the build script.

## Report Generation

When using `--sample-first` or `--sample-all`, the report includes progress analytics:

```bash
python3 tools/build_workbook.py --projects 2 --sample-first \
  --report-output report.md --pdf-output report.pdf
```

Report contents:
- **Progress Summary**: Weighted progress rate, total/completed effort
- **Status Breakdown**: Task counts by status with visual bar chart
- **Completion Rate**: Tasks completed vs total (案件消化度)
- **Measure Progress**: Per-project progress rates (施策別進捗)
- **Owner Workload**: Effort and completion by assignee (担当者別負荷)
- **Master Data**: Cases, measures, statuses, members

Key functions in build script:
- `SampleTask`: Data class for task information
- `calculate_weighted_progress()`: Effort-weighted progress calculation
- `count_by_status()`: Status aggregation
- `generate_report_lines()`: Report text generation

## Documentation

- `docs/IMPLEMENTATION_PLAN.md`: Phase-by-phase implementation status
- `docs/sheet_protection_plan.md`: Protection settings and password policies
- `docs/sheet_protection_test.md`: Test scenarios for verifying sheet protection
- `docs/recalc_performance_report.md`: Performance benchmarks for large WBS
- `docs/gantt_palette.svg`: Color scheme reference for Gantt/status formatting

## Sheet Protection

The build script automatically applies sheet protection with:
- Password-protected sheets (default: `pms-2024`, override via `PMS_SHEET_PASSWORD` env var)
- Unlocked cells for user input (task data, master lists, selectors)
- Locked cells for formulas and calculated fields
- `Workbook_Open` event re-applies protection with `UserInterfaceOnly:=True` for VBA compatibility

## Language

Documentation and comments are in Japanese. The system targets Microsoft 365 Excel (dynamic array functions required).
