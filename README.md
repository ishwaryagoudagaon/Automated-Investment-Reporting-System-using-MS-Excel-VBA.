Investment Reporting Dashboard (Excel & VBA)

Overview

This project is a personal learning project that demonstrates how Excel
can be used to organize investment data, summarize portfolio
information, and present reporting metrics in a dashboard.

The workbook includes investment data, PivotTables, slicers, charts, and
a VBA macro to support the report-refresh workflow. It was built to
practice reporting, data analysis, dashboard design, and basic
automation skills relevant to investment reporting.

Important: This is a demonstration project using illustrative
data. Any FX rates or benchmark figures in
the workbook should be treated as illustrative unless independently
verified.

Project goals

Organize investment information in a structured Excel workbook.

Summarize data using PivotTables.

Explore portfolio information by asset class, region, investment,
and reporting period.

Present key figures and trends with charts and dashboard elements.

Practise Excel slicers and interactive filtering.

Use VBA to reduce repetitive report-refresh steps.

Features

Investment dataset: Example investment records with fields such
as reporting date, investment name, asset class, region, currency,
units, prices, market values, profit/loss, returns, and benchmark
returns.

PivotTables: Summaries that support comparisons across
categories and reporting periods.

Interactive slicers and filters: Allow users to filter selected
PivotTable summaries.

Dashboard visualizations: Charts and summary views for exploring
portfolio value, returns, asset classes, regions, and related
metrics.

VBA refresh macro: A button runs a macro that requests a
workbook refresh, recalculates formulas, and displays a success or
error message.

Tools and skills demonstrated

Microsoft Excel

PivotTables and PivotCharts

Slicers and filters

Excel formulas and KPI reporting

Dashboard layout and data visualization

VBA macro basics and error handling

Data validation and analytical thinking

How to use

Download or clone this repository.

Open Automated_Investment_Reporting_100_Rows(3).xlsm in the
desktop version of Microsoft Excel.

If you trust the workbook and want to test the macro, enable macros
when prompted. Only enable macros in files you trust.

Review the investment data and explore the dashboard.

Use slicers and PivotTable filters to examine different categories
and reporting periods.

Click the Refresh Investment Report button to run the refresh
macro.

If you change the source data, confirm the PivotTable source
range/table and refresh the relevant PivotTables.

Note: Excel for the web and some spreadsheet applications may not
support running VBA macros. Use desktop Excel for the macro
functionality.

VBA macro behavior

The refresh macro is designed to: 1. Temporarily turn off screen
updating. 2. Call This Workbook.RefreshAll. 3. Recalculate workbook
formulas. 4. Display a confirmation message. 5. Display an error message
if the macro encounters a handled error.

The macro is a basic learning example. Depending on workbook connections
and refresh behavior.

Suggested repository structure

```text
Investment-Reporting-Dashboard/
│
├── readme.md
└── Automated_Investment_Reporting.xlsm
├──Dashboard.pdf
├── image
```
## Dashboard Demo

Watch the video demonstration of the Excel Investment Reporting Dashboard:

[Watch Dashboard Demo](PASTE_YOUR_VIDEO_URL_HERE)


Limitations

Data is illustrative and should not be used to make financial or
investment decisions.

FX rates and benchmark returns are not guaranteed to be live or
externally validated.

The workbook is intended to demonstrate Excel reporting and basic
VBA automation, not a production-grade financial reporting system.

The macro refreshes workbook content but does not independently
validate the correctness of source data or financial calculations.

Future improvements

Add automated checks for missing values, duplicate records, and
invalid dates.

Improve the handling and display of blank categories in PivotTables
and slicers.

Standardize number formats for currency, returns, and negative
values.

Add a clear reporting-date selector and chronological month sorting.

Expand VBA automation with refresh-status checks and validation
messages.

Document calculation definitions and data assumptions.

Author

Created as a personal learning project while developing skills in data
analytics, Excel reporting, and VBA automation.

