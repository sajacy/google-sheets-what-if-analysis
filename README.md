# What-If Analysis, Google Sheets Add-On

Excel-style Data Tables for sensitivity analysis and scenario evaluation in Google Sheets.

## Overview

When constructing complex spreadsheet models, outputs frequently depend on a sequence of formulas spread across multiple cells. In such models, evaluating different assumptions requires either duplicating the calculation structure across dozens of cells or testing scenarios manually.

What-If Analysis provides automated one-variable and two-variable Data Tables. The add-on substitutes test values into defined input cells, reads the resulting model outputs, and populates a structured results table.

All calculations run securely and privately in your Google Sheet.

## Getting Started

### Install Add-On

1. Install the add-on from the [Chrome Web Store](https://workspace.google.com/marketplace/app/whatif/430660745752).
2. Open your spreadsheet and click **Extensions > What-If Analysis > Open Sidebar** (or **What-If Analysis > Open Sidebar**).
3. Specify the table range and target input cells, then select **Save & Calc**.

### Manual Installation

Alternatively, you can add the source files directly to a Google Sheet's Apps Script project:

1. Within your Google Sheet, open **Extensions > Apps Script**.
2. Replace `Code.gs` (or create a script file) with the contents of [`src/WhatIfAnalysis.gs`](src/WhatIfAnalysis.gs).
3. Create an HTML file named `sidepanel` (`sidepanel.html`) and paste the contents of [`src/sidepanel.html`](src/sidepanel.html).
4. In **Project Settings**, enable **Show "appsscript.json" manifest file in editor**, and update it with [`src/appsscript.json`](src/appsscript.json).
5. Save the project and reload your Google Sheet.

### Working Example Spreadsheet

Make a copy of the [Reference Example Spreadsheet](https://docs.google.com/spreadsheets/d/1kQ8y3DsnCilScNda1AIcV_-M7TIAWSkdlF__ceZ-W7U/edit?gid=1909248277#gid=1909248277) to inspect working models for:

- **Loan Amortization:** Evaluating interest rates and terms against monthly payments and cumulative interest.
- **Tax Liability Schedules:** Measuring progressive tax impacts across filing statuses and income tiers.

## Example Configurations

### 1. Two-Variable (Bivariate) Matrix

Evaluates a single model output against two variable inputs arranged along the row and column axes:

| Formula: `=B10` (Model Output) | Term: 10 Years | Term: 20 Years | Term: 30 Years |
| :--- | ---: | ---: | ---: |
| **APR: 3.0%** | -$34,794.53 | -$111,214.52 | -$128,497.39 |
| **APR: 4.0%** | -$46,741.94 | -$152,991.85 | -$178,246.28 |
| **APR: 5.0%** | -$58,808.37 | -$196,762.13 | -$231,134.73 |

- Place your model output formula or cell reference in the top-left cell of the table range.
- Place row test values across the top header row.
- Place column test values down the leftmost column.
- Select the table range, assign the corresponding input cells in the sidebar, and run the calculation.

### 2. One-Variable (Univariate) Sweep

Varies a single input variable down a column while computing one or more model outputs simultaneously. Place your target output formulas in the top row directly above the result columns:

| [Optional Label] | Formula 1: `=B6` (Monthly Pmt) | Formula 2: `=B10` (Total Interest) |
| :--- | ---: | ---: |
| **Test Value: 3.0%** | $1,051.38 | -$128,497.39 |
| **Test Value: 4.0%** | $1,189.57 | -$178,246.28 |
| **Test Value: 5.0%** | $1,336.49 | -$231,134.73 |
| **Test Value: 6.0%** | $1,491.42 | -$286,910.92 |

- **Where formulas go:** Put your formula (or a direct cell reference to the model output, such as `=B6` or `=B10`) in the **top row** of each output column (columns 2, 3, etc.). The tool reads these formulas to calculate what fills in the cells below.
- **Where test values go:** Enter your input values down the **leftmost column** starting in row 2.
- **Top-left cell:** Can be left blank or contain a descriptive text header for the input column.
- **Sidebar setup:** Designate the driving variable cell in **Column Input**, leave **Row Input** blank, and select **Save & Calc**.
- **Row-oriented sweeps:** If your test values run horizontally across a row instead, place your formulas down the leftmost column, test values across the top row, and assign the target cell to **Row Input**.

## Key Features

### Named Ranges and Interactive Range Tools

- **Named Range Resolution:** Accepts workbook and sheet-scoped Named Ranges (such as `Example1DataTableRange` and `InterestRate`). If you insert rows or restructure sheets, Named Ranges prevent references from breaking.
- **One-Click Range Capture:** Click the grid icon next to any field to pull the active sheet selection into your configuration.
- **Cell Focus:** Click the eye icon next to any coordinate to jump directly to the target range in the sheet.

### Power Tools and Portability

- **Configuration Portability:** View and export all table configurations as structured JSON records to back up scenarios or migrate setups between workbooks.
- **Revert to Last Save:** Discard experimental table edits and restore the last saved configuration state with one click.
- **Batch Management:** Save, delete, rename, copy, and reorder multiple tables across sheets from a single unified panel.

## Practical Modeling Considerations

1. **When to use Data Tables versus direct formulas:** For simple single-step relationships, writing direct formulas across the grid calculates faster. What-If Data Tables are designed for multi-cell dependencies, loan schedules, or tiered bracket lookups where condensing the logic into a single formula is impractical or difficult to audit.
2. **Multi-output evaluation in one sweep:** In one-variable tables, multiple formulas can be positioned across adjacent output columns. A single parameter sweep can evaluate multiple related metrics side by side without setting up separate tables.
3. **Manual recalculation preserves performance:** Tables do not recalculate on every keystroke. This prevents large scenario tables from introducing spreadsheet calculation latency during normal edits. When base assumptions change, recalculation is triggered on demand through the sidebar.

## Development & DevOps

### Project Structure

- [`src/WhatIfAnalysis.gs`](src/WhatIfAnalysis.gs) — Google Apps Script backend (`V8` runtime, `@OnlyCurrentDoc`)
- [`src/sidepanel.html`](src/sidepanel.html) — Sidebar UI (HTML/CSS/JS)
- [`src/appsscript.json`](src/appsscript.json) — Apps Script manifest and minimal OAuth scopes
- [`test/what-if-analysis.test.js`](test/what-if-analysis.test.js) — Automated unit/integration tests using Node's built-in test runner

### Build

No compilation or bundling step is required. The files in [`src/`](src/) are plain Google Apps Script (`.gs`, `.html`, and `appsscript.json`) and are deployed directly.

### Test

The test suite uses Node.js's built-in `node:test`, `node:assert/strict`, and `node:vm` modules with no external dependencies.

```bash
node --test test/what-if-analysis.test.js
```

### Deploy

You can deploy changes to Google Apps Script using [`clasp`](https://github.com/google/clasp):

1. Install and authenticate `clasp` (one-time setup):

   ```bash
   npm install -g @google/clasp
   clasp login
   ```

2. Point `.clasp.json` (git-ignored or local) to your Apps Script project with `rootDir` set to `src`:

   ```json
   {
     "scriptId": "<YOUR_APPS_SCRIPT_ID>",
     "rootDir": "src"
   }
   ```

3. Push source files and create a new versioned deployment:

   ```bash
   clasp push
   clasp deploy -d "Describe release"
   ```

For Workspace Marketplace updates, create a new deployment version in Apps Script and update the deployment version in the Google Workspace Marketplace SDK configuration in Google Cloud Console.

## Privacy & License

- [Privacy Policy](PRIVACY.md)
- [MIT License](LICENSE) — Copyright (c) 2018 Jeffrey Yang
