# Height Regression Tool

A browser based tool for running height regressions on forestry inventory data. Replicates the Total Access Statistics (TAS) regression workflow: **ln(Ht) = a + b / DBH**, grouped by CrownCode.

Built for F&W Forestry Services timber inventory processing.

## What It Does

- Imports tree data from Excel files or pasted directly from Microsoft Access (Stands2, Plots2, Trees2)
- Joins the three tables automatically (Stands2 → Plots2 → Trees2) matching the standard 0_Plot_Tree make table query
- Groups trees by CoverNo, StndDesc, and YearOrigin for CrownCode assignment
- Runs OLS regressions of ln(Height) on 1/DBH for both Total Height and Merch Height by CrownCode
- Applies predicted heights to trees with missing measurements
- Exports results in the same format as the TAS/Access workflow

## How to Use

1. Open `index.html` in Chrome (or any modern browser)
2. Import your data using one of three methods:
   - **0_Plot_Tree Export**: A pre joined Excel file
   - **Separate .xlsx Files**: Upload Stands2, Plots2, and Trees2 as individual Excel files
   - **Paste from Access**: Copy/paste directly from each Access table (Ctrl+A, Ctrl+C)
3. Assign CrownCode groups using the grouping interface
4. Review QC flags (PulpHt close to TotalHt, MerchHt > TotalHt)
5. Run regressions and review results
6. Apply predictions and export

## Grouping Workflow

The tool sorts species the way you would work through them:

- **Planted pines first** (by species, then year descending)
- **Natural pines** next
- **300s and 400s** (hardwood groups by second digit)
- **500s, 600s, 700s** (hardwood groups by second digit)

Quick select buttons let you grab all 510s, 520s, etc. in one click. A filter dropdown lets you view only specific species groups. CrownCodes auto assign as 0, 1, ... 9, A, B, ... Z matching the TAS convention.

Default minimum of 30 measured total heights per group. Adjustable per job.

## Data Quality Checks

- **PulpHt within 10ft of TotalHt**: Flagged for review. Toggle to exclude from regression if it would pull the curve.
- **MerchHt > TotalHt**: Auto excluded from MHT regression.
- **CoverNo 0099**: Excluded from grouping.
- **ProdCode 9**: Excluded from grouping.
- **ProdCode 4 on 500s, 600s, 700s**: Excluded from grouping.
- **NoTrees > 1**: Excluded from grouping (tally trees).

All excluded trees remain in the dataset and receive predicted heights where applicable.

## Regression Model

```
ln(Ht) = a + b * (1 / DBH)
Ht = exp(a + b / DBH)
```

Produces all TAS statistics: Y-Intercept, Coef_of_X, SE of both coefficients, Beta, t-values, R squared, and sample count.

MHT regressions restricted to ProdCode = 3 (sawtimber).

## Exports

- **HtRegression.xlsx**: Full workbook with Charts data, Regressions (MHT/THT), Regress_Strata, Plot_Tree, and Trees2 sheets matching the TAS output format
- **Export Trees2**: Your original Trees2 table with only predicted heights and CrownCode filled in, ready to paste back into Access
- **Export Coefficients**: Standalone coefficient table
- **Export Charts**: Printable HTML file with scatter plots and fitted curves for every CrownCode

## Requirements

- A modern web browser (Chrome, Edge, Firefox)
- Internet connection on first load (for the SheetJS library and fonts)
- No installation, no server, no account needed

## Technical Notes

- 100% client side. No data leaves your browser.
- Uses [SheetJS](https://sheetjs.com/) for Excel read/write.
- Single HTML file with no dependencies to install.

---

*F&W Forestry Services, Inc.*
