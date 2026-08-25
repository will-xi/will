---
name: xlsx
description: Spreadsheet creation, editing, analysis, formula handling, formatting, and visualization for .xlsx, .xlsm, .csv, and .tsv files. Use this skill when Codex needs to create a spreadsheet, analyze tabular data, modify an existing workbook while preserving formulas and formatting, or validate spreadsheet calculations.
---

# XLSX Skill

## Purpose
Use this skill for spreadsheet tasks involving Excel-compatible files. Prefer reliable, editable workbooks over hardcoded outputs.

## Core rules
- Preserve the existing workbook structure, formulas, styles, merged cells, widths, number formats, and conventions unless the user explicitly asks for a redesign.
- When a workbook already has an established visual or modeling convention, follow it instead of imposing a new standard.
- Keep formulas dynamic. Do not calculate results in Python and write fixed values when an Excel formula is appropriate.
- Before delivery, check for formula errors such as `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, and `#NAME?`.
- Never overwrite the only copy of a source workbook unless the user explicitly requests in-place modification.

## Tool selection
### pandas
Use pandas for:
- reading and profiling tabular data;
- filtering, grouping, joining, reshaping, and bulk transformations;
- statistical analysis;
- fast CSV/TSV processing;
- exporting simple result tables.

### openpyxl
Use openpyxl for:
- preserving and editing existing `.xlsx` workbooks;
- formulas;
- cell styles, fills, fonts, borders, alignment, dimensions, comments, validation, and worksheet structure;
- creating polished Excel outputs.

For macro-enabled `.xlsm` files, load with `keep_vba=True` when macros must be preserved.

## Standard workflow
1. Inspect the workbook before editing: sheet names, dimensions, formulas, tables, merged cells, hidden sheets, named ranges, styles, and number formats.
2. Decide whether pandas, openpyxl, or both are appropriate.
3. Make the requested changes with the smallest necessary scope.
4. Use Excel formulas for totals, percentages, ratios, differences, lookups, and other calculations that should update when source data changes.
5. Save to a new output file unless the user explicitly asks to modify the original.
6. Recalculate formulas using the available spreadsheet engine if needed.
7. Re-open and verify the workbook, including key formulas, styles, sheet integrity, and obvious errors.

## Formula practices
- Prefer cell references over embedded constants when assumptions may change.
- Check row and column offsets carefully; Excel indices are 1-based.
- Validate cross-sheet references and quoted sheet names.
- Guard divisions where zero denominators are possible.
- Test representative formulas before filling large ranges.
- Check for accidental circular references.
- Ensure copied formulas remain consistent across periods and rows.

## Existing workbook editing
When modifying a user-provided workbook:
- inspect neighboring cells before changing formats;
- copy established styles rather than recreating approximate ones;
- do not rename sheets, tables, ranges, or headers unless required;
- do not delete formulas or replace them with values unintentionally;
- preserve hidden rows/columns and hidden sheets unless the request specifically changes them;
- be cautious with merged cells and structured tables.

## New workbook quality
For newly created workbooks:
- use clear sheet names and headers;
- freeze panes when helpful;
- set practical column widths and row heights;
- apply consistent number/date/percentage formats;
- use filters or Excel tables for datasets when appropriate;
- keep input, calculation, and output areas visually distinguishable when the workbook is model-like;
- avoid decorative formatting that reduces readability.

## Financial-model convention
Only when creating a financial model from scratch and there is no existing convention:
- hardcoded user inputs may use blue font;
- formulas may use black font;
- same-workbook linked values may use green font;
- external-workbook links may use red font;
- highlight important assumptions that require review.

Use sensible number formats:
- percentages: usually one decimal place;
- currency: include units in headers where scale matters;
- negatives: prefer parentheses in formal financial models;
- zeros: may be displayed as dashes in financial reports;
- multiples: use an `x` suffix where appropriate.

## Data integrity checks
Before finalizing:
- confirm expected row and column counts;
- check duplicate keys where uniqueness is expected;
- check missing values in critical fields;
- verify date and numeric columns were not converted to text accidentally;
- confirm formulas reference the intended cells;
- compare a few calculated values manually or independently;
- scan all sheets for obvious spreadsheet error values.

## Large files
- Use `read_only=True` for large-workbook inspection when appropriate.
- Use `write_only=True` for very large generated workbooks when formatting needs are simple.
- Limit pandas reads to required columns or sheets when memory is a concern.
- Preserve IDs as strings when leading zeros matter.

## Important openpyxl caution
`data_only=True` is useful for reading cached calculated values, but do not save a workbook loaded only for cached values if formulas must be preserved. Keep a normal formula-preserving workbook instance for edits.

## Output expectations
- Deliver a valid spreadsheet that opens without repair warnings.
- Preserve existing formulas and formatting unless the task says otherwise.
- Prefer formulas over hardcoded calculated values.
- Keep the workbook understandable and editable by a human.
- Report any limitation that prevents full verification, such as unavailable formula recalculation or unsupported external links.
