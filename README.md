# Excel Diff

This project compares two ranges of cells in an Excel workbook and highlights the differences in a new worksheet named `Diffs`.

## What it does

- Opens `Employee Sales.xlsx`
- Prompts you for the names of two worksheets and the cell ranges to compare
- Verifies that both ranges have the same shape
- Compares each matching cell value
- Writes the differing rows to a new `Diffs` sheet
- Includes a percent-difference formula for numeric mismatches

## Requirements

Install the required dependency:

```bash
pip install openpyxl
```

## How to use

1. Place your workbook in the same folder as `exceldiff.py`.
2. Run the script:

```bash
python exceldiff.py
```

3. Enter:
   - the first worksheet name
   - the first cell range, such as `A1:E10`
   - the second worksheet name
   - the second cell range with the same dimensions

4. The script saves the results back to `Employee Sales.xlsx` and adds a `Diffs` worksheet.

## Notes

- The selected ranges must be the same size.
- The comparison ignores the first row of each range and treats it as the header row.
- Numeric differences are calculated with a percentage-based formula in the `Diffs` sheet.