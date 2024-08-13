# Google Apps Scripts by Patrick Goodwin
A collection of Google Apps Scripts I use, especially on Google Sheets to compare SQL outputs when Microsoft SQL server can't do it for me.

## Table of Contents
- [Compare Rows](#compare-rows)

## Compare Rows
Compares 2 rows and returns the differences.

### Parameters
- `row1`: The number of the first row to compare.
- `row2`: The number of the second row to compare.
- `sheetName`: the name of the sheet (defualt: "Sheet1")
- `includeTimestamps`: Displays timestamp differences if TRUE

### Example:
```
=compareRows(2, 3)
=compareRows(2, 3, "Sheet2")
=compareRows(2, 3, , TRUE)
=compareRows(2, 3, "Sheet2", TRUE)
```
