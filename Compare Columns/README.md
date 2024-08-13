# Compare Columns
Compares 2 columns and returns the differences.

## Parameters
- `column1`: The number of the first row to compare - must be in string format.
- `column2`: The number of the second row to compare- must be in string format.
- `sheetName`: the name of the sheet (defualt: "Sheet1")
- `includeTimestamps`: Displays timestamp differences if TRUE

## Examples:
```
=compareColumns("B", "C")
=compareColumns("B", "C", "Sheet2")
=compareColumns("B", "C", , TRUE)
=compareColumns("B", "C", "Sheet2", TRUE)
```
