# Advanced 'Excel Style' Cell References

You reference cells in a similar style to Excel cell references. Excel references have been extended to allow you to reference entire rows or columns.

Smart row references are available to refer to specific parts of the spreadsheet. Very useful when creating templates which may be applied to a range of tables that have different title/message/header/footer options enabled (or disabled).

## Single Cell/Row/Column References

| Ref | Meaning |
|---|---|
| A3 | cell A3 (just like Excel) |
| 4 | row 4, all columns (single number) |
| D | column D, all rows (single capital letter) |

## Range References

Select a range of cells by selecting a starting cell/row/column and an ending one, separated by a colon `:`

| Ref | Meaning |
|---|---|
| 4:6 | rows 4 to 6, all columns |
| B:F | column B to F, all rows |
| D4:D20 | column D from row 4 to 20 |

If nothing is on one side of the range, then it assumes the start, or the end of the table

| Ref | Meaning |
|---|---|
| 3: | from row 3 until the last row, all columns |
| A: | from column A until the last column, all rows |
| B3: | from column B until the last column, from row 3 until the last row |
| :B3 | from column A until column B, from row 1 to row 3 |
| B3:D | from column B until column D, from row 3 until the last row |

## Special references

Special references are available if you want to refer to the last row/column, or count back from the last row/column

### References to the last column

For columns, the greater than sign is used to refer to the last column (like an arrow pointing to the right).

| Ref | Meaning |
|---|---|
| > | all rows, the last column |
| >3:20 | the last column, row 3 to 20 (the number after the > refers to the row, just like A3 is row 3 of the first column, >3 is row 3 of the last column) |

### References counting back from the last column

To count back to the left from the last column, put a negative number before the greater than sign

| Ref | Meaning |
|---|---|
| -3> | three columns back from the last column |
| -2>5 | two columns back from the last column, row 5 |

### References counting back from the last row

| Ref | Meaning |
|---|---|
| -0 | all columns, the last row |
| B-3:B-0 | column B from the third to last row until the last row |

### Reference for everything

| Ref | Meaning |
|---|---|
| : | all columns, all rows (also `1:`, `A:`, `A1:`, `:-0`, `:>`, `:>-0`) |

## Column/Row skipping

This can be used to apply styles to every nth column or row (eg. every 2nd row, every 3rd column)

**Format: n[0-9],[0-9]**

n (stands for every nth column/row), then the column increment followed by the row increment

### Column/Row skipping examples

| Ref | Meaning |
|---|---|
| A3:D10n1,2 | from Column A row 3, to Column D row 10, target every column, target every second row |
| 3:n1,2     | every column from row 3 until the last row, target every second row (use this for row striping) |
| :n1,2      | every column, every second row (also `:n,2`) |
| :n2,1      | every second column, every row (also `:n2`) |

## Smart row references

With the default settings, row references refer to the actual spreadsheet rows (ie. 1 = row 1, 12 = row 12, etc.). This works well, but
can be hard to work with if your spreadsheet has (or doesn't have) the extra title and/or message above the data. Also, if you
include a footer this can be hard to define a template that works for custom excel configurations.

Smart row references adds specific code to refer to these special rows, and redefines row 1 to be the first row of the data

You can enable smart references by adding the `rowref:"smart"` option to your style definition

```js
excelStyles: [
    {
        rowref: "smart",
        cells: "...cell reference...",
        style: { 
            // ...style definition... 
        }
    }
]
```

An alternative way to enable smart row references is by adding a `s` to the beginning of the cell reference, although this may be harder to remember when you come back to it in the future. eg. `sD4` Now refers to column D, row 4 of the data (which in a standard table with a title row and column headers would be found in row 6 of the exported table).

### Available Smart Row References

| Ref | Buttons option | Meaning |
|---|---|---|
| t    | title: | the title row |
| m    | messageTop: | the message row |
| h    | header: | the row with the cell titles |
| 1    |   | the first data row (Numbered rows now **only** refer to the data, NOT the title, header, footer, etc.) |
| -0   |   | the last data row |
| :    |   | all of the data rows (also `1:` or `:-0` or `1:-0`) |
| f    | footer: | row with the cell titles (same content as the header row but at the bottom of the table) |
| b    | messageBottom: | the message row at the bottom of the table. The row below the footer row (it it is enabled) |

## Summary of the order of cell references

| Enable smart row refs | Move left of | Col reference | Smart row ref | Move up from last row | Row number | Range | To ref | Every nth | Col | and | Row |
|---|---|---|---|---|---|---|---|---|---|---|---|
| s | -2 | B | h | - | 3 | : | eg. A3 | n | 1 | , | 2 |