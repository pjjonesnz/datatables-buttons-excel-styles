# DataTables Buttons Html5 Excel Styling

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/pjjonesnz/datatables-buttons-excel-styles)](https://github.com/pjjonesnz/datatables-buttons-excel-styles/releases)
[![GitHub license](https://img.shields.io/github/license/pjjonesnz/datatables-buttons-excel-styles)](https://github.com/pjjonesnz/datatables-buttons-excel-styles/blob/master/LICENSE.md)
[![npm](https://img.shields.io/npm/v/datatables-buttons-excel-styles)](https://www.npmjs.com/package/datatables-buttons-excel-styles)

**Add beautifully styled Excel output to your DataTables.**

[DataTables](https://www.datatables.net) is an amazing tool to display your tables in a user friendly way, and the [Buttons](https://www.datatables.net/extensions/buttons/) extension makes downloading those tables a breeze. 

Now you can simply style your downloaded tables without having to learn the intricacies of SpreadsheetML using:
* Pre-defined Templates: A selection of templates to apply to your table or selected cells, and/or
* Styles: Your own custom defined font, borders and backgrounds.

## Demo

[View the Excel styling demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html)

## Usage

1. Include jQuery

2. Include Datatables css and js (https://www.datatables.net/download/)

3. Include the plugins' style and optional template javascript files

```html
<script src="js/buttons.html5.styles.min.js"></script>
<script src="js/buttons.html5.styles.templates.min.js"></script>
```

4. Add an excelStyles config option to apply a style or template to the Excel file. [View demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/single_template_style.html)

```js
$('#myTable').DataTable({
    dom: 'Bfrtip',
    buttons: [
        {
            extend: 'excel',              // Extend the excel button
            excelStyles: {                // Add an excelStyles definition
                template: 'blue_medium',  // Apply the 'blue_medium' template
            },
        },
    ],
});
```

## Cell References

Use familiar Excel cell references to select a specific cell or cell range:
* `A2` - Select cell A2
* `C17` - Select cell C17
* `B3:D20` - Select the range from cell B3 to cell D20

Extended references are also available to select individual rows and columns, or row/column ranges:
* `4` - All cells in row 4
* `B` - All cells in column B
* `3:7` - All cells from (and including) row 3 to row 7
* `3:` - All cells from row 3 to the end of the table
* `>` - The last column in the table
* `-0` - The last row in the table
* `-2` - The third to last row in the table
* and more...

Smart References can select the various parts of the table (title, header, data, footer, etc.). These are enabled with a `s` prefix in the cell reference:
* `sh` - The header
* `sf` - The footer
* `s1` - Becomes the first data row
* `s-0` - Becomes the last data row
* `sB3` - Column B, row 3 of the data rows
* and more...

For exmplaes of using these cell selections, while the docs are being written, please [view the demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html), or have a look at the source of [buttons.html5.styles.templates.js](https://github.com/pjjonesnz/datatables-buttons-excel-styles/blob/master/js/buttons.html5.styles.templates.js)

** More docs underway - due end of May 2020 **
