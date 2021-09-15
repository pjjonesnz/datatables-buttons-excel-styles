# DataTables Buttons Excel Styling

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/pjjonesnz/datatables-buttons-excel-styles)](https://github.com/pjjonesnz/datatables-buttons-excel-styles/releases)
[![GitHub license](https://img.shields.io/github/license/pjjonesnz/datatables-buttons-excel-styles)](https://github.com/pjjonesnz/datatables-buttons-excel-styles/blob/master/LICENSE.md)
[![npm](https://img.shields.io/npm/v/datatables-buttons-excel-styles)](https://www.npmjs.com/package/datatables-buttons-excel-styles)

**Add beautifully styled Excel output to your DataTables**

* Style the font, border, background, number-format of your table
* Target cells, columns, rows, headers, footers, title, message and messageBottom
* Target cell ranges using familiar Excel cell references
* Conditional formatting
* Insert or replace cells, columns and rows
* Set the default page printing options using a pageStyle config option
* Easy JSON configuration options
* Simple templates built in for fast styling
* Currency formatting, fix the standard US$ display
* Smart targeting of rows useful for styling a spreadsheet that sometimes has message, header, etc. turned off

[DataTables](https://www.datatables.net) is an amazing tool to display your tables in a user friendly way, and the [Buttons](https://www.datatables.net/extensions/buttons/) extension makes downloading those tables a breeze. 

Now you can **easily style the Excel download** of your DataTable without having to learn the intricacies of Office Open XML using either:

* Custom Styles - Your own custom defined font, border, background, number format and alignment styles, or
* Pre-defined Templates - A selection of templates to apply to your entire table or selected cells

[Table of Contents](#table-of-contents)


## Demo

[View the live Excel style demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html) containing lots of examples of how to style your Excel sheet

## Installing

1. If you don't already have the 'Excel' download button running on your DataTable, make sure you add the Buttons Extension and JSZip to your page. [Download from DataTables.net](https://www.datatables.net/download/)

2. Include the javascript files for this plugin from the following cdn, or download from git or npm and add the scripts in the 'js/' folder to your page.

```html
<script src="https://cdn.jsdelivr.net/npm/datatables-buttons-excel-styles@1.2.0/js/buttons.html5.styles.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/datatables-buttons-excel-styles@1.2.0/js/buttons.html5.styles.templates.min.js"></script>
```


## Usage

This plugin adds a new option named `excelStyles` to the [DataTables Buttons configuration](https://datatables.net/reference/option/buttons.buttons).

This option will contain your style which consists of either a single [Excel Style Object](#excel-style-object) or an array of [Excel Style Objects](#excel-style-object) to be applied to your table.


### Style Example

With a custom [Style Object](#style-object) you can customize your spreadsheet to look exactly as you'd like it to. Either use familiar Excel cell references or take advantage of the [Cell Reference](#cell-reference) definitions available, to target specific parts of your worksheet. 

[See this example live](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/single_style.html)

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",                    // Extend the excel button
            excelStyles: {                      // Add an excelStyles definition
                cells: "2",                     // to row 2
                style: {                        // The style block
                    font: {                     // Style the font
                        name: "Arial",          // Font name
                        size: "14",             // Font size
                        color: "FFFFFF",        // Font Color
                        b: false,               // Remove bolding from header row
                    },
                    fill: {                     // Style the cell fill (background)
                        pattern: {              // Type of fill (pattern or gradient)
                            color: "457B9D",    // Fill color
                        }
                    }
                }
            },
        },
    ],
});
```


### Template Example

[Pre-defined templates](#pre-defined-templates) are a quick option for a nice output. 

[See this example live](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/single_template_style.html)

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",              // Extend the excel button
            excelStyles: {                // Add an excelStyles definition
                template: "blue_medium",  // Apply the 'blue_medium' template
            },
        },
    ],
});
```


### Styles and Templates Combined

You can easily combine the two. Start with a nice design and then make it yours! 

[See this example live](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/combine_template_and_style.html)

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",                    // Extend the excel button
            excelStyles: [                      // Add an excelStyles definition
                {                 
                    template: "green_medium",   // Apply the "green_medium" template
                },
                {
                    cells: "sh",                // Use Smart References (s) to target the header row (h)
                    style: {                    // The style definition
                        font: {                 // Style the font
                            size: 14,           // Size 14
                            b: false,           // Turn off the default bolding of the header row
                        },
                        fill: {                 // Style the cell fill
                            pattern: {          // Add a pattern (default is solid)
                                color: "1C3144" // Define the fill color
                            }
                        }
                    }
                }
            ]           
        },
    ],
});
```


### Built-in Styles

Built-in styles can also be used. See the [DataTables built-in style reference](https://datatables.net/reference/button/excelHtml5#Built-in-styles) for pre-defined style definitions.

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",    // Extend the excel button
            excelStyles: {      // Add an excelStyles definition
                cells: "sh",    // Use Smart References (s) to target the header row (h)
                index: 12,      // Apply the built-in style #12 which gives the cells a red background
            },
        },
    ],
});
```


### Conditional Formatting

You can apply [Conditional Styles](#conditional-styles) to cells. Not all Excel conditional formatting is currently supported, but the most common conditional number formatting is as well as formulas which cover most other use cases. Note that you can only apply custom styles (ie. not templates or built-in styles).

ColorScale, DataBar and IconSets are also supported.  See the [test site](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html) for examples.

The major benefit to using this method (as opposed to writing your own customize method to add fixed styles to cells) is that **the applied styles automatically update when you make changes to your data inside of Excel AFTER opening the sheet**.

[See this example live](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/conditional_style.html)

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",                        // Extend the excel button
            excelStyles: {                          // Add an excelStyles definition
                cells: "sF",                        // (s) Smart row reference, All data rows in column F
                condition: {                        // Add this style conditionally
                    type: 'cellIs',                 // Use the 'cellIs' condition type
                    operator: 'between',            // Use the 'between' operator
                    formula: [150000,200000],   // Add the two numbers to match between
                },
                style: {                            // The style block
                    font: {
                        bold: true,
                    },
                    fill: {
                        pattern: {
                            bgColor: "457B9D",      // NOTE: An excel quirk is that conditional solid fills need 
                                                    // the bgColor set, not the fgColor as for normal fills. 
                        }
                    }
                }
            },
        },
    ],
});
```

## Replace or insert cells, columns and rows

This isn't fully documented yet, but is very clear from the example below. Please see the demo site for further examples, but here are the basics. 

[Inserting rows demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/page_style.html)

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",                // Extend the excel button
            insertCells: [                  // Add an insertCells config option 
                {
                    cells: 'sCh',               // Target the header with smart selection
                    content: 'New column C',    // New content for the cells
                    pushCol: true,              // pushCol causes the column to be inserted
                },
                {
                    cells: 'sC1:C-0',           // Target the data
                    content: '',                // Add empty content
                    pushCol: true               // push the columns to the right over one
                },
                {
                    cells: 's5:6',              // Target data row 5 and 6
                    content: '',                // Add empty content
                    pushRow: true               // push the rows down to insert the content
                },
                {
                    cells: 'B3',                // Target cell B3
                    content: 'THIS IS CELL B3', // without pushCol or pushRow defined, the cell
                                                // is overwritten
                }
            ],
            excelStyles: {
                template: 'cyan_medium',    // Add a template to the result
            }
        },
    ],
});
```

## Printer defaults

Printer default settings can be defined using the pageStyle option. I haven't completed a full list of options yet, but it's pretty clear from the example below.

The 'repeatHeading' option can be set to true to repeat the heading at the top of each printed page, or can also be a [cell reference](./docs/cell_references.md) - note that only the row part of the reference is used (eg. `repeatHeading: 'st:h'` will repeat the title and heading row on each printed page.)  The 'repeatCol' option repeats columns when the table width spans multiple pages.

[Printer defaults demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/page_style.html)

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",                // Extend the excel button
            pageStyle: {
                sheetPr: {
                    pageSetUpPr: {
                        fitToPage: 1            // Fit the printing to the page
                    } 
                },
                printOptions: {
                    horizontalCentered: true,
                    verticalCentered: true,
                },
                pageSetup: {
                    orientation: "landscape",   // Orientation
                    paperSize: "9",             // Paper size (1 = Letter, 9 = A4)
                    fitToWidth: "1",            // Fit to page width
                    fitToHeight: "0",           // Fit to page height
                },
                pageMargins: {
                    left: "0.2",
                    right: "0.2",
                    top: "0.4",
                    bottom: "0.4",
                    header: "0",
                    footer: "0",
                },
                repeatHeading: true,    // Repeat the heading row at the top of each page
                repeatCol: 'A:A',       // Repeat column A (for pages wider than a single printed page)
            },
            excelStyles: {
                template: 'blue_gray_medium',    // Add a template style as well if you like
            }
        },
    ],
});
```

[pageSetup options can be found here](https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_pageSetup_topic_ID0EHZ54.html)

## Applying your Styles

In most cases your styles will be automatically applied by this plugin, but please consider the following.

**Please Note:** This plugin hooks in using the [customize option](https://datatables.net/reference/button/excel) in the DataTables Buttons configuration to automatically run when you click on the Excel button. If you are already using the `customize` method to apply other table formatting or modification, you can run this plugin by calling `applyStyles` from within your `customize` method as follows:

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                // ... custom Excel Style Objects defined ...
            },
            customize: function(xlsx) {
                // ... your custom code here ...

                // Apply the excelStyles
                this.applyStyles(xlsx); 
            }
        }
    ]
});
```


## Excel Style Object

The `excelStyles` DataTables Buttons option is added as a configuration item for the DataTables Buttons object. It contains either a single Excel Style Object or an array of Excel Style Objects.

| Attribute | Description | Type | Default |
|---|---|---|---|
| cells | The cell or cell range that the style is being applied to. | String or Array of<br />([Cell References](#cell-reference)) |
| rowref    | Enables smart row references if set to "smart" | Enum<br />( false \| "smart" ) | false |
| style     | The style definition          | [Style Object](#style-object) |
| template  | A template name               | String |
| index     | Built-in style index number   | Integer |
| merge     | Merge this style with the existing cell style | Boolean | true |
| width     | Set the column width          | Double |
| height    | Set the row height            | Double |
| condition | Condition to match for conditional formatting | [Condition Object](#condition-object) |


## Cell Reference

Use familiar Excel cell references to select a specific cell or cell range.

The Cell Reference can be a single string, or an array of references if you wish to apply the style to a 
range of cells in different locations (eg. applying the same style to the header and the footer)

[View this page for a complete list of all cell reference options](./docs/cell_references.md)

**Standard references**
* `A2` - Select cell A2
* `C17` - Select cell C17
* `B3:D20` - Select the range from cell B3 to cell D20

**Extended references** are used to select individual rows and columns, or row/column ranges:
* `4` - All cells in row 4
* `B` - All cells in column B
* `3:7` - All cells from (and including) row 3 to row 7
* `3:` - All cells from row 3 to the end of the table
* `>` - The last column in the table
* `-0` - The last row in the table
* `-2` - The third to last row in the table
* [and more...](./docs/cell_references.md)

**Smart row references** can select the various parts of the table (title, header, data, footer, etc.). These are enabled with a `s` prefix in the cell reference, or with the `rowref: "smart"` config option:
* `sh` - The header
* `sf` - The footer
* `s1` - Becomes the first data row
* `s-0` - Becomes the last data row
* `sB3` - Column B, row 3 of the data rows
* [and more...](./docs/cell_references.md)

For examples of using these cell selections please [view the demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html), 
or have a look at the templates in [buttons.html5.styles.templates.js](https://github.com/pjjonesnz/datatables-buttons-excel-styles/blob/master/js/buttons.html5.styles.templates.js)


## Style Object

There are five main properties available within a Style Object.

| Attribute | Description | Type |
|---|---|---|
| font      | To style the font used in a cell | [Font Object](#font-object) |
| border    | The border of the cell | [Border Object](#border-object) |
| fill      | To style the cell fill (ie. the cell background color and pattern) | [Fill Object](#fill-object) |
| numFmt    | Apply a number format (eg. define currency display, decimal places, etc.) | [NumFmt String](#numfmt-string) |
| alignment | Horizontal and vertical alignment of the cell content | [Alignment Object](#alignment-object) |


## Font Object

The font style is the simplest and consists of an object with the font attributes listed as key:value pairs inside.

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                cells: "A2:",
                style: {

                    // Font Object
                    font: {
                        name: "Arial",
                        size: 18,
                        u: true,          // Single underline
                        color: "D75F41"
                    }
                }
            }
        }
    ]
});
```
### Font Attributes 

The commonly used font attributes are listed below. A full list can be found in the [Office Open XML Spec](https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_font_topic_ID0EAXC6.html)

| Attribute | Meaning | Type | Example | Aliases |
|---|---|---|---|---|
| b         | Bold          | Boolean               | `bold: true`              | strong<br />bold  |
| color     | Color         | String (RGB or ARGB) or<br />[Color Object](#color-object)    | `color: "FF0000"`<br />`color: { rgb: "FF0000", tint: 0.54 }`   |
| family    | Font family   | Integer               | `family: 1`               | 
| i         | Italic        | Boolean               | `i: true`                 | italic |
| name      | Font name     | String                | `name: "Arial"`           |                   
| strike    | Strike through | Boolean              | `strike: true`            |
| sz        | Font size (pt) | Double               | `sz: 14`                  | size |
| u         | Underline      | Boolean or<br />String  | `u: true` (single underline)<br />`u: "singleAccounting"`<br />`u: "double"`<br />`u: "doubleAccounting"` | underline |
| vertAlign | Subscript<br />Superscript | String              | `vertAlign: "subscript"`<br />`vertAlign: "superscript"`  | |


#### Color Object

| Attribute | Meaning | Type | Example | Default |
|---|---|---|---|---|
| rgb   | Hex RGB or ARGB color value           | String | `rgb: "0C96FD"`<br />`rgb: "800C96FD"` |
| tint  | The tint value applied to the color   | Double (-1.0 to 1.0)  | `tint: -0.3` | 0.0 |


## Border Object

The border of a cell can be defined by a simple object

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                cells: "A2:",
                style: {

                    // Border Object
                    border: {
                        top: "thin",            // Thin black border at top of cell/s
                        bottom: {               // At the bottom of the cell/s apply a
                            style: "thick",     // thick border with
                            color: "A9D08E",    // a lovely hue of green
                        },
                    }
                }
            }
        }
    ]
});
```


### Border Attributes

| Attribute | Meaning | Type | Example | 
|---|---|---|---|
| top<br />bottom<br />left<br />right<br />diagonal | Border position | String (Border Style)<br />[Border Style Object](#border-style-object) | `top: "thin"`<br />`bottom: { style: "dashed", color: "A9D08E" }` |


#### Border Style Object

| Attribute | Meaning | Type | Example | 
|---|---|---|---|
| style | The style of the border   | Enum ([Border Styles String](#border-styles-string))      | `style: "medium"` |
| color | The border color          | String or<br />[Color Object](#color-object)  | `color: "FF0000"`<br />`color: { rgb: "FF0000", tint: 0.54 }` |


#### Border Styles String

| Value | Meaning | 
|---|---|
| dashDot           | Dash Dot Pattern                      |
| dashDotDot        | Dash Dot Dot Pattern                  |
| dashed            | Dashed Pattern                        |
| dotted            | Dotted Pattern                        |
| double            | Double Line Border                    |
| hair              | Hairline Border                       |
| medium            | Medium Weight Border                  |
| mediumDashDot     | Medium Weight Dash Dot Pattern        |
| mediumDashDotDot  | Medium Weight Dash Dot Dot Pattern    |
| mediumDashed      | Medium Weight Dashed Pattern          |
| slantDashDot      | Slant Dash Dot Pattern                |
| thick             | Thick Weight Border                   |
| thin              | Thin Weight Border                    |


## Fill Object

The fill style can either be a pattern or a gradient. While these styles are fully supported by Excel on all devices, many of the advanced pattern and gradient options are not completely supported by other spreadsheet viewers (eg. the default ios viewer)


**Solid background color**

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                cells: "A2:",
                style: {

                    // Solid Fill
                    fill: {
                        pattern: {
                            color: "457B9D",
                        }
                    }
                }
            }
        }
    ]
});
```


**Patterned background**

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                cells: "A2:",
                style: {

                    // Patterned Fill
                    fill: {
                        pattern: {
                            type: "lightUp",
                            fgColor: "1C3144",
                            bgColor: "C3D898",
                        }
                    }
                }
            }
        }
    ]
});
```


**Gradient background**

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                cells: "A2:",
                style: {

                    // Gradient Fill
                    fill: {
                        gradient: {
                            degree: 90,
                            stop: [
                                {
                                    position: 0,
                                    color: "000000",
                                },
                                {
                                    position: 1,
                                    color: "CC0000",
                                }
                            ]
                        }
                    }
                }
            }
        }
    ]
});
```


### Fill Attributes

| Attribute | Meaning | Type | Aliases |
|---|---|---|---|
| pattern   | Pattern Fill  | [Pattern Object](#pattern-object)   | patternFill |
| gradient  | Gradient Fill | [Gradient Object](#gradient-object)    | gradientFill |


#### Pattern Object

| Attribute | Meaning | Type | Example | Aliases |
|---|---|---|---|---|
| type      | Type of pattern       | [Pattern Type String](#pattern-type-string) | `type: "lightUp"`<br />Default: `"solid"` | |
| fgColor   | Foreground color      | String or<br />[Color Object](#color-object)    | `fgColor: "FF0000"`<br />`fgColor: { rgb: "FF0000", tint: 0.54 }`   | color |
| bgColor   | Background color      | String or<br />[Color Object](#color-object)    | `bgColor: "FF0000"`<br />`bgColor: { rgb: "FF0000", tint: 0.54 }`   | |


#### Pattern Type String

This list along with examples can be [found here](https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_PatternType_topic_ID0EBYQFB.html). Note that support for pattern types is good with Excel but limited with some other viewers.

| Value | Meaning | 
|---|---|
| darkDown | Dark Down |
| darkGray | Dark Gray |
| darkGrid | Dark Grid |
| darkHorizontal | Dark Horizontal |
| darkTrellis | Dark Trellis |
| darkUp | Dark Up |
| darkVertical | Dark Vertical |
| gray0625 | Gray 0.0625 |
| gray125 | Gray 0.125 |
| lightDown | Light Down |
| lightGray | Light Gray |
| lightGrid | Light Grid |
| lightHorizontal | Light Horizontal |
| lightTrellis | Light Trellis |
| lightUp | Light Up |
| lightVertical | Light Vertical |
| mediumGray | Medium Gray |
| solid | Solid |


#### Gradient Object

| Attribute | Meaning | Type | Example | 
|---|---|---|---|
| type   | Gradient fill type           | Enum<br />( `linear` \| `path` )    | `type: "linear"`<br />`type: "path"` |
| degree | Angle of the gradient<br />for linear gradients | Integer | `degree: "270"` |
| left<br />right<br />top<br />bottom | Edge position percentage of the inner rectangle<br />for path gradients | Double<br />(0.0 - 1.0) | `left: "0.3"` |
| stop   | Array of two or more gradient stops  | Array of [Stop Objects](#stop-object) | `stop: [{ position: "0", color: "#FF0000"}, ..., ...]` |


#### Stop Object

| Attribute | Meaning | Type | Example | 
|---|---|---|---|
| position  | Position percentage | Double<br />(0.0 to 1.0)    | `position: "0"`<br />`position: "1"` |
| color     | Color               | String or<br />[Color Object](#color-object)          | `fgColor: "FF0000"`<br />`fgColor: { rgb: "FF0000", tint: 0.54 }`   |


## NumFmt String

The numFmt attribute is used to apply advanced formatting to cells containing numbers. It consists of a single string with the number formatting code.

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                cells: "A2:",
                style: {

                    // NumFmt String
                    numFmt: "#,##0.0000;(#,##0.0000)"
                }
            }
        }
    ]
});
```

See Microsoft's guide for [Number format codes](https://support.office.com/en-us/article/Number-format-codes-5026BBD6-04BC-48CD-BF33-80F18B4EAE68)

The easiest way to find a custom code using Excel is as follows:

1. Open Excel
2. Modify a cell to format the number in the way you would like it
3. View the 'Format cells...' dialog 
4. Select the 'Number' tab
5. Click 'Custom' in the category list
6. Copy the code from the 'Type' input and use that as your NumFmt String


## Alignment Object

The alignment object applies alignment to the content of your cells.

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                cells: "A2:",
                style: {

                    // Alignment Object
                    alignment: {
                        vertical: "center",
                        horizontal: "left",
                        wrapText: true,
                    }

                }
            }
        }
    ]
});
```


### Alignment Attributes

| Attribute | Meaning | Type | Info |
|---|---|---|---|
| horizontal    | Horizontal Alignment  | [Horizontal Alignment Enum](#horizontal-alignment-enum) |
| indent        | Indent                | Integer                   | value is multiplied by 3 text spaces |
| readingOrder  | Reading Order         | Integer                   | 0 - Context Dependent<br />1 - Left-to-Right<br />2 - Right-to-Left |
| shrinkToFit   | Shrink To Fit         | Boolean                   | Should text be shrunk to fit cell width |
| textRotation  | Text Rotation         | Unsigned Integer (0 - 180)  | Degrees to rotate text
| vertical      | Vertical Alignment    | [Vertical Alignment Enum](#vertical-alignment-enum) |
| wrapText      | Word Wrapping         | Boolean |


#### Horizontal Alignment Enum

| Value | Meaning |
|---|---|
| center                | Centered Horizontal Alignment |
| centerContinuous      | Center Continuous Horizontal Alignment |
| distributed           | Distributed Horizontal Alignment |
| fill                  | Fill |
| general               | General Horizontal Alignment |
| justify               | Justify |
| left                  | Left Horizontal Alignment |
| right                 | Right Horizontal Alignment |


#### Vertical Alignment Enum

| Value | Meaning |
|---|---|
| bottom            | Aligned To Bottom |
| center            | Centered Vertical Alignment |
| distributed       | Distributed Vertical Alignment |
| justify           | Justified Vertically |
| top               | Align Top |


## Pre-defined templates

Predefined templates are located in the `buttons.html5.styles.templates.min.js` javascript file. This file must be included on your page to use the templates.

Templates are as simple to apply as this:

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {                // Add an excelStyles definition
                template: "blue_medium",  // Apply the 'blue_medium' template
            }
        }
    ]
});
```

Complete templates are made up of **template parts** that can also be individually applied:

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {                // Add an excelStyles definition
                template: "header_blue",  // Apply the 'header_blue' template part (white font on a blue background in the header/footer)
            }
        }
    ]
});
```

Multiple templates can be applied by using an array. If a second template target the same cell style as the preceding one, it will overwrite that style.

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {              // Add an excelStyles definition
                template: [             // Apply multiple templates
                    "gold_medium",      // Apply the 'gold_medium' template to the entire table
                    "header_cyan"       // Overwrite the header with the 'header_cyan' style
                ],    
            }
        }
    ]
});
}
```

Some of the template parts are suitable to apply to a specific cell or cell range. 
If you don't define `cells`, the template will be applied to the default range of cells defined within the template itself. If no default range or cell reference is defined, then the template will be skipped.

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {                  // Add an excelStyles definition
                cells: "sD3",               // Column D, row 3 of the data rows
                template: "currency_eu",    // Format the cells with a predefined numFmt 
            }                               // displaying the values as Euro currency
        }
    ]
});
```

For multiple templates targeting different cell ranges, use an array of Style Objects with the cells and template defined in each object

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: [
                {                  
                    cells: "sD",
                    template: "currency_eu",
                },
                {                  
                    cells: "sE",
                    template: "date_long",
                }
            ]
        }
    ]
});
```


## Template List


### Full Table Templates

| Template Name | Description                   |
|---|---|
| black_medium | Complete table: black theme |
| blue_gray_medium | Complete table: blue-gray theme |
| blue_medium | Complete table: blue theme |
| cyan_medium | Complete table: cyan theme |
| gold_medium | Complete table: gold theme |
| gray_medium | Complete table: gray theme |
| green_medium | Complete table: green theme |
| light_gray_medium | Complete table: light-gray theme |
| orange_medium | Complete table: orange theme |

Each of the Full Table Templates above automatically applies the Header & Footer, Row Stripes, Row Borders & Table Outline templates.


### Template Parts

| Part Name     | Description       | Default Cell Reference |
|---|---|---|
| **Basic Types** |
| b             | Bold              | `s1:-0` All data rows |
| u             | Underline         | `s1:-0` All data rows |
| i             | Italic            | `s1:-0` All data rows |
| |
| **Headers and Footers** |
| header_black | Header: black | `['sh', 'sf']` Header and footer |
| header_blue | Header: blue | `['sh', 'sf']` Header and footer |
| header_blue_gray | Header: blue_gray | `['sh', 'sf']` Header and footer |
| header_cyan | Header: cyan | `['sh', 'sf']` Header and footer |
| header_gold | Header: gold | `['sh', 'sf']` Header and footer |
| header_gray | Header: gray | `['sh', 'sf']` Header and footer |
| header_green | Header: green | `['sh', 'sf']` Header and footer |
| header_light_gray | Header: light_gray | `['sh', 'sf']` Header and footer |
| header_orange | Header: orange | `['sh', 'sf']` Header and footer |
| |
| **Row Stripes** |
| stripes_black | Stripes: black | `s1:n,2` All columns, every second row of the data rows |
| stripes_blue | Stripes: blue | `s1:n,2` All columns, every second row of the data rows |
| stripes_blue_gray | Stripes: blue_gray | `s1:n,2` All columns, every second row of the data rows |
| stripes_cyan | Stripes: cyan | `s1:n,2` All columns, every second row of the data rows |
| stripes_gold | Stripes: gold | `s1:n,2` All columns, every second row of the data rows |
| stripes_gray | Stripes: gray | `s1:n,2` All columns, every second row of the data rows |
| stripes_green | Stripes: green | `s1:n,2` All columns, every second row of the data rows |
| stripes_light_gray | Stripes: light_gray | `s1:n,2` All columns, every second row of the data rows |
| stripes_orange | Stripes: orange | `s1:n,2` All columns, every second row of the data rows |
| |
| **Row Borders Top & Bottom** |
| rowlines_black | Rowlines: black | `sh:f` All columns, every row from the header to the footer |
| rowlines_blue | Rowlines: blue | `sh:f` All columns, every row from the header to the footer |
| rowlines_blue_gray | Rowlines: blue_gray | `sh:f` All columns, every row from the header to the footer |
| rowlines_cyan | Rowlines: cyan | `sh:f` All columns, every row from the header to the footer |
| rowlines_gold | Rowlines: gold | `sh:f` All columns, every row from the header to the footer |
| rowlines_gray | Rowlines: gray | `sh:f` All columns, every row from the header to the footer |
| rowlines_green | Rowlines: green | `sh:f` All columns, every row from the header to the footer |
| rowlines_light_gray | Rowlines: light_gray | `sh:f` All columns, every row from the header to the footer |
| rowlines_orange | Rowlines: orange | `sh:f` All columns, every row from the header to the footer |
| |
| **Table Outline** |
| outline_black | Outline: black |
| outline_blue | Outline: blue |
| outline_blue_gray | Outline: blue_gray |
| outline_cyan | Outline: cyan |
| outline_gold | Outline: gold |
| outline_gray | Outline: gray |
| outline_green | Outline: green |
| outline_light_gray | Outline: light_gray |
| outline_orange | Outline: orange |
| |
| **Currency Format** |
| currency_us | US currency number format |
| currency_eu | Euro currency number format |
| currency_gb | GB Pound currency number format |
| |
| **Date Format** |
| date_long | Date: Long format - eg. 24 September 1979 |
| date_medium | Date: Medium format - eg. 4 Dec 1987 |
| |
| **Number Format** |
| int         | Integer number format |
| decimal_1   | Number format - one decimal place, negatives in brackets |
| decimal_2   | Number format - two decimal places, negatives in brackets |
| decimal_3   | Number format - three decimal places, negatives in brackets |
| decimal_4   | Number format - four decimal places, negatives in brackets |


## Conditional Styles

Conditional formatting causes styles to only be applied to cells if the condition is met.

### Condition Object

| Attribute | Description | Type | Default |
|---|---|---|---|
| type      | The type of conditional formatting rule. | [Type Enum](#type-enum) |
| operator  | When type is 'cellIs', sets the comparison type. Don't use with expression. | [Operator Enum](#operator-enum) |
| formula   | The data to compare | Number/String<br>Array (for between or notBetween operators) |
| priority  | The priority of this conditional formatting rule | Number | 1 |

NOTE: As this feature is still not complete, not all [conditional types](https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_CfType_topic_ID0EFYEFB.html) have been tested. 

Presently `cellIs` and `expression` have been tested with examples on the [test site](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html).  The 'expression' type should cover most use cases.  ColorScale, dataBar and iconSet formatting also work with examples on the test site. 

Also

Example:

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                cells: "sF",                    // Smart select Column F

                // Conditional formatting block
                condition: {
                    type: "cellIs",
                    operator: "greaterThan",
                    formula: 150000,
                },
                style: {
                    fill: {
                        pattern: {
                            bgColor: "F78989"   // Brighter red (Note the Excel gotcha requires
                        }                       // bgColor for conditional formatting)
                    }
                }
            }
        }
    ]
});
```

Using multiple values

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                cells: "sF",                    // Smart select Column F
                condition: {
                    type: "cellIs",
                    operator: "notBetween",
                    formula: [150000,200000],
                },
                style: {
                    fill: {
                        pattern: {
                            bgColor: "F78989"   // Brighter red (Note the Excel gotcha requires
                        }                       // bgColor for conditional formatting)
                    }
                }
            }
        }
    ]
});
```

Using a formula

```js
$("#myTable").DataTable({
    dom: "Bfrtip",
    buttons: [
        {
            extend: "excel",
            excelStyles: {
                cells: "s:",                    // Smart select all data rows - ie. apply to entire row
                condition: {
                    type: "expression",         // Use a forumula
                    formula: "AND($F3>150000,$B3=\"Software Engineer\")",   // The formula - make sure you escape strings and use WITHOUT leading = (equals) sign.  You also can't use smart row references in the formula.  Note the $ in front of the cell reference to lock it to the column.
                },
                style: {
                    fill: {
                        pattern: {
                            bgColor: "F78989"   // Brighter red (Note the Excel gotcha requires
                        }                       // bgColor for conditional formatting)
                    }
                }
            }
        }
    ]
});
```

#### Type Enum

| Value | Meaning |
|---|---|
| cellIs               | Compares a cell value using an operator  |
| expression           | Evaluate a formula |
| colorScale           | Creates a gradated color scale on the cells |
| dataBar              | Displays a gradated data bar in the range of cells |
| iconSet              | Applies icons to cells according to their values |

Note: See the demo site for examples as colorScale, dataBar and iconSet need specific attributes set.

#### Operator Enum

| Value | Meaning |
|---|---|
| between               | Value is between the two numbers in the formula array |
| equal                 | Value is equal to the number in the formula |
| greaterThan           | Value is greater than the number in the formula |
| greaterThanOrEqual    | Value is greater than or equal to the number in the formula |
| lessThan              | Value is less than the number in the formula |
| lessThanOrEqual       | Value is less than or equal to the number in the formula |
| notBetween            | Value is NOT between the two numbers in the formula array |
| notEqual              | Value is NOT equal to the number in the formula |


## License

This plugin is released under the MIT license. You are free to use, modify and distribute this software, as long as the copyright header is left intact.


## Contributing

I hope this plugin helps you output beautiful spreadsheets from your DataTables.net enabled tables.

Please let me know if you have any helpful comments or code that you would like to contribute or if this plugin has been helpful.


## Thanks

To the DataTables.net team for making jQuery table display so much easier and more beautiful!

To the C-Rex team for their excellent [Office Open XML File Format Reference](https://c-rex.net/projects/samples/ooxml/e1/index.html)


## Table of Contents

- [DataTables Buttons Excel Styling](#datatables-buttons-excel-styling)
  * [Demo](#demo)
  * [Installing](#installing)
  * [Usage](#usage)
    + [Style Example](#style-example)
    + [Template Example](#template-example)
    + [Styles and Templates Combined](#styles-and-templates-combined)
    + [Built-in Styles](#built-in-styles)
    + [Conditional Formatting](#conditional-formatting)
  * [Replace or insert cells, columns and rows](#replace-or-insert-cells--columns-and-rows)
  * [Printer defaults](#printer-defaults)
  * [Applying your Styles](#applying-your-styles)
  * [Excel Style Object](#excel-style-object)
  * [Cell Reference](#cell-reference)
  * [Style Object](#style-object)
  * [Font Object](#font-object)
    + [Font Attributes](#font-attributes)
      - [Color Object](#color-object)
  * [Border Object](#border-object)
    + [Border Attributes](#border-attributes)
      - [Border Style Object](#border-style-object)
      - [Border Styles String](#border-styles-string)
  * [Fill Object](#fill-object)
    + [Fill Attributes](#fill-attributes)
      - [Pattern Object](#pattern-object)
      - [Pattern Type String](#pattern-type-string)
      - [Gradient Object](#gradient-object)
      - [Stop Object](#stop-object)
  * [NumFmt String](#numfmt-string)
  * [Alignment Object](#alignment-object)
    + [Alignment Attributes](#alignment-attributes)
      - [Horizontal Alignment Enum](#horizontal-alignment-enum)
      - [Vertical Alignment Enum](#vertical-alignment-enum)
  * [Pre-defined templates](#pre-defined-templates)
  * [Template List](#template-list)
    + [Full Table Templates](#full-table-templates)
    + [Template Parts](#template-parts)
  * [Conditional Styles](#conditional-styles)
    + [Condition Object](#condition-object)
      - [Operator Enum](#operator-enum)
  * [License](#license)
  * [Contributing](#contributing)
  * [Thanks](#thanks)

<small><i><a href='http://ecotrust-canada.github.io/markdown-toc/'>Table of contents generated with markdown-toc</a></i></small>