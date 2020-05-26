# DataTables Buttons Excel Styling

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/pjjonesnz/datatables-buttons-excel-styles)](https://github.com/pjjonesnz/datatables-buttons-excel-styles/releases)
[![GitHub license](https://img.shields.io/github/license/pjjonesnz/datatables-buttons-excel-styles)](https://github.com/pjjonesnz/datatables-buttons-excel-styles/blob/master/LICENSE.md)
[![npm](https://img.shields.io/npm/v/datatables-buttons-excel-styles)](https://www.npmjs.com/package/datatables-buttons-excel-styles)

**Add beautifully styled Excel output to your DataTables.**

[DataTables](https://www.datatables.net) is an amazing tool to display your tables in a user friendly way, and the [Buttons](https://www.datatables.net/extensions/buttons/) extension makes downloading those tables a breeze. 

Now you can simply style your downloaded tables without having to learn the intricacies of SpreadsheetML using either:

* Styles: Your own custom defined font, border, background and number format style, and/or
* Pre-defined Templates: A selection of templates to apply to your table or selected cells

## Demo

[View the Excel style demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html)

## Installing

1. If you don't already have DataTables set up to download excel spreadsheets, add jQuery, DataTables, Buttons Extension and JSZip to your page. [Download from DataTables.net](https://www.datatables.net/download/)

2. Include the javascript files for this plugin from the following cdn, or download from this repository and add the scripts in the js/ folder to your page.

```html
<script src="https://cdn.jsdelivr.net/npm/datatables-buttons-excel-styles@0.7.5/js/buttons.html5.styles.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/datatables-buttons-excel-styles@0.7.5/js/buttons.html5.styles.templates.min.js"></script>
```

## Usage

Add an excelStyles config option to apply a style or template to the Excel file. 

### Style Example

With a custom style you can make the table look exactly as you want. With clever cell definitions available to target specific parts of the worksheet. [See this example live](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/single_style.html)

```js
$(document).ready(function () {
    $('#myTable').DataTable({
        dom: 'Bfrtip',
        buttons: [
            {
                extend: 'excel',                    // Extend the excel button
                excelStyles: {                      // Add an excelStyles definition
                    cells: '2',                     // to row 2
                    style: {                        // The style block
                        font: {                     // Style the font
                            name: 'Arial',          // Font name
                            size: '14',             // Font size
                            color: 'FFFFFF',        // Font Color
                            b: false,               // Remove bolding from header row
                        },
                        fill: {                     // Style the cell fill (background)
                            pattern: {              // Type of fill (pattern or gradient)
                                color: '457B9D',    // Fill color
                            }
                        }
                    }
                },
            },
        ],
    });
});
```

### Template Example

Pre-defined templates are a quick option for a nice output. [See this example live](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/single_template_style.html)

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

### Styles and Templates Combined

You can easily combine the two. Start with a nice design and then make it yours! [See this example live](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/combine_template_and_style.html)

```js
$('#myTable').DataTable({
    dom: 'Bfrtip',
    buttons: [
        {
            extend: 'excel',                    // Extend the excel button
            excelStyles: [                      // Add an excelStyles definition
                {                 
                    template: 'green_medium',   // Apply the 'green_medium' template
                },
                {
                    cells: 'sh',                // Use Smart References (s) to target the header row (h)
                    style: {                    // The style definition
                        font: {                 // Style the font
                            size: 14,           // Size 14
                            b: false,           // Turn off the default bolding of the header row
                        },
                        fill: {                 // Style the cell fill
                            pattern: {          // Add a pattern (default is solid)
                                color: '1C3144' // Define the fill color
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

Built-in styles can also be used. See the [built in style reference](https://datatables.net/reference/button/excelHtml5#Built-in-styles)

```js
$('#myTable').DataTable({
    dom: 'Bfrtip',
    buttons: [
        {
            extend: 'excel',    // Extend the excel button
            excelStyles: {      // Add an excelStyles definition
                cells: 'sh',    // Use Smart References (s) to target the header row (h)
                index: 12,      // Apply the built-in style to make the heading bold with a red background
            },
        },
    ],
});
```

## excelStyles Attribute

The excelStyles attribute contains either a single Excel Style Objects or an array of Excel Style Objects

### Excel Style Object

| Attribute | Description | Type | Default |
|---|---|---|---|
| cells | The cell or cell range that the style is being applied to. | String<br />(Cell Reference) |
| rowref | Enables smart row references if set to "smart" | false \| "smart" | false |
| style | The style definition | Style Object |
| template | A template name | String |
| index | Built-in style index number | Integer |
| merge | Merge this style with the existing cell style | Boolean | true |
| width | Set the column width | Double |
| height | Set the row height | Double |

## Cell Reference

Use familiar Excel cell references to select a specific cell or cell range.

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

For examples of using these cell selections, while the docs are being written, please [view the demo](https://www.pauljones.co.nz/github/buttons-html5-styles/examples/simple_table_style.html), or have a look at the source of [buttons.html5.styles.templates.js](https://github.com/pjjonesnz/datatables-buttons-excel-styles/blob/master/js/buttons.html5.styles.templates.js)

## Style Object

There are five main properties available within a Style Object.

| Attribute | Description | Type |
|---|---|---|
| font | To style the font used in a cell | Font Object |
| border | The border of the cell | Border Object |
| fill | To style the cell fill (ie. the cell background color and pattern) | Fill Object |
| numFmt | Apply a number format (eg. define currency display, decimal places, etc.) | NumFmt Object |
| alignment | Horizontal and vertical alignment of the cell content | Alignment Object |

### Font Object

The font style is the simplest and consists of an object with the font attributes listed as key:value pairs inside.

```js
{
    font: {
        name: "Arial",
        size: 18,
        u: true,          // Single underline
        color: "D75F41"
    }
}
```

#### Font attributes 

The commonly used font attributes are listed below. A full list can be found in the [Office Open XML Spec](https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_font_topic_ID0EAXC6.html)

| Attribute | Meaning | Type | Example | Aliases |
|---|---|---|---|---|
| b         | Bold          | Boolean               | `bold: true`              | strong<br />bold  |
| color     | Color         | String (RGB or ARGB)<br />Color Object    | `color: "FF0000"`<br />`color: { rgb: "FF0000", tint: 0.54 }`   |
| family    | Font family   | Integer               | `family: 1`               | 
| i         | Italic        | Boolean               | `i: true`                 | italic |
| name      | Font name     | String                | `name: "Arial"`           |                   
| strike    | Strike through | Boolean              | `strike: true`            |
| sz        | Font size (pt) | Double               | `sz: 14`                  | size |
| u         | Underline      | Boolean<br />String  | `u: true` (single underline)<br />`u: "singleAccounting"`<br />`u: "double"`<br />`u: "doubleAccounting"` | underline |
| vertAlign | Subscript<br />Superscript | String              | `vertAlign: "subscript"`<br />`vertAlign: "superscript"`  | |

**Color Object**
| Attribute | Meaning | Type | Example | Default |
|---|---|---|---|---|
| rgb   | Hex RGB or ARGB color value           | String | `rgb: "0C96FD"`<br />`rgb: "800C96FD"` |
| tint  | The tint value applied to the color   | Double (-1.0 to 1.0)  | `tint: -0.3` | 0.0 |

### Border Object

The border of a cell can be defined by a simple object

```js
border: {
    top: "thin",            // Thin black border at top of cell/s
    bottom: {
        style: "thick",
        color: "A9D08E",
    },
}
```

#### Border attributes

| Attribute | Meaning | Type | Example | 
|---|---|---|---|
| top<br />bottom<br />left<br />right<br />diagonal | Border position | String (Border Style)<br />Border Style Object | `top: "thin"`<br />`bottom: { style: "dashed", color: "A9D08E" }` |

**Border Style Object**
| Attribute | Meaning | Type | Example | 
|---|---|---|---|
| style | The style of the border   | Enum (Border Styles)      | `style: "medium"` |
| color | The border color          | String<br />Color Object  | `color: "FF0000"`<br />`color: { rgb: "FF0000", tint: 0.54 }` |

**Border Styles**
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


### Fill Object

The fill style can either be a pattern or a gradient. While these styles are fully supported by Excel on all devices, many of the advanced pattern and gradient options are not completely supported by other spreadsheet viewers (eg. the default ios viewer)

**Solid background color**
```js
{
    fill: {
        pattern: {
            color: "457B9D",
        }
    }
}
```

**Patterned background**
```js
{
    fill: {
        pattern: {
            type: "lightUp",
            fgColor: "1C3144",
            bgColor: "C3D898",
        }
    }
}
```

**Gradient background**
```js
{
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
```

#### Fill attributes

| Attribute | Meaning | Type | Aliases |
|---|---|---|---|---|
| gradient  | Gradient Fill | Gradient Object    | gradientFill |
| pattern   | Pattern Fill  | Pattern Object    | patternFill |

**Pattern Object**
| Attribute | Meaning | Type | Example | Aliases |
|---|---|---|---|---|
| type      | Type of pattern       | String | `type: "lightUp"`<br />Default: `"solid"` | |
| fgColor   | Foreground color      | String<br />Object    | `fgColor: "FF0000"`<br />`fgColor: { rgb: "FF0000", tint: 0.54 }`   | color |
| bgColor   | Background color      | String<br />Object    | `bgColor: "FF0000"`<br />`bgColor: { rgb: "FF0000", tint: 0.54 }`   | |

**Gradient Object**
| Attribute | Meaning | Type | Example | 
|---|---|---|---|
| type   | Gradient fill type           | Enum<br />( `linear` \| `path` )    | `type: "linear"`<br />`type: "path"` |
| degree | Angle of the gradient<br />for linear gradients | Integer | `degree: "270"` |
| left<br />right<br />top<br />bottom | Edge position percentage of the inner rectangle<br />for path gradients | Double<br />(0.0 - 1.0) | `left: "0.3"` |
| stop   | Array of two or more gradient stops  | Stop Object | `stop: [{ position: "0", color: "#FF0000"}, ..., ...]` |

**Stop Object**
| Attribute | Meaning | Type | Example | 
|---|---|---|---|
| position  | Position percentage | Double<br />(0.0 to 1.0)    | `position: "0"`<br />`position: "1"` |
| color     | Color               | String<br />Object          | `fgColor: "FF0000"`<br />`fgColor: { rgb: "FF0000", tint: 0.54 }`   |

## NumFmt Object ##

coming soon...