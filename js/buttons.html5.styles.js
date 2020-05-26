/**
 * Styling for DataTables Buttons Excel XLSX (Open Office XML) creation
 *
 * @version: 0.7.7
 * @description Add and process a custom 'excelStyles' option to easily customize the DataTables Excel Stylesheet output
 * @file buttons.html5.styles.js
 * @copyright © 2020 Beyond the Box Creative
 * @author Paul Jones <info@pauljones.co.nz>
 * @license MIT
 *
 * Include this file after including the javascript for the DataTables, Buttons, HTML5 and JSZip extensions
 *
 * Create the required styles using the custom 'excelStyles' option in the button's config
 * @see https://datatables.net/reference/button/excel
 *
 * @todo Documentation on 'excelStyles' options - 'Coming soon...'
 */

(function (factory) {
    if (typeof define === 'function' && define.amd) {
        // AMD
        define([
            'jquery',
            'datatables.net',
            'datatables.net-buttons',
            'datatables.net-buttons/js/buttons.html5.js',
        ], function ($) {
            return factory($, window, document);
        });
    } else if (typeof exports === 'object') {
        // CommonJS
        module.exports = function (root, $) {
            if (!root) {
                root = window;
            }

            if (!$ || !$.fn.dataTable) {
                $ = require('datatables.net')(root, $).$;
            }

            if (!$.fn.dataTable.Buttons) {
                require('datatables.net-buttons')(root, $);
            }

            if (!$.fn.dataTable.Buttons.excelHtml5) {
                require('datatables.net-buttons/js/buttons.html5.js')(root, $);
            }

            return factory($, root, root.document);
        };
    } else {
        // Browser
        factory(jQuery, window, document);
    }
})(function ($, window, document, undefined) {
    //(function ($) {
    ('use strict');

    var DataTable = $.fn.dataTable;

    /**
     * Automatically run the applyStyles function if customize isn't redefined
     *
     * @param {object} xlsx
     */
    DataTable.ext.buttons.excelHtml5.customize = function (xlsx) {
        this.applyStyles(xlsx);
    };

    /**
     * Allow applyStyles to be triggered from a custom customize function
     * If excelStyles is defined but customize isn't, then it will
     * automatically be run so you don't need to do this.
     *
     * @example
     * buttons: {
     *       excelStyles: {
     *          ... custom styles defined ...
     *       },
     *       customize: function(xlsx) {
     *           this.applyStyles(xlsx);
     *           ... custom code here ...
     *       }
     *    }
     * }
     */

    DataTable.ext.buttons.excelHtml5.applyStyles = function (xlsx) {
        this._applyExcelStyles(xlsx);
    };

    /**
     * Parse cell names into a row, col, from:to like structure
     *
     * @example
     * Cell reference examples
     *
     * Single Range References
     *  'A3'   = cell A3
     *  '4'    = row 4, all columns
     *  'D'    = column D, all rows
     *
     * Multiple Range References (separated by :)
     *  '4:6'       = rows 4 to 6, all columns
     *  'B:F'       = column B to F, all rows
     *  'D4:D20'    = column D from row 4 to 20
     *  '3:'        = from row 3 until the last row, all columns
     *  'A:'        = from column A until the last column, all rows
     *  'B3:'       = from column B until the last column, from row 3 until the last row
     *  ':B3'       = from column A until column B, from row 1 to row 3
     *  'B3:D'      = from column B until column D, from row 3 until the last row
     *
     * References to the last column
     *  '>'         = all rows, the last column
     *  '>3:>20'    = the last column, row 3 to 20
     *
     * References counting back from the last column
     *  '-3>'       = three columns back from the last column
     *  '-2>5'      = two columns back from the last column, row 5
     *
     * References counting back from the last row
     *  '-0'        = all columns, the last row
     *  'B-3:B-0'   = column B from the third to last row until the last row
     *
     * Reference for everything
     *  ':'         = all columns, all rows (also '1:', 'A:', '', 'A1:', ':-0', ':>', ':>-0')
     *
     * Column/Row skipping
     *
     * Used to apply styles to every nth column or row (eg. every 2nd row, every 3rd column)
     *
     * Format: n[0-9],[0-9]
     * n (stands for every nth column/row), then the column increment followed by the row increment
     *
     * Column/Row skipping examples
     *  'A3:D10n1,2'    = from Column A row 3, to Column D row 10, target every column, target every second row
     *  '3:n1,2'        = every column from row 3 until the last row, target every second row (use this for row striping)
     *  ':n1,2'         = every column, every second row (also ':n,2')
     *  ':n2,1'         = every second column, every row (also ':n2')
     *
     * Smart row references
     *
     * With the default settings row references refer to the actual Excel spreadsheet rows (ie. 1 = row 1, 12 = row 12). This works well, but
     * can be hard to work with if your spreadsheet has (or doesn't have) the extra title and/or message above the data. Also, if you
     * include a footer this can be hard to define a template that works for custom excel configurations.
     *
     * Smart row references adds specific code to refer to these special rows, and redefines row 1 to be the first row of the data
     *
     * You can enable smart references by adding the following to your style definition, or by prefixing your cell reference with a lower case s
     * excelStyles: [
     *      {
     *          rowref: "smart",
     *          cells: "...cell reference..."
     *          style: { ...style definition... }
     *      }
     * ]
     *
     * Once enabled, the following row references are available, along with the row option that they refer to:
     *
     *  't'     = title: the title row (usually this is excel row 1)
     *  'm'     = messageTop: the message row (if enabled)
     *  'h'     = header: the row with the cell titles
     *  '1:-0'  = the data rows (also '1:' or ':-0' or ':') - note that will now ONLY refer to the data
     *  'f'     = footer: row with the cell titles (same content as the header row but at the bottom of the table)
     *  'b'     = messageBottom: the message row at the bottom of the table. The row below the footer row (it it is enabled)
     *
     * @param {string} cells Cell names in an Excel-like structure
     * @param {object} sheet The worksheet to enable finding of the last column/row
     * @return {object} Parsed rows and columns, in number format (ie. columns referenced by number, not letter)
     */
    var _parseExcellyReference = function (cells, sheet, smartRowOption) {
        //var pattern = /^(-\d+(?=\>))*([A-Z]*|[>])*(-)*([0-9]*)(\:)*(-\d+(?=\>))*([A-Z]*|[>])*(-)*([0-9]*)(?:n([0-9]*)(?:,)*([0-9]*))*$/;
        var pattern = /^(s)*(?:-(\d*)(?=\>))*([A-Z]*|[>])*([tmhfb]{1})*(-(?=[0-9]+))*([0-9]*)(?:(\:)(?:-(\d*)(?=\>))*([A-Z]*|[>])*([tmhfb]{1})*(-(?=[0-9]+))*([0-9]*)(?:n([0-9]*)(?:,)*([0-9]*))*)*$/;
        var matches = pattern.exec(cells);
        if (matches === null) {
            return false;
        }

        var results = {
            smartRow: matches[1],
            fromColEndSubtractAmount: matches[2],
            fromCol: matches[3],
            fromSmartRow: matches[4],
            fromRowEndSubtract: matches[5],
            fromRow: matches[6],
            range: matches[7],
            toColEndSubtractAmount: matches[8],
            toCol: matches[9],
            toSmartRow: matches[10],
            toRowEndSubtract: matches[11],
            toRow: matches[12],
            nthCol: matches[13],
            nthRow: matches[14],
            pattern: cells,
        };

        var _smartRow = function (index) {
            return parseInt(index) + _rowRefs.dt - 1;
        };

        /**
         * Modify the parsed cell results to account for smart row references
         *
         * @param {object} results The parsed cells
         * @param {boolean} smartRowOption Has the smart row option been set in excelStyles
         * @returns {boolean} True if a positive match has been made and resolved, or if this is not a smart row. False otherwise
         */
        function _patternMatchSmartRow(results, smartRowOption) {
            if (
                !smartRowOption &&
                (!results.smartRow || results.smartRow != 's')
            ) {
                results.smartRow = false;
                return true;
            }
            results.smartRow = true;

            if (results.fromRow && !results.fromRowEndSubtract) {
                results.fromRow = _smartRow(results.fromRow);
            }

            if (results.toRow && !results.toRowEndSubtract) {
                results.toRow = _smartRow(results.toRow);
            }

            var pattern = /['tmhfb']{1}/;
            if (results.fromSmartRow !== undefined) {
                var match = pattern.exec(results.fromSmartRow);
                if (match && _rowRefs[match[0]] !== false) {
                    results.fromRow = _rowRefs[match[0]];
                } else {
                    return false;
                }
            }
            if (results.toSmartRow !== undefined) {
                var match = pattern.exec(results.toSmartRow);
                if (match && _rowRefs[match[0]] !== false) {
                    results.toRow = _rowRefs[match[0]];
                } else {
                    return false;
                }
            }
            return true;
        }

        if (!_patternMatchSmartRow(results, smartRowOption)) {
            return false;
        }

        // Refine column results

        results.toCol =
            (results.toCol // if a to column has been specified
                ? !results.toColEndSubtractAmount // if we are NOT subtracting from the last column
                    ? results.toCol // return the selected column
                    : _getMaxColumnIndex(sheet) - results.toColEndSubtractAmount // else return last column minus this column number
                : null) || // else return null and continue
            (results.range || !results.fromCol // if there is a range selected, but no fromCol
                ? _getMaxColumnIndex(sheet) // return the maximum column
                : !results.fromColEndSubtractAmount // else if we are NOT subtracting from the last column for the from source
                ? results.fromCol // return the from column
                : _getMaxColumnIndex(sheet) - results.fromColEndSubtractAmount); // else return the last column minus the from column number

        results.toCol = _parseColumnName(results.toCol, sheet);
        results.fromCol = results.fromCol
            ? !results.fromColEndSubtractAmount
                ? results.fromCol
                : _getMaxColumnIndex(sheet) - results.fromColEndSubtractAmount
            : 1;
        results.fromCol = _parseColumnName(results.fromCol, sheet);
        results.nthCol = results.nthCol ? parseInt(results.nthCol) : 1;

        // Reverse the column results if from is higher than to

        if (results.fromCol > results.toCol) {
            var tempCol = results.fromCol;
            results.fromCol = results.toCol;
            results.toCol = tempCol;
        }

        // Refine row results
        results.toRow =
            (results.toRow // if a to row has been specified
                ? !results.toRowEndSubtract // if we are NOT subtracting from the last row
                    ? results.toRow // return the selected row
                    : _getMaxRow(sheet, results) - results.toRow // else return last row minus this row number
                : null) || // else return null and continue
            (results.range || !results.fromRow // if there is a range selected, but no fromRow
                ? _getMaxRow(sheet, results) // return the maximum row
                : !results.fromRowEndSubtract // else if we are NOT subtracting from the last row for the from source
                ? results.fromRow // return the from row
                : _getMaxRow(sheet, results) - results.fromRow); // else return the last row minus the from row number

        results.toRow = parseInt(results.toRow);

        results.fromRow = results.fromRow
            ? parseInt(
                  !results.fromRowEndSubtract
                      ? results.fromRow
                      : _getMaxRow(sheet, results) - results.fromRow
              )
            : _getMinRow(results);
        results.nthRow = results.nthRow ? parseInt(results.nthRow) : 1;

        // Reverse the row results if from is higher than to

        if (results.fromRow > results.toRow) {
            var tempRow = results.fromRow;
            results.fromRow = results.toRow;
            results.toRow = tempRow;
        }

        return results;
    };

    /**
     * Get the maximum row index - adjusts for smart row references
     *
     * @param {object} sheet Worksheet
     * @param {object} results Cell parsing results to check for smart row refs
     * @return {int} The maximum row number
     */
    var _getMaxRow = function (sheet, results) {
        if (results.smartRow) {
            return _rowRefs.db;
        }
        return _getMaxSheetRow(sheet);
    };

    /**
     * Get the minimum row index - adjusts for smart row references
     *
     * @param {object} results Cell parsing results to check for smart row refs
     */
    var _getMinRow = function (results) {
        if (results.smartRow) {
            return _rowRefs.dt;
        }
        return 1;
    };

    var _getMaxSheetRow = function (sheet) {
        return Number($('sheetData row', sheet).last().attr('r'));
    };

    /**
     * Get the maximum column index in the worksheet
     *
     * @param {object} sheet Worksheet
     * @return {int} The maximum column index
     */
    var _getMaxColumnIndex = function (sheet) {
        var maxColumn = 0;
        $('cols col', sheet).each(function () {
            var colMax = $(this).attr('max');
            if (colMax > maxColumn) {
                maxColumn = colMax;
            }
        });
        return Number(maxColumn);
    };

    /**
     * Convert column name to index
     *
     * @param {string} columnName Name of the excel column, eg. A, B, C, AB, etc.
     * @return {number} Index number of the column
     */
    var _parseColumnName = function (columnName, sheet) {
        if (typeof columnName == 'number') {
            return columnName;
        }
        // Match last column selector
        if (columnName == '>') {
            return _getMaxColumnIndex(sheet);
        }
        var alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
            i,
            j,
            result = 0;

        for (
            i = 0, j = columnName.length - 1;
            i < columnName.length;
            i += 1, j -= 1
        ) {
            result +=
                Math.pow(alpha.length, j) * (alpha.indexOf(columnName[i]) + 1);
        }

        return Number(result);
    };

    /**
     * Convert index number to Excel column name
     *
     * @param {int} index Index number of column
     * @return {string} Column name
     */
    var _parseColumnIndex = function (index) {
        index -= 1;
        var letter = String.fromCharCode(65 + (index % 26));
        var nextNumber = parseInt(index / 26);
        return nextNumber > 0 ? _parseColumnIndex(nextNumber) + letter : letter;
    };

    /**
     * Datatables config settings, used to calculate smart row references
     */
    //var _tableConfig = {};

    /**
     * Row references for smart row references
     */
    var _rowRefs = {
        t: false, // title
        m: false, // messageTop
        h: false, // header
        dt: false, // Data top row
        db: false, // Data bottom row
        f: false, // footer
        b: false, // messageBottom
    };

    /**
     * Load the row references for smart rows into an object
     *
     * @param {object} config Config options that affect the index of the rows
     * @param {object} sheet Spreadsheet - to calculate length
     */
    function _loadRowRefs(config, sheet) {
        var currentRow = 1;
        // title: Row 1 if it exists
        if (typeof config.title === 'string' && config.title !== '') {
            _rowRefs.t = currentRow;
            currentRow++;
        }
        if (config.messageTop !== null && config.messageTop !== '') {
            _rowRefs.m = currentRow;
            currentRow++;
        }
        if (config.header !== false) {
            _rowRefs.h = currentRow;
            currentRow++;
        }
        _rowRefs.dt = currentRow;

        // Get last row in sheet
        var currentRow = _getMaxSheetRow(sheet);
        if (config.messageBottom !== null && config.messageBottom !== '') {
            _rowRefs.b = currentRow;
            currentRow--;
        }
        if (config.footer !== false) {
            _rowRefs.f = currentRow;
            currentRow--;
        }
        _rowRefs.db = currentRow;
    }

    /**
     * Apply excelStyles to the XML stylesheet
     *
     * @param {object} xlsx
     */
    DataTable.ext.buttons.excelHtml5._applyExcelStyles = function (xlsx) {
        // Load excelStyles and also check exportOptions for backwards compatibility
        var excelStyles = this.excelStyles || this.exportOptions.excelStyles;
        if (excelStyles === undefined) {
            return;
        }
        if (!Array.isArray(excelStyles)) {
            excelStyles = [excelStyles];
        }

        var sheet = xlsx.xl.worksheets['sheet1.xml'];

        // load config settings for smart row references
        var config = DataTable.Api().buttons.exportInfo(this);
        config.header = this.header;
        config.footer = this.footer;
        _loadRowRefs(config, sheet);

        for (var i in excelStyles) {
            var style = excelStyles[i];
            /**
             * A lookup table of existing cell styles and what they should be turned into
             *
             * eg. if existing style is 0, and this style becomes number 54, then any cells with style 1 get turned into 54
             * if there isn't a match in the table, then create the new style.
             */
            var styleLookup = {};

            /**
             * A list of styles created and the cell selectors to apply them to
             */
            var applyTable = {};

            /**
             * Are we using an existing style index rather than a style definition object
             */
            var styleId = false;
            if (style.index !== undefined && typeof style.index === 'number') {
                styleId = style.index;
            }
            var cells = style.cells !== undefined ? style.cells : ['1:'];
            if (!Array.isArray(cells)) {
                cells = [cells];
            }
            var smartRowRef = false;
            if (style.rowref && style.rowref == 'smart') {
                smartRowRef = true;
            }
            for (var i in cells) {
                var selection = _parseExcellyReference(
                    cells[i],
                    sheet,
                    smartRowRef
                );

                // If a valid cell selection is not found, skip this style
                if (selection === false) {
                    continue;
                }

                for (
                    var col = selection.fromCol;
                    col <= selection.toCol;
                    col += selection.nthCol
                ) {
                    var colLetter = _parseColumnIndex(col);
                    for (
                        var row = selection.fromRow;
                        row <= selection.toRow;
                        row += selection.nthRow
                    ) {
                        var tag =
                            'row[r="' +
                            row +
                            '"] c[r="' +
                            colLetter +
                            row +
                            '"]';

                        // Get current style from cell
                        var currentCellStyle = $(tag, sheet).attr('s') || 0;

                        // If a new style hasn't been created, based on this currentCellStyle, then...
                        if (styleLookup[currentCellStyle] == undefined) {
                            var newStyleId;
                            if (currentCellStyle === 0 && styleId) {
                                newStyleId = styleId;
                            } else {
                                // Add a new style based on this current style
                                var merge =
                                    style.merge !== undefined
                                        ? style.merge
                                        : true;
                                var mergeWithCellStyle = merge
                                    ? currentCellStyle
                                    : 0;
                                if (!styleId) {
                                    newStyleId = _addXMLStyle(
                                        xlsx,
                                        style,
                                        mergeWithCellStyle
                                    );
                                } else {
                                    newStyleId = _addXMLStyle(
                                        xlsx,
                                        styleId,
                                        mergeWithCellStyle
                                    );
                                }
                            }
                            styleLookup[currentCellStyle] = newStyleId;
                            applyTable[styleLookup[currentCellStyle]] = [];
                        }
                        applyTable[styleLookup[currentCellStyle]].push(tag);
                    }
                    // Set column width
                    $('col[min="' + col + '"]', sheet)
                        .attr('width', style.width)
                        .attr('customWidth', true);
                }

                // Set row heights
                for (
                    var row = selection.fromRow;
                    row <= selection.toRow;
                    row += selection.nthRow
                ) {
                    if (style.height !== undefined) {
                        $('row[r="' + row + '"]', sheet)
                            .attr('ht', style.height)
                            .attr('customHeight', true);
                    }
                }
            }
            for (var i in applyTable) {
                $(applyTable[i].join(), sheet).attr('s', i);
            }
        }
    };

    /**
     * Attributes to use when translating the simplified excelStyles object
     * to a format that Excel understands
     *
     * @example
     * [rootStyleTag]: { // Main style tag (font|fill|border)
     *    default: {
     *        tagName1: '', // Objects that are required by excel in a particular order
     *        tagName2: '',
     *    },
     *    translate: { // Used to translate commonly used tag names to XML spec name
     *        tagName: 'translatedTagName',
     *    },
     *    [tagName]: { // eg. color, bottom, top, left, right (children of the main style tag)
     *        default: {
     *            tagName1: '', // Child objects required by excel in a particular order
     *            tagName2: '',
     *        },
     *        translate: { // Used to translate commonly used tag names to XML spec name
     *            tagName: 'translatedTagName',
     *        },
     *        val: 'defaultAttributeName', // The attribute name to use in the XML output if value passed as a non-object
     *        [attributeName]: {
     *            tidy: function(val) { // The tidy function to run on attributeName value
     *            },
     *        },
     *        attributeName: 'child', // Any attributes that should be create as a child of the parent tagName
     *    },
     * }
     * @var {object} _translateAttributes
     */
    var _translateAttributes = {
        font: {
            translate: {
                size: 'sz',
                strong: 'b',
                bold: 'b',
                italic: 'i',
                underline: 'u',
            },
            color: {
                val: 'rgb',
            },
        },
        fill: {
            translate: {
                pattern: 'patternFill',
                gradient: 'gradientFill',
            },
            patternFill: {
                default: {
                    patternType: 'solid',
                    fgColor: '',
                    bgColor: '',
                },
                translate: {
                    type: 'patternType',
                    color: 'fgColor',
                },
                replace: 'gradientFill',
                fgColor: {
                    child: true,
                    val: 'rgb',
                },
                bgColor: {
                    child: true,
                    val: 'rgb',
                },
            },
            gradientFill: {
                replace: 'patternFill',
                merge: false,
                stop: {
                    merge: false,
                    child: true,
                    color: {
                        child: true,
                        val: 'rgb',
                    },
                },
            },
        },
        border: {
            default: {
                left: '',
                right: '',
                top: '',
                bottom: '',
                diagonal: '',
                vertical: '',
                horizontal: '',
            },
            top: {
                val: 'style',
                color: {
                    child: true,
                    val: 'rgb',
                },
            },
            bottom: {
                val: 'style',
                color: {
                    child: true,
                    val: 'rgb',
                },
            },
            left: {
                val: 'style',
                color: {
                    child: true,
                    val: 'rgb',
                },
            },
            right: {
                val: 'style',
                color: {
                    child: true,
                    val: 'rgb',
                },
            },
            diagonal: {
                val: 'style',
                color: {
                    child: true,
                    val: 'rgb',
                },
            },
            horizontal: {
                val: 'style',
                color: {
                    child: true,
                    val: 'rgb',
                },
            },
            vertical: {
                val: 'style',
                color: {
                    child: true,
                    val: 'rgb',
                },
            },
        },
    };

    /**
     * Find the node value in the _translateAttributes object
     *
     * @param {array} keyArray Hierarchy of nodes to search
     * @return {any|undefined} Value of the node
     */
    var _findNodeValue = function (keyArray) {
        var val = _translateAttributes;
        for (var i in keyArray) {
            if (keyArray[i] !== null) {
                if (val[keyArray[i]] === undefined) {
                    return undefined;
                }
                val = val[keyArray[i]];
            }
        }
        return val;
    };

    /**
     * Merge object with defaults to fix Excel needing certain fields in a particular order
     *
     * @param {array} nodeHierarchy
     * @param {object} obj Attribute object
     * @return {object} Attribute object merged with object defaults if they exist
     */
    var _mergeDefault = function (nodeHierarchy, obj) {
        var mergeObj = _findNodeValue(nodeHierarchy.concat(['default']));
        if (mergeObj !== undefined) {
            return $.extend({}, mergeObj, obj);
        }
        return obj;
    };

    /**
     * Should this attribute be created as a child node?
     *
     * @param {array} nodeHierarchy
     * @param {string} tagName
     * @param {string} attributeName
     * @return {boolean}
     */
    var _isChildAttribute = function (nodeHierarchy, attributeName) {
        var value = _findNodeValue(nodeHierarchy.concat([attributeName]));
        return (
            value !== undefined &&
            value.child !== undefined &&
            value.child === true
        );
    };

    /**
     * Get translated tagName to translate commonly used html names to XML name (eg size: 'sz')
     *
     * @param {array} nodeHierarchy
     * @param {string} tagName
     * @return {string} Translated tagName if found, otherwise tagName
     */
    var _getTranslatedKey = function (nodeHierarchy, tagName) {
        var newKey = _findNodeValue(
            nodeHierarchy.concat(['translate', tagName])
        );
        return newKey !== undefined ? newKey : tagName;
    };

    /**
     * Get the attributes to add to the node
     *
     * @param {string} styleType
     * @param {string} tagName
     * @param {string} attributeName
     * @param {string} value
     * @param {array}  nodeHierarchy   Array of node names in this tree
     */
    var _getStringAttribute = function (attributeValue, nodeHierarchy) {
        var attributeName = 'val';
        var tKey = _findNodeValue(nodeHierarchy.concat([attributeName]));
        if (tKey !== undefined) {
            attributeName = tKey;
            tKey = _findNodeValue(nodeHierarchy.concat([attributeName]));
        }
        if (tKey !== undefined && tKey.tidy !== undefined) {
            attributeValue = tKey.tidy(attributeValue);
        }
        var obj = {};
        obj[attributeName] = attributeValue;
        return obj;
    };

    /**
     * Add attributes to a node
     *
     * @param {string}          styleType       The type being added (ie. font, fill, border)
     * @param {string}          attributeName   The name of the attribute to add
     * @param {string|object}   attributeValue  The value of the attribute to add
     * @param {object}          parentNode      The parent xml node
     * @param {array}           nodeHierarchy   Array of node names in this tree
     *
     * @todo Replace jQuery function setting attributes when passed an object with plain javascript
     */
    var _addXMLAttribute = function (
        styleType,
        attributeName,
        attributeValues,
        parentNode,
        nodeHierarchy
    ) {
        if (typeof attributeValues === 'object') {
            attributeValues = _mergeDefault(nodeHierarchy, attributeValues);
            for (var attributeKey in attributeValues) {
                var value = attributeValues[attributeKey];
                var key = _getTranslatedKey(nodeHierarchy, attributeKey);
                // if the type is child, create a child node
                if (_isChildAttribute(nodeHierarchy, key)) {
                    _addXMLNode(
                        styleType,
                        key,
                        value,
                        parentNode,
                        nodeHierarchy
                    );
                } else {
                    $(parentNode).attr(key, value);
                }
            }
        } else if (attributeValues !== '') {
            var txAttr = _getStringAttribute(attributeValues, nodeHierarchy);
            $(parentNode).attr(txAttr);
        }
    };

    /**
     * The xml Doc we're working on
     */
    var _xmlStyleDoc;

    /**
     * The xml Doc we're working on
     */
    var _xmlStyleDoc;

    /**
     * Add an XML Node to the tree
     *
     * @param {string}          styleType       The type being added (ie. font, fill, border)
     * @param {string}          attributeName   The name of the attribute to add
     * @param {string|object}   attributeValue  The value of the attribute to add
     * @param {object}          parentNode      The parent xml node
     * @param {array}           nodeHierarchy   Array of node names in this tree
     */
    var _addXMLNode = function (
        styleType,
        attributeName,
        attributeValue,
        parentNode,
        nodeHierarchy
    ) {
        var attributeName = _getTranslatedKey(nodeHierarchy, attributeName);
        _purgeUnwantedSiblings(attributeName, parentNode, nodeHierarchy);
        if (!Array.isArray(attributeValue)) {
            attributeValue = [attributeValue];
        }

        var mergeWith = _doWeMerge(attributeName, nodeHierarchy);

        for (var i in attributeValue) {
            var childNode;
            if ( !mergeWith || parentNode.getElementsByTagName(attributeName).length === 0)
                childNode = parentNode.appendChild(
                    _xmlStyleDoc.createElement(attributeName)
                );
            else {
                childNode = parentNode.getElementsByTagName(attributeName)[0];
            }

            _addXMLAttribute(
                styleType,
                attributeName,
                attributeValue[i],
                childNode,
                nodeHierarchy.concat(attributeName)
            );
        }
    };

    var _doWeMerge = function (attributeName, nodeHierarchy) {
        var merge = _findNodeValue(
            nodeHierarchy.concat([attributeName, 'merge'])
        );
        if( merge !== undefined && merge === false) {
            return false;
        }
        return true;
    };

    var _purgeUnwantedSiblings = function (
        attributeName,
        parentNode,
        nodeHierarchy
    ) {
        var replace = _findNodeValue(
            nodeHierarchy.concat([attributeName, 'replace'])
        );
        if (replace !== undefined) {
            var match = parentNode.getElementsByTagName(replace);
            if (match.length > 0) {
                parentNode.removeChild(match[0]);
            }
        }
    };

    /**
     * Add Style to the stylesheet
     *
     * @param {object} xlsx
     * @param {object|int} addStyle Definition of style to add as an object, or (int) styleID if using a built in style
     */
    var _addXMLStyle = function (xlsx, addStyle, currentCellStyle) {
        if (typeof addStyle === 'object' && addStyle.style === undefined) {
            return currentCellStyle;
        }
        _xmlStyleDoc = xlsx.xl['styles.xml'];
        if (typeof addStyle === 'object') {
            return _mergeWithStyle(addStyle, currentCellStyle);
        } else {
            return _mergeWithBuiltin(addStyle, currentCellStyle);
        }
    };

    /**
     * Merge built-in style with new built-in style to be applied
     *
     * @param {int} builtInIndex Index of the built-in style to apply
     * @param {int} currentCellStyle Current index of the cell being updated
     * @return {int} Index of the newly created style
     */
    var _mergeWithBuiltin = function (builtInIndex, currentCellStyle) {
        var cellXfs = _xmlStyleDoc.getElementsByTagName('cellXfs')[0];

        var currentStyleXf = cellXfs.getElementsByTagName('xf')[
            currentCellStyle
        ];
        var mergeStyleXf = cellXfs.getElementsByTagName('xf')[builtInIndex];

        var xf = cellXfs.appendChild(currentStyleXf.cloneNode(true));

        // Go through all types if any of the type ids are different, clone the elements of those types and change as required
        var types = ['font', 'fill', 'border', 'numFmt'];
        for (var i = 0; i < types.length; i++) {
            var id = types[i] + 'Id';

            if (mergeStyleXf.hasAttribute(id)) {
                if (xf.hasAttribute(id)) {
                    var mergeId = mergeStyleXf.getAttribute(id);
                    var typeId = xf.getAttribute(id);
                    var parentNode = _xmlStyleDoc.getElementsByTagName(
                        types[i] + 's'
                    )[0];

                    var mergeNode = parentNode.childNodes[mergeId];
                    if (mergeId != typeId) {
                        if (id == 'numFmtId') {
                            if (mergeId > 0) {
                                xf.setAttribute(id, mergeId);
                            }
                        } else {
                            var childNode = parentNode.childNodes[
                                typeId
                            ].cloneNode(true);
                            parentNode.appendChild(childNode);
                            _updateContainerCount(parentNode);
                            xf.setAttribute(
                                id,
                                parentNode.childNodes.length - 1
                            );

                            // Cycle through merge children and add/replace
                            var mergeNodeChildren = mergeNode.childNodes;

                            for (
                                var key = 0;
                                key < mergeNodeChildren.length;
                                key++
                            ) {
                                var newAttr = mergeNodeChildren[key].cloneNode(
                                    true
                                );

                                var attr = childNode.getElementsByTagName(
                                    mergeNodeChildren[key].nodeName
                                );
                                if (attr[0]) {
                                    childNode.replaceChild(newAttr, attr[0]);
                                } else {
                                    childNode.appendChild(newAttr);
                                }
                            }
                        }
                    }
                }
            }
        }
        return cellXfs.childNodes.length - 1;
    };

    var _mergeWithStyle = function (addStyle, currentCellStyle) {
        var cellXfs = _xmlStyleDoc.getElementsByTagName('cellXfs')[0];
        var style = addStyle.style;
        var existingStyleXf = cellXfs.getElementsByTagName('xf')[
            currentCellStyle
        ];
        var xf = cellXfs.appendChild(existingStyleXf.cloneNode(true));

        for (var type in style) {
            var typeNode = _xmlStyleDoc.getElementsByTagName(type + 's')[0];
            var parentNode;
            var styleId = type + 'Id';
            if (type == 'alignment') {
                continue;
            } else if (type == 'numFmt') {
                // Handle numFmt style separately as they are a different format
                if (typeof style[type] == 'number') {
                    xf.setAttribute(styleId, style[type]);
                } else {
                    parentNode = _xmlStyleDoc.createElement(type);
                    parentNode.setAttribute('formatCode', style[type]);

                    var lastNumFmtChild = typeNode.lastChild;
                    var lastId = lastNumFmtChild.getAttribute('numFmtId');

                    var numFmtId = Number(lastId) + 1;
                    parentNode.setAttribute('numFmtId', numFmtId);

                    typeNode.appendChild(parentNode);
                    _updateContainerCount(typeNode);

                    xf.setAttribute(styleId, numFmtId);
                }
            } else {
                if (xf.hasAttribute(styleId)) {
                    var existingTypeId = xf.getAttribute(styleId);
                    parentNode = typeNode.childNodes[existingTypeId].cloneNode(
                        true
                    );
                } else {
                    parentNode = _xmlStyleDoc.createElement(type);
                }

                typeNode.appendChild(parentNode);
                style[type] = _mergeDefault([type], style[type]);

                for (var attributeName in style[type]) {
                    var attributeValue = style[type][attributeName];
                    _addXMLNode(
                        type,
                        attributeName,
                        attributeValue,
                        parentNode,
                        [type]
                    ); // fill, patternFill, object|string, parentNode
                }
                xf.setAttribute(styleId, typeNode.childNodes.length - 1);
                _updateContainerCount(typeNode);
            }
        }
        // Add alignment separately
        if (style.alignment !== undefined) {
            _addXMLNode('xf', 'alignment', style.alignment, xf, 'xf');
            xf.setAttribute('applyAlignment', '1');
        }
        _updateContainerCount(cellXfs);
        return cellXfs.childNodes.length - 1;
    };

    /**
     * Update the count attribute on style type containers
     *
     * @param {object} Container node
     */
    var _updateContainerCount = function (container) {
        container.setAttribute('count', container.childNodes.length);
    };

    return DataTable.Buttons;
});
