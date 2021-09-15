/**
 * Styling for DataTables Buttons Excel XLSX (Open Office XML) creation
 *
 * @version: 1.2.0
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
     * @param {boolean} smartRowOption Has the smartRow option been set
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

    /**
     * Get the index number of the last row in the worksheet
     *
     * @param {object} sheet Worksheet
     */
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
            var colMax = Number($(this).attr('max'));
            if (colMax > maxColumn) {
                maxColumn = colMax;
            }
        });
        return maxColumn;
    };

    /**
     * Convert column name to index
     *
     * @param {string} columnName Name of the excel column, eg. A, B, C, AB, etc.
     * @param {object} sheet Worksheet
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
     * Convert a cell name into col and row object
     *
     * @param {string} cellName Name of a cell, eg. B4
     * @return {object} Column and row index
     */
    var _parseCellName = function (cellName) {
        var pattern = /^([A-Z]+)([0-9]+)$/;
        var matches = pattern.exec(cellName);
        if (matches === null) {
            return false;
        }
        return { col: _parseColumnName(matches[1]), row: matches[2] };
    };

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
     * Get a smart reference for the row number
     *
     * @param {int} rowIndex The index of the row
     */
    function _getSmartRefFromIndex(rowIndex) {
        if (rowIndex >= _rowRefs.dt && rowIndex <= _rowRefs.db) {
            return rowIndex - _rowRefs.dt + 1;
        }
        switch (rowIndex) {
            case _rowRefs.t:
                return 't';
            case _rowRefs.m:
                return 'm';
            case _rowRefs.h:
                return 'h';
            case _rowRefs.f:
                return 'f';
            case _rowRefs.b:
                return 'b';
            default:
                return undefined;
        }
    }

    /**
     * Load the row references for smart rows into the _rowRefs object
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
     * Turn a value into an array if it isn't already one
     *
     * @param {any|array} value
     */
    var _makeArray = function (value) {
        if (!Array.isArray(value)) {
            return [value];
        }
        return value;
    };

    /**
     * Insert cells into a spreadsheet
     *
     * // Add cell information (without pushCol or pushRow it replaces any existing data in those cells)
     *
     * insertCells: [
     * {
     *      cells: 'sEh',
     *      content: 'column E',
     * },
     * {
     *      cells: 'sE1:-0',
     *      content: '',
     * }]
     *
     * Use pushCol to push the columns to the right over
     *
     * insertCells: [
     * {
     *      cells: 'sEh',
     *      content: 'column E',
     *      pushCol: true,
     * },
     * {
     *      cells: 'sE1:-0',
     *      content: '',
     *      pushCol: true
     * }]
     *
     * Use pushRow to insert the row, pushing the existing row down by one
     *
     * insertCells: [
     * {
     *   cells: 'sA5',
     *   content: 'THIS IS A ROW BREAK',
     *   pushRow: true,
     * }]
     *
     * @param {*} cells
     * @param {*} xlsx
     */
    var _insertCells = function (insertCells, sheet, config) {
        insertCells = _makeArray(insertCells);
        var maxCol = 0;
        var initialWidth = $('col', sheet).length;
        var maxWidth = 0;
        for (var j in insertCells) {
            var insertObject = insertCells[j];
            var cells =
                insertObject.cells !== undefined
                    ? _makeArray(insertObject.cells)
                    : ['1:'];

            var smartRowRef = false;
            if (insertObject.rowref && style.rowref == 'smart') {
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
                var contentArrayIndex = 0;
                for (
                    var col = selection.fromCol;
                    col <= selection.toCol;
                    col += selection.nthCol
                ) {
                    maxWidth = 0;
                    if (col > maxCol) {
                        maxCol = col;
                    }
                    var colLetter = _parseColumnIndex(col);
                    for (
                        var row = selection.fromRow;
                        row <= selection.toRow;
                        row += selection.nthRow
                    ) {
                        var cellId = String(colLetter) + String(row);
                        var smartRowID = _getSmartRefFromIndex(row);

                        var text = insertObject.content;
                        if (typeof insertObject.content === 'function') {
                            text = insertObject.content(
                                cellId,
                                col,
                                row,
                                smartRowID
                            );
                        }
                        if (Array.isArray(text)) {
                            if (contentArrayIndex >= text.length) {
                                contentArrayIndex = 0;
                            }
                            text = text[contentArrayIndex];
                            contentArrayIndex++;
                        }
                        var width = _calcColWidth(text);
                        if (width > maxWidth) {
                            maxWidth = width;
                        }
                        var cell = _createNode(sheet, 'c', {
                            attr: {
                                t: 'inlineStr',
                                r: cellId,
                            },
                            children: {
                                row: _createNode(sheet, 'is', {
                                    children: {
                                        row: _createNode(sheet, 't', {
                                            text: text,
                                            attr: {
                                                'xml:space': 'preserve',
                                            },
                                        }),
                                    },
                                }),
                            },
                        });
                        var existingCell = _getExistingCell(cellId, sheet);
                        var newCol;
                        if (existingCell !== false) {
                            if (
                                insertObject.pushRow !== undefined &&
                                insertObject.pushRow === true
                            ) {
                                // Insert row
                                var newRow = _createNode(sheet, 'row', {
                                    attr: { r: row },
                                });
                                existingCell.parent().before(newRow);
                                _pushRow(existingCell.parent(), 1);
                                existingCell
                                    .parent()
                                    .nextAll()
                                    .each(function () {
                                        _pushRow($(this), 1);
                                    });
                                newRow.appendChild(cell);
                            } else if (
                                insertObject.pushCol !== undefined &&
                                insertObject.pushCol === true
                            ) {
                                // Insert content
                                existingCell.before(cell);
                                newCol = _pushCol(existingCell, 1);
                                if (newCol > maxCol) {
                                    maxCol = newCol;
                                }
                                existingCell.nextAll().each(function () {
                                    newCol = _pushCol($(this), 1);
                                    if (newCol > maxCol) {
                                        maxCol = newCol;
                                    }
                                });
                            } else {
                                // Replace content
                                existingCell.replaceWith(cell);
                            }
                        } else {
                            // Add content to end
                            $('row', sheet)[row - 1].appendChild(cell);
                        }
                    }
                    _addColIfRequired(col, maxCol, maxWidth, sheet);
                }
            }

            // Update smart row references
            _loadRowRefs(config, sheet);
        }
        _pushMergedColEnd(initialWidth, sheet);
    };

    var _calcColWidth = function (str) {
        var len, lineSplit;

        // from buttons.html5.js
        if (str.indexOf('\n') !== -1) {
            lineSplit = str.split('\n');
            lineSplit.sort(function (a, b) {
                return b.length - a.length;
            });

            len = lineSplit[0].length;
        } else {
            len = str.length;
        }
        return len;
    };

    var _addColIfRequired = function (insertedCol, maxCol, maxWidth, sheet) {
        // Update columns
        var sheetColCount = sheet.getElementsByTagName('col').length;
        if (sheetColCount < maxCol) {
            if (maxWidth > 40) {
                maxWidth = 40;
            }
            if (maxWidth < 6) {
                maxWidth = 6;
            }
            maxWidth *= 1.35;
            var insertBefore = sheet.getElementsByTagName('col')[
                insertedCol - 1
            ];
            for (var i = sheetColCount + 1; i <= maxCol; i++) {
                var newCol = _createNode(sheet, 'col', {
                    attr: {
                        min: i,
                        max: i,
                        width: maxWidth,
                        customWidth: 1,
                    },
                });

                sheet
                    .getElementsByTagName('cols')[0]
                    .insertBefore(newCol, insertBefore);
            }
            _updateCellMinMax(sheet);
        } else {
            // update width if required
            if (maxWidth > 40) {
                maxWidth = 40;
            }
            if (maxWidth < 6) {
                maxWidth = 6;
            }
            maxWidth *= 1.35;
            var column = sheet.getElementsByTagName('col')[insertedCol - 1];
            var currentWidth = $(column).attr('width');
            if (currentWidth < 6) {
                currentWidth = 6;
            }
            if (maxWidth > currentWidth) {
                $(column).attr('width', maxWidth);
                $(column).attr('customWidth', 1);
            }
        }
    };

    var _updateCellMinMax = function (sheet) {
        var cells = sheet.getElementsByTagName('col');
        for (var i = 0; i < cells.length; i++) {
            var cell = $(cells[i]);
            cell.attr('min', i + 1);
            cell.attr('max', i + 1);
        }
    };

    var _getExistingCell = function (cellId, sheet) {
        var cell = $('sheetData row c[r="' + cellId + '"]', sheet);
        if (cell.length === 0) {
            return false;
        } else {
            return cell;
        }
    };

    var _pushMergedColEnd = function (initialWidth, sheet) {
        var newWidth = $('col', sheet).length;
        if (newWidth == initialWidth) {
            return;
        }
        var mergeCells = sheet.getElementsByTagName('mergeCell');
        if (mergeCells.length > 0) {
            for (var i = 0; i < mergeCells.length; i++) {
                var mc = mergeCells[i];
                var ref = _parseExcellyReference(
                    $(mc).attr('ref'),
                    sheet,
                    false
                );
                if (ref.toCol >= initialWidth) {
                    var newRef =
                        _parseColumnIndex(ref.fromCol) +
                        String(ref.fromRow) +
                        ':' +
                        _parseColumnIndex(newWidth) +
                        String(ref.toRow);
                    $(mc).attr('ref', newRef);
                }
            }
        }
    };

    var _pushRow = function (row, rowsToPush) {
        var rowID = row.attr('r');
        var newRowID = parseInt(rowID) + rowsToPush;
        row.attr('r', newRowID);
        row.children().each(function () {
            var cell = $(this);
            var cellID = cell.attr('r');
            var cellColRow = _parseCellName(cellID);
            var newCellID =
                String(_parseColumnIndex(cellColRow.col)) + String(newRowID);
            cell.attr('r', newCellID);
        });
    };

    var _pushCol = function (cell, colsToPush) {
        var cellID = cell.attr('r');
        var cellColRow = _parseCellName(cellID);
        var newCellID =
            String(_parseColumnIndex(cellColRow.col + colsToPush)) +
            String(cellColRow.row);
        cell.attr('r', newCellID);
        return cellColRow.col + 1;
    };

    var _createNode = function (doc, nodeName, opts) {
        var tempNode = doc.createElement(nodeName);

        if (opts) {
            if (opts.attr) {
                $(tempNode).attr(opts.attr);
            }

            if (opts.children) {
                $.each(opts.children, function (key, value) {
                    tempNode.appendChild(value);
                });
            }

            if (opts.text !== null && opts.text !== undefined) {
                tempNode.appendChild(doc.createTextNode(opts.text));
            }
        }

        return tempNode;
    };

    /**
     * Apply excelStyles to the XML stylesheet
     *
     * @param {object} xlsx
     */
    DataTable.ext.buttons.excelHtml5._applyExcelStyles = function (xlsx) {
        // Load excelStyles and also check exportOptions for backwards compatibility
        // var excelStyles = this.excelStyles || this.exportOptions.excelStyles;
        // if (excelStyles === undefined) {
        //     return;
        // }
        // if (!Array.isArray(excelStyles)) {
        //     excelStyles = [excelStyles];
        // }

        var sheet = xlsx.xl.worksheets['sheet1.xml'];
        _xmlStyleDoc = xlsx.xl['styles.xml'];

        // load config settings for smart row references
        var config = DataTable.Api().buttons.exportInfo(this);
        config.header = this.header;
        config.footer = this.footer;
        _loadRowRefs(config, sheet);

        if (this.insertCells !== undefined) {
            _insertCells(this.insertCells, sheet, config);
        }

        if (this.pageStyle !== undefined) {
            _applyPageStyle(this.pageStyle, sheet, xlsx);
        }

        // Load excelStyles and also check exportOptions for backwards compatibility
        var excelStyles = this.excelStyles || this.exportOptions.excelStyles;
        if (excelStyles === undefined) {
            return;
        }
        excelStyles = _makeArray(excelStyles);

        /**
         * Cache the links to the spreadsheet cells
         */
        var tag_cache = [];

        var sheet_data = sheet.querySelectorAll('sheetData row');

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
             * Are we using an existing style index rather than a style definition object
             */
            var styleId = false;
            if (style.index !== undefined && typeof style.index === 'number') {
                styleId = style.index;
            }

            var cells =
                style.cells !== undefined ? _makeArray(style.cells) : ['1:'];

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

                // If a condition is supplied, add this style as a conditional style
                if (style.condition != undefined) {
                    _addConditionalStyle(sheet, style, selection);
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
                            colLetter +
                            row;

                        
                        // Get current style from cell
                        if(tag_cache[tag] == undefined) {

                            // var searchTag =
                            //     'row[r="' +
                            //     row +
                            //     '"] c[r="' +
                            //     colLetter +
                            //     row +
                            //     '"]';
                            //var cacheCellRef = sheet.querySelector(searchTag);

                            // New - cell selection in version 1.2 
                            // The next four lines can be replaced with the above commented code if this doesn't work
                            if(sheet_data[row-1] == undefined || sheet_data[row-1].childNodes[col-1] == undefined) {
                                continue;
                            }
                            var cacheCellRef = sheet_data[row-1].childNodes[col-1];

                            var cellInitialStyle = cacheCellRef.getAttribute('s') || 0;
                            tag_cache[tag] = {
                                cellRef: cacheCellRef,
                                initialStyle: cellInitialStyle,
                                style: cellInitialStyle,
                            };
                        }
                        var currentCellStyle = tag_cache[tag].style;

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
                                        style,
                                        mergeWithCellStyle
                                    );
                                } else {
                                    newStyleId = _addXMLStyle(
                                        styleId,
                                        mergeWithCellStyle
                                    );
                                }
                            }
                            styleLookup[currentCellStyle] = newStyleId;
                        }
                        tag_cache[tag].style = styleLookup[currentCellStyle];
                    }
                    // Set column width
                    if (style.width !== undefined) {
                        var colref = sheet.querySelector('col[min="' + col + '"]');
                        colref.setAttribute('width', style.width);
                        colref.setAttribute('customWidth', true);
                    }
                }

                // Set row heights
                for (
                    var row = selection.fromRow;
                    row <= selection.toRow;
                    row += selection.nthRow
                ) {
                    if (style.height !== undefined) {
                        var rwref = sheet.querySelector('row[r="' + row + '"]');
                        rwref.setAttribute('ht', style.height);
                        rwref.setAttribute('customHeight', true);
                    }
                }
            }
        }
        for(var i in tag_cache) {
            if(tag_cache[i].style != tag_cache[i].initialStyle) {
                tag_cache[i].cellRef.setAttribute('s', tag_cache[i].style);
            }
        }
    };

    var _applyPageStyle = function (pageStyle, sheet, xlsx) {
        pageStyle = _mergeDefault(['worksheet'], pageStyle);
        for (var type in pageStyle) {
            var attributeValue = pageStyle[type];
            switch (type) {
                case 'repeatHeading':
                case 'repeatRow':
                    _addRepeatHeading(attributeValue, sheet, xlsx);
                    break;
                case 'repeatCol':
                    _addRepeatColumns(attributeValue, sheet, xlsx);
                    break;
                default:
                    var parentNode = sheet.getElementsByTagName('worksheet')[0];
                    _addXMLNode('pageStyle', type, attributeValue, parentNode, [
                        'worksheet',
                    ]);
            }
        }
    };

    /**
     * Add the xml to repeat the page heading on each printed page
     * 
     * Use 'repeatHeading: value' in the pageStyle object to define.
     * 
     * The value can be:
     *      true - to repeat the heading row on every page
     *      An excelly row reference (eg. st:h to repeat the title and heading on each page)
     */
    var _addRepeatHeading = function(value, sheet, xlsx) {
        var rows = 'sh:h';
        if (value !== true && value !== false) {
            rows = value;
        }
        var rowSelection = _parseExcellyReference(rows, sheet, false);
        var selectionString =
            'Sheet1!$' + rowSelection.fromRow + ':$' + rowSelection.toRow;

        _addRepeat(selectionString, xlsx);
    }

    /**
     * Allow repeating of columns as well using repeatCol option
     */
    var _addRepeatColumns = function (value, sheet, xlsx) {
        var cols = 'A:A';
        if (value !== true && value !== false) {
            cols = value;
        }
        var colSelection = _parseExcellyReference(cols, sheet, false);
        var selectionString =
            'Sheet1!$' +
            _parseColumnIndex(colSelection.fromCol) +
            ':$' +
            _parseColumnIndex(colSelection.toCol);

        _addRepeat(selectionString, xlsx);
    };

    var _addRepeat = function (selectionString, xlsx) {
        var workbook = xlsx.xl['workbook.xml'];
        var parentNode = workbook.getElementsByTagName('workbook')[0];

        var repeats = []
        
        var existing = workbook.getElementsByName('_xlnm.Print_Titles');
        if(existing.length > 0) {
            repeats.push(existing[0].textContent);
            existing[0].textContent = '';
        }
        repeats.push(selectionString);

        var addObject = {
            definedName: {
                name: '_xlnm.Print_Titles',
                localSheetId: '0',
                rows: repeats.join(','),
            },
        };
        _addXMLNode('workbook', 'definedNames', addObject, parentNode, [
            'workbook',
        ]);
    };

    /**
     * Internal attributes to use when translating the simplified Excel Style Objects
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
        conditionalFormatting: {
            cfRule: {
                default: {
                    priority: '1',
                },
                formula: {
                    child: true,
                    merge: false,
                    val: 'formulaValue',
                    formulaValue: {
                        textNode: true,
                    },
                },
                dataBar: {
                    child: true,
                    default: {
                        cfvo: [
                            { type: 'min', val: 0 },
                            { type: 'max', val: 0 },
                        ],
                    },
                    cfvo: {
                        child: true,
                        merge: false,
                    },
                    color: {
                        child: true,
                        val: 'rgb',
                    },
                },
                colorScale: {
                    child: true,
                    default: {
                        cfvo: [
                            { type: 'min', val: 0 },
                            { type: 'max', val: 0 },
                        ],
                    },
                    cfvo: {
                        child: true,
                        merge: false,
                    },
                    color: {
                        child: true,
                        merge: false,
                        val: 'rgb',
                    },
                },
                iconSet: {
                    child: true,
                    default: {
                        iconSet: '4Rating',
                        cfvo: [
                            { type: 'percentile', val: 0 },
                            { type: 'percentile', val: 33 },
                            { type: 'percentile', val: 67 },
                            { type: 'percentile', val: 100 },
                        ],
                    },
                    cfvo: {
                        child: true,
                        merge: false,
                    },
                },
            },
        },
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
        worksheet: {
            default: {
                printOptions: '',
                pageMargins: {
                    left: '0.7',
                    right: '0.7',
                    top: '0.75',
                    bottom: '0.75',
                    header: '0.3',
                    footer: '0.3',
                },
                pageSetup: '',
            },
            mergeCells: {
                updateCount: true,
                mergeCell: {
                    child: true,
                    merge: false,
                    val: 'ref',
                }
            },
            sheetPr: {
                insertBefore: 'cols',
                pageSetUpPr: {
                    child: true,
                },
            },
        },
        // The workbook area is only used at this stage to allow the repeatHeading option
        workbook: {
            definedNames: {
                insertBefore: 'calcPr',
                definedName: {
                    child: true,
                    rows: {
                        textNode: true,
                        merge: false,
                    },
                },
            },
        },
    };

    /**
     * Find a node value in the _translateAttributes object
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

    var _doUpdateCount = function (nodeHierarchy, attributeName) {
        var value = _findNodeValue(nodeHierarchy.concat([attributeName]));
        return (
            value !== undefined &&
            value.updateCount !== undefined &&
            value.updateCount === true
        );
    };

    /**
     * Should this attribute be created as a textNode?
     *
     * @param {array} nodeHierarchy
     * @param {string} attributeName
     * @return {boolean}
     */
    var _isTextNode = function (nodeHierarchy, attributeName) {
        var value = _findNodeValue(nodeHierarchy.concat([attributeName]));
        return (
            value !== undefined &&
            value.textNode !== undefined &&
            value.textNode === true
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

    var _getAppendPosition = function (nodeHierarchy, tagName) {
        var value = _findNodeValue(
            nodeHierarchy.concat([tagName, 'insertBefore'])
        );
        return value === undefined ? 'end' : value;
    };

    /**
     * Get the attributes to add to the node
     *
     * @param {string} attributeValue
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
                    if (_isTextNode(nodeHierarchy, key)) {
                        parentNode.appendChild(
                            _xmlStyleDoc.createTextNode(value)
                        );
                    } else {
                        $(parentNode).attr(key, value);
                    }
                }
            }
        } else if (attributeValues !== '') {
            var txAttr = _getStringAttribute(attributeValues, nodeHierarchy);
            for (var i in txAttr) {
                if (_isTextNode(nodeHierarchy, i)) {
                    parentNode.appendChild(
                        _xmlStyleDoc.createTextNode(txAttr[i])
                    );
                } else {
                    parentNode.setAttribute(i, txAttr[i]);
                }
            }
        }
    };

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
        attributeValue = _makeArray(attributeValue);
        var mergeWith = _doWeMerge(attributeName, nodeHierarchy);

        for (var i in attributeValue) {
            var childNode;
            if (
                !mergeWith ||
                parentNode.getElementsByTagName(attributeName).length === 0
            ) {
                var position = _getAppendPosition(nodeHierarchy, attributeName);
                if (position === 'end') {
                    childNode = parentNode.appendChild(
                        _xmlStyleDoc.createElement(attributeName)
                    );
                } else {
                    var beforeNode = parentNode.getElementsByTagName(
                        position
                    )[0];
                    childNode = parentNode.insertBefore(
                        _xmlStyleDoc.createElement(attributeName),
                        beforeNode
                    );
                }
            } else {
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
        if(_doUpdateCount(nodeHierarchy, attributeName)) {
            _updateContainerCount(childNode);
        }
    };

    /**
     * Determine if we should merge attributes or replace them
     *
     * To fix issues with gradientFill options causing Excel to throw an error
     *
     * @param {string} attributeName Name of the attributes
     * @param {array} nodeHierarchy Array of node names in this tree
     */
    var _doWeMerge = function (attributeName, nodeHierarchy) {
        var merge = _findNodeValue(
            nodeHierarchy.concat([attributeName, 'merge'])
        );
        if (merge !== undefined && merge === false) {
            return false;
        }
        return true;
    };

    /**
     * Remove node siblings which would cause Excel to throw an error
     *
     * eg. You can't apply a patternFill and a gradientFill to the same call
     *
     * @param {string} attributeName Name of the attribute
     * @param {object} parentNode The parent xml node
     * @param {array} nodeHierarchy Array of node names in this tree
     */
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
     * Add Style to the stylesheet using either a built-in style or a custom defined style
     *
     * @param {object|int} addStyle Definition of style to add as an object, or (int) styleID if using a built in style
     * @param {object|int} currentCellStyle The current style of the cell to merge with
     * @return {int} Style ID
     */
    var _addXMLStyle = function (addStyle, currentCellStyle) {
        if (typeof addStyle === 'object' && addStyle.style === undefined) {
            return currentCellStyle;
        }
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

    /**
     * Merge existing cell style with the new custom Excel Style to be applied
     *
     * @param {object} addStyle Excel Style Object to be applied to the cell
     * @param {int} currentCellStyle Current index of the cell being updated
     * @return {int} Index of the newly created style
     */
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
     * Add conditional formatting to a spreadsheet
     *
     * @param {xls} sheet
     * @param {object} excelStyle ExcelStyle object
     * @param {array} selection The cell range selected
     */
    var _addConditionalStyle = function (sheet, excelStyle, selection) {
        // Create new dxf incremental formatting style
        var dxfs = _xmlStyleDoc.getElementsByTagName('dxfs')[0];
        var dxfNode = _xmlStyleDoc.createElement('dxf');
        dxfs.appendChild(dxfNode);
        _updateContainerCount(dxfs);

        // Add style to dxf block
        var style = excelStyle.style ? excelStyle.style : {};
        var parentNode;
        for (var type in style) {
            parentNode = _xmlStyleDoc.createElement(type);
            dxfNode.appendChild(parentNode);
            style[type] = _mergeDefault([type], style[type]);

            for (var attributeName in style[type]) {
                var attributeValue = style[type][attributeName];
                _addXMLNode(type, attributeName, attributeValue, parentNode, [
                    type,
                ]);
            }
        }

        var dxfId = dxfs.childNodes.length - 1;

        var worksheet = sheet.getElementsByTagName('worksheet')[0];
        var conditionalFormatting = sheet.createElement(
            'conditionalFormatting'
        );

        var cellRef = _getRangeFromSelection(selection);
        conditionalFormatting.setAttribute('sqref', cellRef);
        worksheet.appendChild(conditionalFormatting);

        var condition = excelStyle.condition;
        condition.dxfId = dxfId;
        _addXMLNode(
            'conditionalFormatting',
            'cfRule',
            condition,
            conditionalFormatting,
            ['conditionalFormatting']
        );
    };

    /**
     * Convert a cell selection into a Range (ignoring cell skipping, etc.)
     *
     * @param {array} selection Parsed excelly reference
     * @return {string} Cell range, eg. "A3:A45"
     */
    var _getRangeFromSelection = function (selection) {
        return (
            _parseColumnIndex(selection.fromCol) +
            String(selection.fromRow) +
            ':' +
            _parseColumnIndex(selection.toCol) +
            String(selection.toRow)
        );
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
