/**
 * Styling for Datatables Buttons Excel XLSX (OOXML) creation
 *
 * @version: 0.5
 * @description Add and process a custom 'excelStyles' option to easily customize the Datatables Excel Stylesheet output
 * @file buttons.html5.styles.js
 * @copyright © 2020 Beyond the Box Creative
 * @author Paul Jones <info@pauljones.co.nz>
 * @license MIT
 *
 * Include this file after including the javascript for the Datatables, Buttons, HTML5 and JSZip extensions
 *
 * Create the required styles using the custom 'excelStyles' option in the button's 'exportOptions'
 * @see https://datatables.net/reference/button/excelHtml5 For exportOptions information
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
     *      exportOptions: {
     *          excelStyles: {
     *              ... custom styles defined ...
     *          },
     *          customize: function(xlsx) {
     *              this.applyStyles(xlsx);
     *              ... custom code here ...
     *          }
     *      }
     * }
     */
    DataTable.ext.buttons.excelHtml5.applyStyles = function (xlsx) {
        var excelStyles = this.exportOptions.excelStyles;
        if (excelStyles !== undefined) {
            if (!Array.isArray(excelStyles)) {
                excelStyles = [excelStyles];
            }
            this._applyExcelStyles(xlsx, excelStyles);
        }
    };

    /**
     * Parse cell names into a row, col, from:to like structure
     *
     * @example
     * Cell reference examples
     *
     * Single Range References
     * ['A3']   = cell A3
     * ['4']    = row 4, all columns
     * ['D']    = column D, all rows
     *
     * Multiple Range References (seperated by :)
     * ['4:6']      = rows 4 to 6, all columns
     * ['B:F']      = column B to F, all rows
     * ['D4:D20']   = column D from row 4 to 20
     * ['3:']       = from row 3 until the last row, all columns
     * ['A:']       = from column A until the last column, all rows
     * ['B3:']      = from column B until the last column, from row 3 until the last row
     * [':B3']      = from column A until column B, from row 1 to row 3
     * ['B3:D']     = from column B until column D, from row 3 until the last row
     *
     * References to the last column
     * ['>']        = all rows, the last column
     * ['>3:>20']   = the last column, row 3 to 20
     *
     * References counting back from the last column
     * ['-3>']      = three columns back from the last column
     * ['-2>5']     = two columns back from the last column, row 5
     *
     * References counting back from the last row
     * ['-0']       = all columns, the last row
     * ['B-3:B-0']  = column B from the third to last row until the last row
     *
     * Reference for everything
     * [':']        = all columns, all rows (also ['1:'], ['A:'], [''], ['A1:'], [':-0'], [':>'], [':>-0'])
     *
     *
     * Column/Row skipping
     *
     * Used to apply styles to every nth column or row (eg. every 2nd row, every 3rd column)
     *
     * Format: n[0-9],[0-9]
     * n (stands for every nth column/row), then the column increment followed by the row increment
     *
     * Column/Row skipping examples
     * ['A3:D10n1,2']   = from Column A row 3, to Column D row 10, target every column, target every second row
     * ['3:n1,2']       = every column from row 3 until the last row, target every second row (use this for row striping)
     * [':n1,2']        = every column, every second row (also [':n,2'])
     * [':n2,1']        = every second column, every row (also [':n2'])
     *
     *
     * @param {string} cells Cell names in an Excel-like structure
     * @param {object} sheet The worksheet to enable finding of the last column/row
     * @return {object} Parsed rows and columns, in number format (ie. columns refernced by number, not letter)
     */
    var _parseExcellyReference = function (cells, sheet) {
        //var pattern = /^(-\d+(?=\>))*([A-Z]*|[>])*(-)*([0-9]*)(\:)*(-\d+(?=\>))*([A-Z]*|[>])*(-)*([0-9]*)(?:n([0-9]*)(?:,)*([0-9]*))*$/;
        var pattern = /^(?:-(\d*)(?=\>))*([A-Z]*|[>])*(-(?=[0-9]+))*([0-9]*)(?:(\:)(?:-(\d*)(?=\>))*([A-Z]*|[>])*(-(?=[0-9]+))*([0-9]*)(?:n([0-9]*)(?:,)*([0-9]*))*)*$/;
        var matches = pattern.exec(cells);
        var results = {
            fromColEndSubtractAmount: matches[1],
            fromCol: matches[2],
            fromRowEndSubtract: matches[3],
            fromRow: matches[4],
            range: matches[5],
            toColEndSubtractAmount: matches[6],
            toCol: matches[7],
            toRowEndSubtract: matches[8],
            toRow: matches[9],
            nthCol: matches[10],
            nthRow: matches[11],
        };

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
                    : _getMaxRow(sheet) - results.toRow // else return last row minus this row number
                : null) || // else return null and continue
            (results.range || !results.fromRow // if there is a range selected, but no fromRow
                ? _getMaxRow(sheet) // return the maximum row
                : !results.fromRowEndSubtract // else if we are NOT subtracting from the last row for the from source
                ? results.fromRow // return the from row
                : _getMaxRow(sheet) - results.fromRow); // else return the last row minus the from row number

        results.toRow = parseInt(results.toRow);
        results.fromRow = results.fromRow
            ? parseInt(
                  !results.fromRowEndSubtract
                      ? results.fromRow
                      : _getMaxRow(sheet) - results.fromRow
              )
            : 1;
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
     * Get the maximum row number in the worksheet
     *
     * @param {object} sheet Worksheet
     * @return {int} The maximum row number
     */
    var _getMaxRow = function (sheet) {
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
    var _parseColumnName = function(columnName, sheet) {
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
    }

    /**
     * Convert index number to Excel column name
     *
     * @param {int} index Index number of column
     * @return {string} Column name
     */
    var _parseColumnIndex = function(index) {
        index -= 1;
        var letter = String.fromCharCode(65 + (index % 26));
        var nextNumber = parseInt(index / 26);
        return nextNumber > 0 ? _parseColumnIndex(nextNumber) + letter : letter;
    }

    

    /**
     * Apply exportOptions.excelStyles to the OOXML stylesheet
     *
     * @param {object} xlsx
     */
    DataTable.ext.buttons.excelHtml5._applyExcelStyles = function (xlsx, excelStyles) {
        var sheet = xlsx.xl.worksheets['sheet1.xml'];

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

            for (var i in cells) {
                var selection = _parseExcellyReference(cells[i], sheet);

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
     * styleTag: { // Main style tag (font|fill|border)
     *    default: {
     *        tagName1: '', // Objects that are required by excel in a particular order
     *        tagName2: '',
     *    },
     *    translate: { // Used to translate commonly used tag names to OOXML name
     *        tagName: 'translatedTagName',
     *    },
     *    tagName: { // eg. color, bottom, top, left, right (children of the main style tag)
     *        default: {
     *            tagName1: '', // Child objects required by excel in a particular order
     *            tagName2: '',
     *        },
     *        val: 'defaultAttributeName', // The attribute name to use in the OOXML if value passed as a non-object
     *        attributeName: {
     *            tidy: function(val) { // The tidy function to run on attributeName value
     *            },
     *        },
     *        attributeName: 'child', // Any attributes that should be create as a child of the parent tagName
     *    },
     * }
     * @var {object} _translateAttribues
     */
    var _translateAttributes = {
        font: {
            translate: {
                size: 'sz',
                strong: 'b',
            },
            color: {
                val: 'rgb',
                rgb: {
                    tidy: function (val) {
                        return /([A-F0-9]{3,6})/.exec(val)[1].toUpperCase();
                    },
                },
            },
        },
        fill: {
            fgColor: {
                val: 'rgb',
            },
            bgColor: {
                val: 'rgb',
            },
            patternFill: {
                default: {
                    patternType: '',
                    fgColor: '',
                    bgColor: '',
                },
                fgColor: 'child',
                bgColor: 'child',
            },
        },
        border: {
            default: {
                left: '',
                right: '',
                top: '',
                bottom: '',
                diagonal: '',
            },
            color: {
                val: 'rgb',
            },
            top: {
                val: 'style',
                color: 'child',
            },
            bottom: {
                val: 'style',
                color: 'child',
            },
            left: {
                val: 'style',
                color: 'child',
            },
            right: {
                val: 'style',
                color: 'child',
            },
        },
    };

    /**
     * Find the node value in the _translateAttribues object
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
     * @param {string} parentTagName
     * @param {string} tagName
     * @param {object} obj Attribute object
     * @return {object} Attribute object merged with object defaults if they exist
     */
    var _mergeDefault = function (parentTagName, tagName, obj) {
        var mergeObj = _findNodeValue([parentTagName, tagName, 'default']);
        if (mergeObj !== undefined) {
            return $.extend({}, mergeObj, obj);
        }
        return obj;
    };

    /**
     * Should this attribute be created as a child node?
     *
     * @param {string} parentTagName
     * @param {string} tagName
     * @param {string} attributeName
     * @return {boolean}
     */
    var _isChildAttribute = function (parentTagName, tagName, attributeName) {
        return (
            _findNodeValue([parentTagName, tagName, attributeName]) == 'child'
        );
    };

    /**
     * Get translated tagName to translate commonly used html names to OOXML name (eg size: 'sz')
     *
     * @param {string} parentTagName
     * @param {string} tagName
     * @return {string} Translated tagName if found, otherwise tagName
     */
    var _getTranslatedKey = function (parentTagName, tagName) {
        var newKey = _findNodeValue([parentTagName, 'translate', tagName]);
        return newKey !== undefined ? newKey : tagName;
    };

    /**
     * Get the attributes to add to the node
     *
     * @param {string} parentTagName
     * @param {string} tagName
     * @param {string} attributeName
     * @param {string} value
     */
    var _getTagAttributes = function (
        parentTagName,
        tagName,
        attributeName,
        value
    ) {
        var tKey = _findNodeValue([parentTagName, tagName, attributeName]);
        if (tKey !== undefined && attributeName == 'val') {
            attributeName = tKey;
            tKey = _findNodeValue([parentTagName, tagName, attributeName]);
        }
        if (tKey !== undefined && tKey.tidy !== undefined) {
            value = tKey.tidy(value);
        }
        var obj = {};
        obj[attributeName] = value;
        return obj;
    };

    /**
     * Add attributes to a node
     *
     * @param {string} tagName
     * @param {string} attributeName
     * @param {string|object} value Attribute Value
     * @param {obj} parentNode
     *
     * @todo Replace jQuery function setting attributes when passed an object with plain javascript
     */
    var _addXMLAttribute = function (
        tagName,
        attributeName,
        value,
        parentNode
    ) {
        if (typeof value === 'object') {
            value = _mergeDefault(tagName, attributeName, value);
            for (var i in value) {
                // if the type is child, create a child node
                if (_isChildAttribute(tagName, attributeName, i)) {
                    _addXMLNode(tagName, i, value[i], parentNode);
                } else {
                    $(parentNode).attr(i, value[i]);
                }
            }
        } else if (value != '') {
            var txAttr = _getTagAttributes(
                tagName,
                attributeName,
                'val',
                value
            );
            $(parentNode).attr(txAttr);
        }
    };

    /**
     * The xml Doc we're working on
     */
    var _xmlStyleDoc;

    /**
     * Add an XML Node to the tree
     *
     * @param {string} tagName
     * @param {string} attributeName
     * @param {string|object} value
     * @param {object} parentNode
     */
    var _addXMLNode = function (tagName, attributeName, value, parentNode) {
        var key = _getTranslatedKey(tagName, attributeName);
        var childNode;
        if (parentNode.getElementsByTagName(key).length === 0)
            childNode = parentNode.appendChild(_xmlStyleDoc.createElement(key));
        else {
            childNode = parentNode.getElementsByTagName(key)[0];
        }
        _addXMLAttribute(tagName, attributeName, value, childNode);
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

    /**
     * Merge current cell style with new custom style to be applied
     *
     * @param {object} addStyle Excelstyles style object to be applied to cell
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
            var node;
            var styleId = type + 'Id';
            if (type == 'alignment') {
                continue;
            } else if (type == 'numFmt') {
                // Handle numFmt style seperately as they are a different format
                if (typeof style[type] == 'number') {
                    xf.setAttribute(styleId, style[type]);
                } else {
                    node = _xmlStyleDoc.createElement(type);
                    node.setAttribute('formatCode', style[type]);

                    //var numFmts = _xmlStyleDoc.getElementsByTagName('numFmts')[0];
                    var lastNumFmtChild = typeNode.lastChild;
                    var lastId = lastNumFmtChild.getAttribute('numFmtId');

                    var numFmtId = Number(lastId) + 1;
                    node.setAttribute('numFmtId', numFmtId);

                    typeNode.appendChild(node);
                    _updateContainerCount(typeNode);

                    xf.setAttribute(styleId, numFmtId);
                }
            } else {
                if (xf.hasAttribute(styleId)) {
                    var existingTypeId = xf.getAttribute(styleId);
                    node = typeNode.childNodes[existingTypeId].cloneNode(true);
                } else {
                    node = _xmlStyleDoc.createElement(type);
                }

                typeNode.appendChild(node);

                style[type] = _mergeDefault(type, null, style[type]);

                for (var attr in style[type]) {
                    var value = style[type][attr];
                    _addXMLNode(type, attr, value, node); // fill, patternFill, object|string, parentNode
                }
                xf.setAttribute(styleId, typeNode.childNodes.length - 1);

                _updateContainerCount(typeNode);
            }
        }
        // Add alignment seperately
        if (style.alignment !== undefined) {
            _addXMLNode('xf', 'alignment', style.alignment, xf);
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
