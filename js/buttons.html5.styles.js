/**
 * Styling for Datatables Buttons Excel XLSX (OOXML) creation
 *
 * @version: 0.3
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
 * Documentation on 'excelStyles' options
 * 'Coming soon...'
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
        var sheet = xlsx.xl.worksheets['sheet1.xml'];
        console.log(sheet);
    };

    /**
     * Allow applyStyles to be triggered from a custom customize function
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
        if (this.exportOptions.excelStyles !== undefined) {
            _applyStyles(xlsx, this.exportOptions.excelStyles);
        }
    };

    /**
     * Apply exportOptions.excelStyles to the OOXML stylesheet
     *
     * @param {object} xlsx
     */
    var _applyStyles = function (xlsx, excelStyles) {
        var sheet = xlsx.xl.worksheets['sheet1.xml'];
        //var excelStyles = this.exportOptions.excelStyles;

        for (var i in excelStyles) {
            var style = excelStyles[i];
            var styleId;
            if (style.index !== undefined && typeof style.index === 'number') {
                styleId = style.index;
            } else {
                styleId = _addXMLStyle(xlsx, style);
            }

            var rows = style.row !== undefined ? style.row : [0];
            var columns = style.column !== undefined ? style.column : [];
            if (!Array.isArray(columns)) {
                columns = [columns];
            }
            if (!Array.isArray(rows)) {
                rows = [rows];
            }

            var selectors = [];

            for (var i in rows) {
                var prow = _parseRowSelector(rows[i]);
                var rowStart = prow.start;
                var rowEnd = prow.end;
                var rowInc = prow.inc;
                if (rowInc < 1) {
                    rowInc = 1;
                }

                if (rowEnd == '') {
                    rowEnd = $('row', sheet).length;
                }

                if (rowEnd < rowStart) {
                    continue;
                }

                for (var row = rowStart; row <= rowEnd; row += rowInc) {
                    var rowSelector =
                        'row' + (row > 0 ? '[r="' + row + '"]' : '');
                    if (columns.length == 0) {
                        selectors.push(rowSelector + ' c');
                    } else {
                        for (var i in columns) {
                            var colSelector = ' c[r^="' + columns[i] + '"]';
                            selectors.push(rowSelector + colSelector);
                        }
                    }
                    if(style.height !== undefined) {
                        $('row', sheet).eq(row-1).attr('ht',style.height).attr('customHeight', true);
                    }
                }
                
            }
            if (styleId >= 0) {
                $(selectors.join(), sheet).attr('s', styleId);
            } 
            if (columns.length > 0 && style.width !== undefined) {
                for (var i in columns) {
                    $('col', sheet).eq(_excelColumnToIndex(columns[i])-1).attr('width', style.width).attr('customWidth', true);
                }
            }
        }
    };

    /**
     * Convert column name to index
     * 
     * @param {string} columnName Name of the excel column, eg. A, B, C, AB, etc.
     * @return {number} Index number of the column
     */
    function _excelColumnToIndex (columnName) {
        var alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
            i,
            j,
            result = 0;

        for (i = 0, j = columnName.length - 1; i < columnName.length; i += 1, j -= 1) {
            result += Math.pow(alpha.length, j) * (alpha.indexOf(columnName[i]) + 1);
        }

        return result;
    };


    /**
     * Parse the row selector to determine start row, end row and row increment
     *
     * Rows can be defined as follows:
     *
     *      number           (single row)           - eg. 7       = row 7
     *      number-number    (from row - to row)    - eg. '4-7'   = rows 4,5,6 and 7
     *      number-          (from row)             - eg. '5-'    = from row 5 to the last row
     *      -number          (to row)               - eg. '-3'    = from row 1 to row 3
     *
     * @param {string|number} row
     * @return {object} {start: rowStart, end: rowEnd, inc: rowIncrement}
     */
    function _parseRowSelector(row) {
        if (typeof row === 'number') {
            return { start: row, end: row, inc: 1 };
        }
        var matches;
        var inc = 1;
        matches = /^(.*)i(\d+)$/.exec(row);
        if (matches) {
            inc = parseInt(matches[2]);
            row = matches[1];
        }
        matches = /^(\d+)$/.exec(row); // match number
        if (matches) {
            return {
                start: parseInt(matches[1]),
                end: parseInt(matches[1]),
                inc: inc,
            };
        }
        matches = /^(\d*)-(\d*)$/.exec(row); // match number range
        if (matches) {
            return {
                start: matches[1] != '' ? parseInt(matches[1]) : 1,
                end: matches[2] != '' ? parseInt(matches[2]) : '',
                inc: inc,
            };
        }
        return { start: 0, end: 0, inc: inc };
    }

    /**
     * Attributes to use when translating the simplified excelStyles object
     * to a format that Excel understands
     *
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
                    parentNode.attr(i, value[i]);
                }
            }
        } else if (value != '') {
            var txAttr = _getTagAttributes(
                tagName,
                attributeName,
                'val',
                value
            );
            parentNode.attr(txAttr);
        }
    };

    /**
     * Add an XML Node to the tree
     *
     * @param {string} type
     * @param {string} attr
     * @param {string|object} value
     * @param {object} parentNode
     */
    var _addXMLNode = function (tagName, attributeName, value, parentNode) {
        var key = _getTranslatedKey(tagName, attributeName);

        var childNode = parentNode
            .append('<' + key + '/>')
            .children()
            .last();

        _addXMLAttribute(tagName, attributeName, value, childNode);
    };

    /**
     * Add Style to the stylesheet
     *
     * @param {object} xlsx
     * @param {object} addStyle Definition of style to add
     */
    var _addXMLStyle = function (xlsx, addStyle) {
        if(addStyle.style === undefined) {
            return -1;
        }
        var xml = xlsx.xl['styles.xml'];
        var cellXfs = $('cellXfs', xml);
        var style = addStyle.style;
        var xf = cellXfs
            .append(
                '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1" applyFill="1" applyBorder="1" />'
            )
            .children()
            .last();

        for (var type in style) {
            var node = $(xml)
                .find(type + 's') // get fonts node
                .append('<' + type + '/>') // append font
                .children()
                .last();
            style[type] = _mergeDefault(type, null, style[type]);
            for (var attr in style[type]) {
                var value = style[type][attr];
                _addXMLNode(type, attr, value, node); // fill, patternFill, object|string, parentNode
            }

            xf.attr(
                type + 'Id',
                $(xml)
                    .find(type + 's')
                    .children().length - 1
            );

            var container = $(type + 's', xml);
            _updateContainerCount(container);
        }
        // Add alignment seperately
        if (addStyle.alignment !== undefined) {
            _addXMLNode('xf', 'alignment', addStyle.alignment, xf);
            xf.attr('applyAlignment', '1');
        }
        _updateContainerCount(cellXfs);
        return cellXfs.children().length - 1;
    };

    /**
     * Update the count attribute on style type containers
     *
     * @param {object} Container node
     */
    var _updateContainerCount = function (container) {
        container.attr('count', container.children().length);
    };

    //return DataTable.Buttons;
});
