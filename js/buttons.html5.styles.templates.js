/**
 * Style templates for html5.styles
 *
 * @version: 0.1
 * @description Easy templats for 'excelStyles'
 * @file buttons.html5.styles.templates.js
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
     * Override the html5.styles.js applyStyles function to initialize the templates
     */
    DataTable.ext.buttons.excelHtml5.applyStyles = function (xlsx) {
        var excelStyles = this.exportOptions.excelStyles;
        if (excelStyles !== undefined) {
            if (!Array.isArray(excelStyles)) {
                excelStyles = [excelStyles];
            }
            excelStyles = _replaceTemplatesWithStyles(excelStyles);
            this._applyExcelStyles(xlsx, excelStyles);
        }
    };

    /**
     * Replace any template names found in the styles with the template style content
     * 
     * @param {array} excelStyles The excel styles to apply
     */
    var _replaceTemplatesWithStyles = function (excelStyles) {
        var templatesLoaded = false;
        if (DataTable.ext.buttons.excelHtml5.getTemplate !== undefined) {
            templatesLoaded = true;
        }
        var newStyles = [];
        for (var i in excelStyles) {
            if (excelStyles[i].template !== undefined) {
                var templateName = excelStyles[i].template;
                if (templatesLoaded) {
                    var template = DataTable.ext.buttons.excelHtml5.getTemplate(
                        templateName
                    );
                    if (template !== false) {
                        for (var j in template.excelStyles) {
                            newStyles.push(template.excelStyles[j]);
                        }
                    } else {
                        console.log(
                            "Error: Template '" +
                                templateName +
                                "' not found. Ignoring template."
                        );
                    }
                } else {
                    console.log(
                        "Error: the style.templates.js library has not been included - template '" +
                            templateName +
                            "' ignored"
                    );
                }
            } else {
                newStyles.push(excelStyles[i]);
            }
        }
        return newStyles;
    };

    DataTable.ext.buttons.excelHtml5.getTemplate = function(templateName) {
        return _templates[templateName] || false;
    }

    var _templates = {
        blue_medium: {
            description: 'Blue Medium Weight',
            excelStyles: [
                {
                    cells: '2',
                    style: {
                        font: {
                            color: 'FFFFFF',
                        },
                        fill: {
                            patternFill: {
                                patternType: 'solid',
                                fgColor: '4472C4',
                                bgColor: '4472C4',
                            },
                        },
                    },
                },
                {
                    cells: '3:n,2',
                    style: {
                        fill: {
                            patternFill: {
                                patternType: 'solid',
                                fgColor: 'D9E1F2',
                                bgColor: 'D9E1F2',
                            },
                        },
                    },
                },
                {
                    cells: '2:',
                    style: {
                        border: {
                            top: {
                                style: 'thin',
                                color: '8EA9DB',
                            },
                            bottom: {
                                style: 'thin',
                                color: '8EA9DB',
                            },
                        },
                    },
                },
            ],
        },
        green_medium: {
            description: 'Green Medium Weight',
            excelStyles: [
                {
                    cells: '2',
                    style: {
                        font: {
                            color: 'FFFFFF',
                        },
                        fill: {
                            patternFill: {
                                patternType: 'solid',
                                fgColor: '70AD47',
                                bgColor: '70AD47',
                            },
                        },
                    },
                },
                {
                    cells: '3:n,2',
                    style: {
                        fill: {
                            patternFill: {
                                patternType: 'solid',
                                fgColor: 'E2EFDA',
                                bgColor: 'E2EFDA',
                            },
                        },
                    },
                },
                {
                    cells: '2:',
                    style: {
                        border: {
                            top: {
                                style: 'thin',
                                color: 'A9D08E',
                            },
                            bottom: {
                                style: 'thin',
                                color: 'A9D08E',
                            },
                        },
                    },
                },
            ],
        },
    };

    return DataTable.Buttons;
});


