/**
 * Style templates for html5.styles
 *
 * @version: 0.7.0
 * @description Easy templates for 'excelStyles'
 * @file buttons.html5.styles.templates.js
 * @copyright © 2020 Beyond the Box Creative
 * @author Paul Jones <info@pauljones.co.nz>
 * @license MIT
 *
 * Include this file after including the buttons.html5.styles.js (along with the required DataTables dependencies)
 *
 * @todo Add extra templates really soon
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
        var excelStyles = this.excelStyles || this.exportOptions.excelStyles;
        if (excelStyles !== undefined) {
            if (!Array.isArray(excelStyles)) {
                excelStyles = [excelStyles];
            }
            this.excelStyles = _replaceTemplatesWithStyles(
                excelStyles
            );
            this._applyExcelStyles(xlsx);
        }
    };

    /**
     * Replace any template names found in the styles with the template style content
     *
     * @param {array} excelStyles The excel styles to apply
     */
    var _replaceTemplatesWithStyles = function (excelStyles) {
        var newStyles = [];
        for (var i in excelStyles) {
            if (excelStyles[i].template !== undefined) {
                var templateName = excelStyles[i].template;
                var template = _getTemplate(templateName);
                if (template !== false) {
                    if (Array.isArray(template.es)) {
                        for (var j in template.es) {
                            if (excelStyles[i].cells !== undefined) {
                                template.es[j].cells = excelStyles[i].cells;
                            }
                            newStyles.push(template.es[j]);
                        }
                    }
                    else {
                        if (excelStyles[i].cells !== undefined) {
                            template.es.cells = excelStyles[i].cells;
                        }
                        newStyles.push(template.es);
                    }
                } else {
                    console.log(
                        "Error: Template '" +
                            templateName +
                            "' not found. Ignoring template."
                    );
                }
            } else {
                newStyles.push(excelStyles[i]);
            }
        }
        return newStyles;
    };

    var _getTemplate = function (templateName) {
        return _templates[templateName] || false;
    };

    /**
     * Template parts to be used to create excelStyles, and also be available as styles in themselves
     * Note: excelStyles key shortened to es for brevity
     */
    var _tp = {
        header_blue: {
            es: {
                cells: ['sh', 'sf'],
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
        },
        header_green: {
            es: {
                cells: ['sh', 'sf'],
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
        },
        stripes_blue: {
            es: {
                cells: 's1:n,2',
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
        },
        stripes_green: {
            es: {
                cells: 's1:n,2',
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
        },
        rowlines_blue: {
            es: {
                cells: 'sh:f',
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
        },
        rowlines_green: {
            es: {
                cells: 'sh:f',
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
        },
        currency_us: {
            es: {
                style: {
                    numFmt: '[$$-en-US] #,##0.00',
                },
            },
        },
        currency_eu: {
            es: {
                style: {
                    numFmt: '[$€-x-euro2] #,##0.00',
                },
            },
        },
        currency_gb: {
            es: {
                style: {
                    numFmt: '[$£-en-GB]#,##0.00',
                },
            },
        },
        int: {
            es: {
                style: {
                    numFmt: '#,##0;(#,##0)',
                },
            },
        },
        decimal_1: {
            es: {
                style: {
                    numFmt: '#,##0.0;(#,##0.0)',
                },
            },
        },
        decimal_2: {
            es: {
                style: {
                    numFmt: '#,##0.00;(#,##0.00)',
                },
            },
        },
        decimal_3: {
            es: {
                style: {
                    numFmt: '#,##0.000;(#,##0.000)',
                },
            },
        },
        decimal_4: {
            es: {
                style: {
                    numFmt: '#,##0.0000;(#,##0.0000)',
                },
            },
        },
    };

    /**
     * Templates available for styles
     */
    var _templates = {
        blue_medium: {
            desc: 'Blue Medium Weight',
            es: [
                _tp.header_blue.es,
                _tp.stripes_blue.es,
                _tp.rowlines_blue.es,
            ],
        },
        green_medium: {
            desc: 'Green Medium Weight',
            es: [
                _tp.header_green.es,
                _tp.stripes_green.es,
                _tp.rowlines_green.es,
            ],
        },
    };

    $.extend(_templates, _tp);

    return DataTable.Buttons;
});
