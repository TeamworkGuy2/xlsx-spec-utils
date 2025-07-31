"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.xlsxAdditionalNamespaces = exports.openxmlNamespaces = void 0;
/** A map of Open XML namespace prefixes to their schema URIs.
 */
exports.openxmlNamespaces = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
};
/** A map of additional XLSX namespace prefixes to their schema URIs.
 * These namespaces are ignored in OpenXML files using a root level 'mc:Ignorable' attribute.
 */
exports.xlsxAdditionalNamespaces = {
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x15ac": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac",
    "x16r2": "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
    "mx": "http://schemas.microsoft.com/office/mac/excel/2008/main",
    "xcalcf": "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "xr6": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
    "xr10": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
};
