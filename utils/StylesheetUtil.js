"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.StylesheetUtil = void 0;
/** Static functions for finding/creating cell formats, fonts, borders, and other Stylesheet part Open XML elements
 */
var StylesheetUtil;
(function (StylesheetUtil) {
    /** Try to find a cell format matching the given parameters, return the cell format's index if found, -1 if no match
     */
    function findCellFormat(stylesheet, font, numFmt, border, alignment) {
        var fontRes = (typeof font === "number") ? { index: font, apply: true } : font;
        var numFmtRes = (typeof numFmt === "number") ? { id: numFmt, apply: true } : numFmt;
        var borderRes = (typeof border === "number") ? { index: border, apply: true } : border;
        return _findCellFormat(stylesheet, fontRes, numFmtRes, borderRes, alignment);
    }
    StylesheetUtil.findCellFormat = findCellFormat;
    function _findCellFormat(stylesheet, font, numFmt, border, alignment) {
        var _a, _b, _c;
        var cellFormats = (_a = stylesheet.cellXfs) === null || _a === void 0 ? void 0 : _a.xfs;
        var cellStyleFormats = (_b = stylesheet.cellStyleXfs) === null || _b === void 0 ? void 0 : _b.xfs;
        if (cellFormats != null && cellStyleFormats != null) {
            for (var i = 0, size = cellFormats.length; i < size; i++) {
                var fmt = cellFormats[i];
                var style = cellStyleFormats[(_c = fmt === null || fmt === void 0 ? void 0 : fmt.xfId) !== null && _c !== void 0 ? _c : -1];
                // the complex logic and '... || ...Id == 0' is for handling nulls and default values (zero being a default ID)
                if (((font == null) || (((fmt.applyFont == font.apply || font.apply == undefined) && fmt.fontId === font.index) || (style.applyFont == font.apply && style.fontId === font.index))) &&
                    ((numFmt == null) || (((fmt.applyNumberFormat == numFmt.apply || numFmt.apply == undefined || numFmt.id == 0) && fmt.numFmtId === numFmt.id) || ((style.applyNumberFormat == numFmt.apply || numFmt.id == 0) && style.numFmtId === numFmt.id))) &&
                    ((border == null) || (((fmt.applyBorder == border.apply || border.apply == undefined || border.index == 0) && fmt.borderId === border.index) || ((style.applyBorder == border.apply || border.index == 0) && style.borderId === border.index))) &&
                    compareCellFormatAlignment(fmt, style, alignment)) {
                    return i;
                }
            }
        }
        return -1;
    }
    /** Create an Open XML CellFormat object, add it to the stylesheet and return the new cell format's index
     */
    function createCellFormat(stylesheet, font, numFmt, border, alignment) {
        var fontRes = (typeof font === "number") ? { index: font, apply: true } : font;
        var numFmtRes = (typeof numFmt === "number") ? { id: numFmt, apply: true } : numFmt;
        var borderRes = (typeof border === "number") ? { index: border, apply: true } : border;
        return _createCellFormat(stylesheet, fontRes, numFmtRes, borderRes, alignment);
    }
    StylesheetUtil.createCellFormat = createCellFormat;
    function _createCellFormat(stylesheet, font, numFmt, border, alignment) {
        // allow null 'apply' to mean false while still setting the corresponding 'id' or 'index'
        var style = {
            alignment: alignment,
            applyAlignment: alignment != null,
            applyBorder: border != null && ((border.apply === undefined && border.index != null) || border.apply),
            applyFill: false,
            applyFont: font != null && ((font.apply === undefined && font.index != null) || font.apply),
            applyNumberFormat: numFmt != null && ((numFmt.apply === undefined && numFmt.id != null) || numFmt.apply),
            applyProtection: false,
            borderId: border != null && border.index != null ? border.index : 0,
            fillId: 0,
            fontId: font != null && font.index != null ? font.index : 0,
            numFmtId: numFmt != null && numFmt.id != null ? numFmt.id : 0,
            pivotButton: false,
            protection: null,
            quotePrefix: false,
            xfId: 0,
        };
        var cellFormats = stylesheet.cellXfs || (stylesheet.cellXfs = { xfs: [], count: 0 });
        var idx = cellFormats.xfs.push(style) - 1;
        cellFormats.count = cellFormats.xfs.length;
        return idx;
    }
    /** Try to find a CellFormat matching the given parameters, if one cannot be found, create one, return the index of the cell format found or the index of the newly created cell format
     */
    function findOrCreateCellFormat(stylesheet, font, numFmt, border, alignment) {
        var idx = findCellFormat(stylesheet, font, numFmt, border, alignment);
        if (idx < 0) {
            idx = createCellFormat(stylesheet, font, numFmt, border, alignment);
        }
        return idx;
    }
    StylesheetUtil.findOrCreateCellFormat = findOrCreateCellFormat;
    /** Try to find a border matching the given parameters, return the border's index if found, null if no match
     */
    function findBorder(stylesheet, left, right, top, bottom, diagonal) {
        var _a;
        var borders = (_a = stylesheet.borders) === null || _a === void 0 ? void 0 : _a.borders;
        if (borders != null) {
            for (var i = 0, size = borders.length; i < size; i++) {
                var brd = borders[i];
                if (compareBorder(left, brd.left) &&
                    compareBorder(right, brd.right) &&
                    compareBorder(top, brd.top) &&
                    compareBorder(bottom, brd.bottom) &&
                    compareBorder(diagonal, brd.diagonal)) {
                    return i;
                }
            }
        }
        return null;
    }
    StylesheetUtil.findBorder = findBorder;
    /** Create an Open XML Border object, add it to the stylesheet and return the new border's index
     */
    function createBorder(stylesheet, left, right, top, bottom, diagonal) {
        var border = {
            bottom: _createBorder(bottom),
            diagonal: _createBorder(diagonal),
            diagonalDown: false,
            diagonalUp: false,
            end: null,
            horizontal: null,
            left: _createBorder(left),
            outline: false,
            right: _createBorder(right),
            start: null,
            top: _createBorder(top),
            vertical: null,
        };
        var borders = stylesheet.borders || (stylesheet.borders = { borders: [], count: 0 });
        var idx = borders.borders.push(border) - 1;
        borders.count = borders.borders.length;
        return idx;
    }
    StylesheetUtil.createBorder = createBorder;
    /** Try to find a Border matching the given parameters, if one cannot be found, create one, return the index of the border found or the index of the newly created border
     */
    function findOrCreateBorder(stylesheet, left, right, top, bottom, diagonal) {
        var idx = findBorder(stylesheet, left, right, top, bottom, diagonal);
        if (!idx) {
            idx = createBorder(stylesheet, left, right, top, bottom, diagonal);
        }
        return idx;
    }
    StylesheetUtil.findOrCreateBorder = findOrCreateBorder;
    /** Try to find a Font matching the given parameters, return the font's index if found, null if no match
     */
    function findFont(stylesheet, fontSize, colorTheme, fontName, fontFamily, bold, italic, underline) {
        var _a;
        var fonts = (_a = stylesheet.fonts) === null || _a === void 0 ? void 0 : _a.fonts;
        if (fonts != null) {
            for (var i = 0, size = fonts.length; i < size; i++) {
                var fnt = fonts[i];
                if (((fnt.sz && fnt.sz.val == fontSize) || (!fnt.sz && fontSize == null)) &&
                    ((fnt.color && fnt.color.theme == colorTheme) || (!fnt.color && colorTheme == null)) &&
                    ((fnt.name && fnt.name.val == fontName) || (!fnt.name && fontName == null)) &&
                    ((fnt.family && fnt.family.val == fontFamily) || (!fnt.family && fontFamily == null)) &&
                    ((fnt.b && fnt.b.val == bold) || (!fnt.b && (bold == null || bold == false))) &&
                    ((fnt.i && fnt.i.val == italic) || (!fnt.i && (italic == null || italic == false))) &&
                    ((fnt.u && fnt.u.val == underline) || (!fnt.u && underline == null))) {
                    return i;
                }
            }
        }
        return null;
    }
    StylesheetUtil.findFont = findFont;
    /** Create an Open XML Font object, add it to the stylesheet and return the new font's index
     */
    function createFont(stylesheet, fontSize, colorTheme, fontName, fontFamily, bold, italic, underline) {
        var fnt = {
            b: bold == true ? { val: bold } : null,
            charset: null,
            color: colorTheme != null ? { theme: colorTheme } : null,
            condense: null,
            extend: null,
            family: fontFamily ? { val: fontFamily } : null,
            i: italic == true ? { val: italic } : null,
            name: { val: fontName },
            outline: null,
            scheme: null,
            shadow: null,
            strike: null,
            sz: { val: fontSize },
            u: underline != null ? { val: underline } : null,
            vertAlign: null,
        };
        var fonts = stylesheet.fonts || (stylesheet.fonts = { fonts: [], count: 0 });
        var idx = fonts.fonts.push(fnt) - 1;
        fonts.count = fonts.fonts.length;
        return idx;
    }
    StylesheetUtil.createFont = createFont;
    /** Try to find a font matching the given parameters, if one cannot be found, create one, return the index of the font found or the index of the newly created font
     */
    function findOrCreateFont(stylesheet, fontSize, colorTheme, fontName, fontFamily, bold, italic, underline) {
        var idx = findFont(stylesheet, fontSize, colorTheme, fontName, fontFamily, bold, italic, underline);
        if (!idx) {
            idx = createFont(stylesheet, fontSize, colorTheme, fontName, fontFamily, bold, italic, underline);
        }
        return idx;
    }
    StylesheetUtil.findOrCreateFont = findOrCreateFont;
    /** Try to find a NumberingFormat matching the given parameters, return the number format's ID if found, null if no match
     */
    function findNumberFormatId(stylesheet, formatCode) {
        var _a;
        var numFmts = (_a = stylesheet.numFmts) === null || _a === void 0 ? void 0 : _a.numFmts;
        if (numFmts != null) {
            for (var i = 0, size = numFmts.length; i < size; i++) {
                var numFmt = numFmts[i];
                if (numFmt.formatCode == formatCode) {
                    return numFmt.numFmtId;
                }
            }
        }
        return null;
    }
    StylesheetUtil.findNumberFormatId = findNumberFormatId;
    /** Create an Open XML NumberingFormat object, add it to the stylesheet and return the new number format's ID
     */
    function createNumberFormat(stylesheet, formatCode) {
        var fmts = stylesheet.numFmts ? stylesheet.numFmts.numFmts : null;
        // assumption: based on the MSDN Open XML documentation, the highest built-in numFmt ID is ~90
        var highestId = (fmts && fmts.length > 0) ? Math.max(fmts.map(function (nf) { return nf.numFmtId; }).sort(function (a, b) { return b - a; })[0], 100) : 100;
        var numFmt = {
            formatCode: formatCode,
            numFmtId: highestId + 1,
        };
        var numFmts = stylesheet.numFmts || (stylesheet.numFmts = { numFmts: [], count: 0 });
        var idx = numFmts.numFmts.push(numFmt) - 1;
        numFmts.count = numFmts.numFmts.length;
        return idx;
    }
    StylesheetUtil.createNumberFormat = createNumberFormat;
    /** Try to find a number format matching the given parameters, if one cannot be found, create one, return the ID of the number format found or the ID of the newly created number format
     */
    function findOrCreateNumberFormatId(stylesheet, formatCode) {
        var id = findNumberFormatId(stylesheet, formatCode);
        if (!id) {
            id = createNumberFormat(stylesheet, formatCode);
        }
        return id;
    }
    StylesheetUtil.findOrCreateNumberFormatId = findOrCreateNumberFormatId;
    /** Create default 'extList' element found in XLSX spreadsheets
     */
    function createDefaultExtLst(domBldr) {
        return domBldr.create("extLst")
            .addChild(domBldr.create("ext").attrString("uri", "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}").attrString("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
            .addChild(domBldr.create("x14:slicerStyles").attrString("defaultSlicerStyle", "SlicerStyleLight1").element).element)
            .addChild(domBldr.create("ext").attrString("uri", "{9260A510-F301-46a8-8635-F512D64BE5F5}").attrString("xmlns:x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
            .addChild(domBldr.create("x15:timelineStyles").attrString("defaultTimelineStyle", "TimeSlicerStyleLight1").element).element)
            .element;
    }
    StylesheetUtil.createDefaultExtLst = createDefaultExtLst;
    /** Check if either a cell format's alignment or an inherited parent cell format's alignment are equivalent to a given alignment element.
     * @return true if the 'fmt' or 'fmtParent' are equivalent to 'alignment'
     */
    function compareCellFormatAlignment(fmt, fmtParent, alignment) {
        if (alignment == null) {
            return (fmt.alignment == null && fmtParent.alignment == null);
        }
        return (fmt.applyAlignment && fmt.alignment && compareAlignment(alignment, fmt.alignment)) ||
            (fmtParent.applyAlignment && fmtParent.alignment && compareAlignment(alignment, fmtParent.alignment));
    }
    StylesheetUtil.compareCellFormatAlignment = compareCellFormatAlignment;
    /** Check if two OpenXml.Alignment elements are equivalent
     */
    function compareAlignment(a, b) {
        // '==' equality so we don't have to manually check for empty strings, null vs. undefined, etc., may not always compare correctly
        return (a == null && b == null) ||
            (a.horizontal == b.horizontal) &&
                (a.indent == b.indent) &&
                (a.justifyLastLine == b.justifyLastLine) &&
                (a.readingOrder == b.readingOrder) &&
                (a.relativeIndent == b.relativeIndent) &&
                (a.shrinkToFit == b.shrinkToFit) &&
                (a.textRotation == b.textRotation) &&
                (a.vertical == b.vertical) &&
                (a.wrapText == b.wrapText);
    }
    StylesheetUtil.compareAlignment = compareAlignment;
    /** Check if a simple border property is equivalent to an OpenXml.BorderProperty
     */
    function compareBorder(a, b) {
        return b && (a.style == b.style) &&
            (a.auto == (b.color && b.color.auto)) &&
            (a.indexed == (b.color && b.color.indexed)) &&
            (a.rgb == (b.color && b.color.rgb)) &&
            (a.theme == (b.color && b.color.theme)) &&
            (a.tint == (b.color && b.color.tint));
    }
    StylesheetUtil.compareBorder = compareBorder;
    /** Create an OpenXml.BorderProperty from a simple border property
     */
    function _createBorder(borderData) {
        return {
            style: borderData.style,
            color: {
                auto: borderData.auto,
                indexed: borderData.indexed,
                rgb: borderData.rgb,
                theme: borderData.theme,
                tint: borderData.tint,
            },
        };
    }
    StylesheetUtil._createBorder = _createBorder;
})(StylesheetUtil = exports.StylesheetUtil || (exports.StylesheetUtil = {}));
