/** Static functions for finding/creating cell formats, fonts, borders, and other Stylesheet part Open XML elements
 */
module StylesheetUtil {

    interface SimpleBorderProperty {
        indexed?: number;
        style?: string;
        rgb?: string | number;
        theme?: number;
        tint?: number;
        auto?: boolean;
    }


    interface IndexFlag {
        index: number;
        apply?: boolean;
    }


    interface IdFlag {
        id: number;
        apply?: boolean;
    }


    /** Try to find a cell format matching the given parameters, return the cell format's index if found, -1 if no match
     */
    export function findCellFormat(stylesheet: OpenXml.Stylesheet, font: number | IndexFlag, numFmt: number | IdFlag, border: number | IndexFlag, alignment: OpenXml.Alignment): number {
        var fontRes = (typeof font === "number") ? { index: font, apply: true } : font;
        var numFmtRes = (typeof numFmt === "number") ? { id: numFmt, apply: true } : numFmt;
        var borderRes = (typeof border === "number") ? { index: border, apply: true } : border;
        return _findCellFormat(stylesheet, fontRes, numFmtRes, borderRes, alignment);
    }


    function _findCellFormat(stylesheet: OpenXml.Stylesheet, font: IndexFlag, numFmt: IdFlag, border: IndexFlag, alignment: OpenXml.Alignment): number {
        var cellFormats = stylesheet.cellXfs?.xfs;
        var cellStyleFormats = stylesheet.cellStyleXfs?.xfs;
        if (cellFormats != null && cellStyleFormats != null) {
            for (var i = 0, size = cellFormats.length; i < size; i++) {
                var fmt = cellFormats[i];
                var style = cellStyleFormats[fmt?.xfId ?? -1];
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
    export function createCellFormat(stylesheet: OpenXml.Stylesheet, font: number | IndexFlag, numFmt: number | IdFlag, border: number | IndexFlag, alignment: OpenXml.Alignment): number {
        var fontRes = (typeof font === "number") ? { index: font, apply: true } : font;
        var numFmtRes = (typeof numFmt === "number") ? { id: numFmt, apply: true } : numFmt;
        var borderRes = (typeof border === "number") ?  { index: border, apply: true } : border;
        return _createCellFormat(stylesheet, fontRes, numFmtRes, borderRes, alignment);
    }


    function _createCellFormat(stylesheet: OpenXml.Stylesheet, font: IndexFlag, numFmt: IdFlag, border: IndexFlag, alignment: OpenXml.Alignment): number {
        // allow null 'apply' to mean false while still setting the corresponding 'id' or 'index'
        var style: OpenXml.CellFormat = {
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
    export function findOrCreateCellFormat(stylesheet: OpenXml.Stylesheet, font: number | IndexFlag, numFmt: number | IdFlag, border: number | IndexFlag, alignment: OpenXml.Alignment): number {
        var idx = findCellFormat(stylesheet, font, numFmt, border, alignment);
        if (idx < 0) {
            idx = createCellFormat(stylesheet, font, numFmt, border, alignment);
        }
        return idx;
    }


    /** Try to find a border matching the given parameters, return the border's index if found, null if no match
     */
    export function findBorder(stylesheet: OpenXml.Stylesheet, left: SimpleBorderProperty, right: SimpleBorderProperty, top: SimpleBorderProperty, bottom: SimpleBorderProperty, diagonal: SimpleBorderProperty): number | null {
        var borders = stylesheet.borders?.borders;
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


    /** Create an Open XML Border object, add it to the stylesheet and return the new border's index
     */
    export function createBorder(stylesheet: OpenXml.Stylesheet, left: SimpleBorderProperty, right: SimpleBorderProperty, top: SimpleBorderProperty, bottom: SimpleBorderProperty, diagonal: SimpleBorderProperty): number {
        var border: OpenXml.Border = {
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


    /** Try to find a Border matching the given parameters, if one cannot be found, create one, return the index of the border found or the index of the newly created border
     */
    export function findOrCreateBorder(stylesheet: OpenXml.Stylesheet, left: SimpleBorderProperty, right: SimpleBorderProperty, top: SimpleBorderProperty, bottom: SimpleBorderProperty, diagonal: SimpleBorderProperty): number {
        var idx = findBorder(stylesheet, left, right, top, bottom, diagonal);
        if (!idx) {
            idx = createBorder(stylesheet, left, right, top, bottom, diagonal);
        }
        return idx;
    }


    /** Try to find a Font matching the given parameters, return the font's index if found, null if no match
     */
    export function findFont(stylesheet: OpenXml.Stylesheet, fontSize: number, colorTheme: number, fontName: string, fontFamily: number, bold: boolean, italic: boolean, underline: string): number | null {
        var fonts = stylesheet.fonts?.fonts;
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


    /** Create an Open XML Font object, add it to the stylesheet and return the new font's index
     */
    export function createFont(stylesheet: OpenXml.Stylesheet, fontSize: number, colorTheme: number, fontName: string, fontFamily?: number, bold?: boolean, italic?: boolean, underline?: string): number {
        var fnt: OpenXml.Font = {
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
            u: underline != null ? { val: <OpenXml.ST_UnderlineValues>underline } : null,
            vertAlign: null,
        };
        var fonts = stylesheet.fonts || (stylesheet.fonts = { fonts: [], count: 0 });
        var idx = fonts.fonts.push(fnt) - 1;
        fonts.count = fonts.fonts.length;
        return idx;
    }


    /** Try to find a font matching the given parameters, if one cannot be found, create one, return the index of the font found or the index of the newly created font
     */
    export function findOrCreateFont(stylesheet: OpenXml.Stylesheet, fontSize: number, colorTheme: number, fontName: string, fontFamily: number, bold: boolean, italic: boolean, underline: string): number {
        var idx = findFont(stylesheet, fontSize, colorTheme, fontName, fontFamily, bold, italic, underline);
        if (!idx) {
            idx = createFont(stylesheet, fontSize, colorTheme, fontName, fontFamily, bold, italic, underline);
        }
        return idx;
    }



    /** Try to find a NumberingFormat matching the given parameters, return the number format's ID if found, null if no match
     */
    export function findNumberFormatId(stylesheet: OpenXml.Stylesheet, formatCode: string): number | null {
        var numFmts = stylesheet.numFmts?.numFmts;
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


    /** Create an Open XML NumberingFormat object, add it to the stylesheet and return the new number format's ID
     */
    export function createNumberFormat(stylesheet: OpenXml.Stylesheet, formatCode: string): number {
        var fmts = stylesheet.numFmts ? stylesheet.numFmts.numFmts : null;
        // assumption: based on the MSDN Open XML documentation, the highest built-in numFmt ID is ~90
        var highestId = (fmts && fmts.length > 0) ? Math.max(fmts.map((nf) => nf.numFmtId).sort((a, b) => b - a)[0], 100) : 100;

        var numFmt: OpenXml.NumberingFormat = {
            formatCode: formatCode,
            numFmtId: highestId + 1,
        };
        var numFmts = stylesheet.numFmts || (stylesheet.numFmts = { numFmts: [], count: 0 });
        var idx = numFmts.numFmts.push(numFmt) - 1;
        numFmts.count = numFmts.numFmts.length;
        return idx;
    }


    /** Try to find a number format matching the given parameters, if one cannot be found, create one, return the ID of the number format found or the ID of the newly created number format
     */
    export function findOrCreateNumberFormatId(stylesheet: OpenXml.Stylesheet, formatCode: string): number {
        var id = findNumberFormatId(stylesheet, formatCode);
        if (!id) {
            id = createNumberFormat(stylesheet, formatCode);
        }
        return id;
    }


    /** Create default 'extList' element found in XLSX spreadsheets
     */
    export function createDefaultExtLst(domBldr: DomBuilderFactory): ElementLike {
        return domBldr.create("extLst")
            .addChild(domBldr.create("ext").attrString("uri", "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}").attrString("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
                .addChild(domBldr.create("x14:slicerStyles").attrString("defaultSlicerStyle", "SlicerStyleLight1").element).element)
            .addChild(domBldr.create("ext").attrString("uri", "{9260A510-F301-46a8-8635-F512D64BE5F5}").attrString("xmlns:x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
                .addChild(domBldr.create("x15:timelineStyles").attrString("defaultTimelineStyle", "TimeSlicerStyleLight1").element).element)
            .element;
    }


    /** Check if either a cell format's alignment or an inherited parent cell format's alignment are equivalent to a given alignment element.
     * @return true if the 'fmt' or 'fmtParent' are equivalent to 'alignment'
     */
    export function compareCellFormatAlignment(fmt: OpenXml.CellFormat, fmtParent: OpenXml.CellFormat, alignment: OpenXml.Alignment) {
        if (alignment == null) {
            return (fmt.alignment == null && fmtParent.alignment == null);
        }
        return (fmt.applyAlignment && fmt.alignment && compareAlignment(alignment, fmt.alignment)) ||
                (fmtParent.applyAlignment && fmtParent.alignment && compareAlignment(alignment, fmtParent.alignment));
    }


    /** Check if two OpenXml.Alignment elements are equivalent
     */
    export function compareAlignment(a: OpenXml.Alignment, b: OpenXml.Alignment) {
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


    /** Check if a simple border property is equivalent to an OpenXml.BorderProperty
     */
    export function compareBorder(a: SimpleBorderProperty, b: OpenXml.BorderProperty | null | undefined) {
        return b && (a.style == b.style) &&
            (a.auto == (b.color && b.color.auto)) &&
            (a.indexed == (b.color && b.color.indexed)) &&
            (a.rgb == (b.color && b.color.rgb)) &&
            (a.theme == (b.color && b.color.theme)) &&
            (a.tint == (b.color && b.color.tint));
    }


    /** Create an OpenXml.BorderProperty from a simple border property
     */
    export function _createBorder(borderData: SimpleBorderProperty): OpenXml.BorderProperty {
        return {
            style: <OpenXml.ST_BorderStyle | null>borderData.style,
            color: {
                auto: borderData.auto,
                indexed: borderData.indexed,
                rgb: borderData.rgb,
                theme: borderData.theme,
                tint: borderData.tint,
            },
        };
    }

}

export = StylesheetUtil;