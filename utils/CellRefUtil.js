"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.CellRefUtil = void 0;
/** Utilities for working with cell references (i.e. 'A3' or 'BC26'), column names/indexes, and cell spans (i.e. '3:8' or 'C2:D6')
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
var CellRefUtil;
(function (CellRefUtil) {
    var columnNames = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
    var charCodeA = "A".charCodeAt(0);
    var charCode0 = "0".charCodeAt(0);
    var charCode9 = "9".charCodeAt(0);
    /** Get the column name of a column index (i.e. 0 == 'A', 25 == 'Z')
     * @param i0: the 0-based column index to conver to a column name
     */
    function columnIndexToName(i0) {
        var len = columnNames.length;
        var remaining = i0 + 1;
        var colName = "";
        var modulo;
        while (remaining > 0) {
            modulo = (remaining - 1) % len;
            colName = columnNames[modulo] + colName;
            remaining = Math.floor((remaining - modulo) / len);
        }
        return colName;
    }
    CellRefUtil.columnIndexToName = columnIndexToName;
    /** Get the column index of a column name (i.e. 'A' == 0, 'Z' == 25)
     * @param name: the column name to convert to an index
     */
    function columnNameToIndex(name) {
        var len = columnNames.length;
        var num = 0;
        var pow = 1;
        for (var i = name.length - 1; i >= 0; i--) {
            num += (name.charCodeAt(i) - charCodeA + 1) * pow;
            pow *= len;
        }
        return num - 1;
    }
    CellRefUtil.columnNameToIndex = columnNameToIndex;
    /** Get the column index of a 'A1' formatted cell reference
     * @param ref: the cell reference string in the format 'A1'
     * @return a 0-based 'col' (i.e. 'A' == 0)
     */
    function parseCellRefColumn(ref) {
        var rowDigitOff = 0;
        for (var i = 0, size = ref.length; i < size; i++) {
            var ch = ref.charCodeAt(i);
            if (ch <= charCode9 && ch >= charCode0) {
                rowDigitOff = i;
                break;
            }
        }
        var columnName = ref.substr(0, rowDigitOff);
        var col = columnNameToIndex(columnName);
        return col;
    }
    CellRefUtil.parseCellRefColumn = parseCellRefColumn;
    /** Get the column index and row number of a 'A1' formatted cell reference
     * @param ref: the cell reference string in the format 'A1'
     * @return a 0-based 'col' (i.e. 'A' == 0), and a 1-based row (i.e. 'A1' row == 1)
     */
    function parseCellRef(ref) {
        var rowDigitOff = 0;
        for (var i = 0, size = ref.length; i < size; i++) {
            var ch = ref.charCodeAt(i);
            if (ch <= charCode9 && ch >= charCode0) {
                rowDigitOff = i;
                break;
            }
        }
        var columnName = ref.substr(0, rowDigitOff);
        var col = columnNameToIndex(columnName);
        var row = parseInt(ref.substr(rowDigitOff));
        return { col: col, row: row };
    }
    CellRefUtil.parseCellRef = parseCellRef;
    /** Given a min and max column index (inclusive) and an existing 'spans' row attribute, create a new 'spans' attribute string that represents the merged span of both
     * @param offsetIdx the new column range offset, 0-based
     * @param length the number of columns in the new range
     * @param spansStr the existing 'spans' attribute; i.e. 'spans="3:8"' represents a row containing 6 cells, 3 through 8, inclusive
     */
    function createCellSpans(offsetIdx, length) {
        return '' + (offsetIdx + 1) + ":" + (offsetIdx + length); // 1 based, inclusive
    }
    CellRefUtil.createCellSpans = createCellSpans;
    /** Given a min and max column index (inclusive) and an existing 'spans' row attribute, create a new 'spans' attribute string that represents the merged span of both
     * @param offsetIdx the new column range offset, 0-based
     * @param length the number of columns in the new range
     * @param spansStr the existing 'spans' attribute; i.e. 'spans="3:8"' represents a row containing 6 cells, 3 through 8, inclusive
     */
    function mergeCellSpans(offsetIdx, length, spansStr) {
        var min1 = offsetIdx + 1;
        var max1 = offsetIdx + length;
        var _a = spansStr.split(":"), min2Str = _a[0], max2Str = _a[1];
        var min2 = parseInt(min2Str);
        var max2 = parseInt(max2Str);
        var min = Math.min(min1, min2);
        var max = Math.max(max1, max2);
        return { spans: min + ":" + max, min: min, max: max }; // 1 based, inclusive
    }
    CellRefUtil.mergeCellSpans = mergeCellSpans;
})(CellRefUtil = exports.CellRefUtil || (exports.CellRefUtil = {}));
