/** Utilities for working with cell references (i.e. 'A3' or 'BC26'), column names/indexes, and cell spans (i.e. '3:8' or 'C2:D6')
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
export module CellRefUtil {
    var columnNames = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
    var charCodeA = "A".charCodeAt(0);
    var charCode0 = "0".charCodeAt(0);
    var charCode9 = "9".charCodeAt(0);


    /** Get the column name of a column index (i.e. 0 == 'A', 25 == 'Z')
     * @param i0: the 0-based column index to conver to a column name
     */
    export function columnIndexToName(i0: number): string {
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


    /** Get the column index of a column name (i.e. 'A' == 0, 'Z' == 25)
     * @param name: the column name to convert to an index
     */
    export function columnNameToIndex(name: string): number {
        var len = columnNames.length;
        var num = 0;
        var pow = 1;
        for (var i = name.length - 1; i >= 0; i--) {
            num += (name.charCodeAt(i) - charCodeA + 1) * pow;
            pow *= len;
        }
        return num - 1;
    }


    /** Get the column index of a 'A1' formatted cell reference
     * @param ref: the cell reference string in the format 'A1'
     * @return a 0-based 'col' (i.e. 'A' == 0)
     */
    export function parseCellRefColumn(ref: string): number {
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


    /** Get the column index and row number of a 'A1' formatted cell reference
     * @param ref: the cell reference string in the format 'A1'
     * @return a 0-based 'col' (i.e. 'A' == 0), and a 1-based row (i.e. 'A1' row == 1)
     */
    export function parseCellRef(ref: string): { col: number; row: number } {
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
        return { col, row };
    }


    /** Given a min and max column index (inclusive) and an existing 'spans' row attribute, create a new 'spans' attribute string that represents the merged span of both
     * @param offsetIdx the new column range offset, 0-based
     * @param length the number of columns in the new range
     * @param spansStr the existing 'spans' attribute; i.e. 'spans="3:8"' represents a row containing 6 cells, 3 through 8, inclusive
     */
    export function createCellSpans(offsetIdx: number, length: number): string {
        return '' + (offsetIdx + 1) + ":" + (offsetIdx + length); // 1 based, inclusive
    }


    /** Given a min and max column index (inclusive) and an existing 'spans' row attribute, create a new 'spans' attribute string that represents the merged span of both
     * @param offsetIdx the new column range offset, 0-based
     * @param length the number of columns in the new range
     * @param spansStr the existing 'spans' attribute; i.e. 'spans="3:8"' represents a row containing 6 cells, 3 through 8, inclusive
     */
    export function mergeCellSpans(offsetIdx: number, length: number, spansStr: string): { spans: string; min: number; max: number; } {
        var min1 = offsetIdx + 1;
        var max1 = offsetIdx + length;
        var [min2Str, max2Str] = spansStr.split(":");
        var min2 = parseInt(min2Str);
        var max2 = parseInt(max2Str);
        var min = Math.min(min1, min2);
        var max = Math.max(max1, max2);
        return { spans: min + ":" + max, min, max }; // 1 based, inclusive
    }

}