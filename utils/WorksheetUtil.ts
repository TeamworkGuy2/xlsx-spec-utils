import CellRefUtil = require("./CellRefUtil");
import CellFormulaValues = require("../../xlsx-spec-models/enums/CellFormulaValues");
import CellValues = require("../../xlsx-spec-models/enums/CellValues");
import SharedStringsUtil = require("./SharedStringsUtil");

/**
 * @author TeamworkGuy2
 * @since 2016-5-28
 */
module WorksheetUtil {

    export interface SimpleCellData {
        val?: string | number;
        vals?: string[];
        cellType: CellValues;
        isFormula?: boolean;
        /** e.g. 'SUM(C2,C8)' or '(G7 + G9) * (G13 - G17)' */
        formulaString?: string;
        /** e.g. 'E3:G12' */
        formulaRange?: string;
        /** an index referencing one of the style IDs in the stylesheet part of this spreadsheet, see StylesheetUtil */
        styleId?: number;
    }


    export function addCalcChainRef(calcChain: OpenXml.CalculationChain, sheetId: number, colRowName: string, childChain?: boolean, newDependencyLevel?: boolean) {
        calcChain.cs.push({
            i: sheetId,
            l: newDependencyLevel,
            r: colRowName,
            s: childChain,
        });
    }


    // TODO support more complex cells
    export function addPlainRow(worksheet: OpenXml.Worksheet, columnVals: (string | number)[], dyDescent: number): OpenXml.Row {
        var res: SimpleCellData[] = [];

        for (var i = 0, size = columnVals.length; i < size; i++) {
            var cellVal = columnVals[i];
            res.push({
                val: cellVal != null ? String(cellVal) : null,
                cellType: getCellValueType(cellVal),
                isFormula: isFormulaString(cellVal),
            });
        }

        return addRow(worksheet, 0, dyDescent, res);
    }


    // add a row to the end/bottom of the spreadsheet
    export function addRow(worksheet: OpenXml.Worksheet, columnOffset: number, dyDescent: number, columnVals: SimpleCellData[]): OpenXml.Row {
        var rowNum = getGreatestRowNum(worksheet.sheetData.rows) + 1;
        return setRow(worksheet, rowNum, columnOffset, dyDescent, columnVals);
    }


    /** Create or overwrite a particular row in a worksheet
     * @param worksheet the worksheet to modify
     * @param rowNum the row number, 1-based
     * @param columnOffset the column offset at which the first cell starts, 0-based
     * @param dyDescent
     * @param columnVals the cell data to place in the row
     */
    export function setRow(worksheet: OpenXml.Worksheet, rowNum: number, columnOffset: number, dyDescent: number, cellVals: SimpleCellData[]): OpenXml.Row {
        if (cellVals.length < 1) { return; }

        var cells: OpenXml.Cell[] = [];
        // create an OpenXml.Cell for each column in the row
        for (var i = 0, size = cellVals.length; i < size; i++) {
            var cell = createCell(rowNum, columnOffset + i, cellVals[i]);
            cells.push(cell);
        }

        // create an OpenXml.Row from the cells
        var newRow: OpenXml.Row = {
            cs: cells,
            r: rowNum,
            spans: CellRefUtil.createCellSpans(columnOffset, cells.length), // 1 based, inclusive
            dyDescent: dyDescent,
        };
        // add the new row to the spreadsheet
        insertOrOverwriteRow(worksheet.sheetData.rows, newRow, true);
        return newRow;
    }


    /** Create a cell based on some simple data
     * @param rowNum the cell's row number, 1-based
     * @param columnIdx the cell's column index, 0-based
     * @param cellData the simple data used to create an 'OpenXml.Cell' object
     */
    export function createCell(rowNum: number, columnIdx: number, cellData: SimpleCellData): OpenXml.Cell {
        var col = CellRefUtil.columnIndexToName(columnIdx);
        var cell: OpenXml.Cell;
        // allow null cells
        if (cellData == null) {
            var cell: OpenXml.Cell = {
                f: null,
                is: null,
                v: null,
                r: col + rowNum,
                t: null,
            };
        }
        else {
            var cellType = cellData.cellType != null ? cellData.cellType.xmlValue : null;
            var isInlineStr = cellType == CellValues.InlineString.xmlValue;
            cell = {
                f: createCellSimpleFormula(cellData),
                is: isInlineStr && cellData.vals != null ? { rs: WorksheetUtil._simpleCellDataToRichTextRuns(cellData.vals) } : null,
                v: !isInlineStr && cellData.val != null ? { content: String(cellData.val) } : null,
                // attributes
                r: col + rowNum,
                s: cellData.styleId,
                t: cellType,
            };
        }
        return cell;
    }


    /** Create or overwrite a particular cell in a worksheet
     * @param worksheet
     * @param columnOffset
     * @param dyDescent
     * @param cellRef
     * @param cellVal
     */
    export function setCell(worksheet: OpenXml.Worksheet, sharedStrings: OpenXml.SharedStringTable, dyDescent: number,
            cellRef: string, cellVal: SimpleCellData): OpenXml.Cell {
        return _mergeOrSetCell(worksheet, sharedStrings, dyDescent, cellRef, cellVal, false, false);
    }


    export function mergeCell(worksheet: OpenXml.Worksheet, sharedStrings: OpenXml.SharedStringTable, dyDescent: number,
            cellRef: string, cellVal: SimpleCellData, overwriteSharedStrings: boolean = false): OpenXml.Cell {
        return _mergeOrSetCell(worksheet, sharedStrings, dyDescent, cellRef, cellVal, true, overwriteSharedStrings);
    }


    export function _mergeOrSetCell(worksheet: OpenXml.Worksheet, sharedStrings: OpenXml.SharedStringTable, dyDescent: number,
            cellRef: string, cellVal: SimpleCellData, mergeIfExisting: boolean, overwriteSharedStrings: boolean): OpenXml.Cell {
        var { col, row } = CellRefUtil.parseCellRef(cellRef);
        var rowIdx = getRowIndex(worksheet.sheetData.rows, row);
        if (rowIdx > -1) {
            var rowData = worksheet.sheetData.rows[rowIdx];
            var cell = createCell(rowIdx + 1, col, cellVal);
            return insertOrOverwriteCell(rowData.cs, cell, undefined, mergeIfExisting, overwriteSharedStrings, sharedStrings);
        }
        else {
            var resRow = setRow(worksheet, row, col, dyDescent, [cellVal]);
            return resRow.cs[0];
        }
    }


    export function updateBounds(ws: OpenXml.Worksheet) {
        var sheetBounds = getLeastAndGreatestRef(ws.sheetData.rows);
        ws.dimension.ref = sheetBounds.min + ":" + sheetBounds.max;
    }


    /** Given a row number, return the index of that row in an array of 'OpenXml.Row' objects.
     * A row number and its index in the 'rows' array may be wildly different since the rows array in a worksheet is sparcely populated (only rows with data exist in the array)
     * @param rows the array of rows to search
     * @param rowNumber the row number, 1-based
     * @return index into the 'rows' array where the specified 'rowNumber' exists, -1 if no match was found
     */
    export function getRowIndex(rows: OpenXml.Row[], rowNumber: number): number {
        for (var i = 0, size = rows.length; i < size; i++) {
            var row = rows[i];
            if (row.r == rowNumber) {
                return i;
            }
        }
        return -1;
    }


    /** Give a cell's column index (i.e. 0 == 'A', 25 == 'Z'), find the index of that cell in an array of 'OpenXml.Cell' objects.
     * A cell's column index may be wildly different from its index in the 'cells' array since the cells array in a worksheet is sparcely populated (only cells with data exist in the array)
     * @param cells the array of cells to search
     * @param cellIdx the column index, 0-based
     * @return index into the 'cells' array where the specified 'cellIdx' exists, -1 if no match was found
     */
    export function getCellIndex(cells: OpenXml.Cell[], cellIdx: number): number {
        for (var i = 0, size = cells.length; i < size; i++) {
            var otherIdx = CellRefUtil.parseCellRefColumn(cells[i].r);
            if (otherIdx == cellIdx) {
                return i;
            }
        }
        return -1;
    }


    /** Get the highest row number reference (i.e. 1 based)
     */
    export function getGreatestRowNum(rows: OpenXml.Row[]): number {
        var lastRow = rows[rows.length - 1];
        if (lastRow) {
            return lastRow.r;
        }
        return 0;
    }


    export function getCellValueType(val: any): CellValues {
        var type = typeof val;
        if (type === "string") {
            return CellValues.String;
        }
        else if (type === "number") {
            return CellValues.Number;
        }
        else if (type === "boolean") {
            return CellValues.Boolean;
        }
        else if (type === "object") {
            if (val != null && typeof val.getTime === "function") {
                return CellValues.Date;
            }
        }
        return val == null ? null : CellValues.Error;
    }


    function createCellSimpleFormula(cell: SimpleCellData): OpenXml.CellFormula {
        if (cell.isFormula) {
            return {
                content: cell.formulaString != null ? cell.formulaString : <string>cell.val,
                ref: cell.formulaRange,
                t: null,
                si: null
            };
        }
        return null;
    }


    function getLeastAndGreatestRef(rows: OpenXml.Row[]): { min: string; max: string; } {
        var rowIdx = getLeastAndGreatestRowIndex(rows);
        var colIdx = getLeastAndGreatestColumnIndex(rows);
        var min = (CellRefUtil.columnIndexToName(colIdx.min) || "A") + (rowIdx.min + 1);
        var max = (CellRefUtil.columnIndexToName(colIdx.max) || "A") + (rowIdx.max + 1);
        return { min, max };
    }


    function getLeastAndGreatestRowIndex(rows: OpenXml.Row[]): { min: number; max: number; } {
        var min = Number.MAX_SAFE_INTEGER;
        var max = Number.MIN_SAFE_INTEGER;
        for (var i = 0, size = rows.length; i < size; i++) {
            var row = rows[i];
            if (row && row.cs && row.cs.length > 0) {
                var rowIdx = row.r - 1;
                if (rowIdx < min) { min = rowIdx; }
                if (rowIdx > max) { max = rowIdx; }
            }
        }
        return { min, max };
    }


    function getLeastAndGreatestColumnIndex(rows: OpenXml.Row[]): { min: number; max: number; } {
        var min = Number.MAX_SAFE_INTEGER;
        var max = Number.MIN_SAFE_INTEGER;
        for (var i = 0, size = rows.length; i < size; i++) {
            var row = rows[i];
            if (row && row.cs && row.cs.length > 0) {
                var firstCell = row.cs[0];
                var lastCell = row.cs[row.cs.length - 1];
                var firstColumnIdx = CellRefUtil.parseCellRefColumn(firstCell.r);
                var lastColumnIdx = CellRefUtil.parseCellRefColumn(lastCell.r);
                if (firstColumnIdx < min) { min = firstColumnIdx; }
                if (lastColumnIdx > max) { max = lastColumnIdx; }
            }
        }
        return { min, max };
    }


    function isFormulaString(val: any) {
        return typeof val === "string" && val.startsWith("=");
    }


    /** Insert a row into an array of rows based on 'r' row number, if a row already exists with the same 'r' row number, overwrite it (only if 'allowOverwrite' = true)
     * @param rows
     * @param newRow the row to add (it's 'r' property is the (1-based) row number to insert or overwrite)
     */
    export function insertOrOverwriteRow(rows: OpenXml.Row[], newRow: OpenXml.Row, allowOverwrite: boolean = false) {
        var rowNum = newRow.r;

        // if an existing row has the same row number, overwrite it (if allowed), return
        var rowIdx = getRowIndex(rows, rowNum);
        if (rowIdx > -1) {
            if (allowOverwrite) {
                rows[rowIdx] = newRow;
            }
            return;
        }

        // if this row has a lower row number than any of the existing cells, insert it at the beginning of the array, return
        if (rowNum < rows[0].r) {
            rows.unshift(newRow);
            return;
        }

        // idx = rows.length, so if no insertion point found, add to end of array
        var idx = rows.length;
        var insert = true;
        // search for a point between two rows where this row number should be
        for (var i = 0, size = rows.length - 1; i < size; i++) {
            if (rows[i].r <= rowNum && rows[i + 1].r > rowNum) {
                if (rows[i].r == rowNum) {
                    idx = i;
                    insert = false;
                }
                else {
                    idx = i + 1;
                    insert = true;
                }
                break;
            }
        }

        // if an insert point was found, insert and shift remaining rows up by one in the array
        if (insert) {
            for (var i = rows.length - 1; i >= idx; i--) {
                rows[i + 1] = rows[i];
            }
            rows[idx] = newRow;
        }
        // else, an overwrite point was found
        else if (allowOverwrite) {
            rows[idx] = newRow;
        }
    }


    /** Insert a cell into an array of cells based on 'r' cell reference, if a cell already exists with the same 'r' cell reference, overwrite it (only if 'allowOverwrite' = true)
     * @param cells
     * @param newCell the cell to add (it's 'r' property is the cell reference to insert into or overwrite)
     * @param allowOverwrite
     */
    export function insertOrOverwriteCell(cells: OpenXml.Cell[], newCell: OpenXml.Cell, allowOverwrite: boolean, allowMerge: boolean, overwriteSharedStrings: boolean, sharedStrings?: OpenXml.SharedStringTable): OpenXml.Cell {
        var parseCol = CellRefUtil.parseCellRefColumn;
        var cellIdx = parseCol(newCell.r);

        // if an existing cell has the same index, overwrite it (if allowed), return
        var findIdx = getCellIndex(cells, cellIdx);
        if (findIdx > -1) {
            if (overwriteSharedStrings) {
                if (sharedStrings == null) { throw new Error("cannot overwrite shared strings without shared string table"); }
                _lookupAndOverwriteSharedStrings(sharedStrings, cells[findIdx], newCell);
            }

            if (allowOverwrite) {
                cells[findIdx] = newCell;
            }
            else if (allowMerge) {
                cells[findIdx] = _mergeCells(cells[findIdx], newCell);
            }
            return cells[findIdx];
        }

        // if this cell has a lower index than any of the existing cells, insert it at the beginning of the array, return
        if (cellIdx < parseCol(cells[0].r)) {
            cells.unshift(newCell);
            return cells[0];
        }

        // idx = cells.length, so if no insertion point found, add to end of array
        var idx = cells.length;
        var insert = true;
        // search for a point between two cells where this cell index should be
        for (var i = 0, size = cells.length - 1; i < size; i++) {
            if (parseCol(cells[i].r) <= cellIdx && parseCol(cells[i + 1].r) > cellIdx) {
                if (parseCol(cells[i].r) == cellIdx) {
                    idx = i;
                    insert = false;
                }
                else {
                    idx = i + 1;
                    insert = true;
                }
                break;
            }
        }

        // if an insert point was found, insert and shift remaining cells up by one in the array
        if (insert) {
            for (var i = cells.length - 1; i >= idx; i--) {
                cells[i + 1] = cells[i];
            }
            cells[idx] = newCell;
            return newCell;
        }
        // else, an overwrite point was found
        else {
            if (overwriteSharedStrings) {
                if (sharedStrings == null) { throw new Error("cannot overwrite shared strings without shared string table"); }
                _lookupAndOverwriteSharedStrings(sharedStrings, cells[findIdx], newCell);
            }

            if (allowOverwrite) {
                cells[idx] = newCell;
            }
            else if (allowMerge) {
                cells[idx] = _mergeCells(cells[idx], newCell);
            }
            return cells[idx];
        }
    }


    /** Merge two cells into one new copy, 'c1' properties take precedence
     * @param c1 base
     * @param c2 takes precendence
     */
    export function _mergeCells(c1: OpenXml.Cell, c2: OpenXml.Cell): OpenXml.Cell {
        var useC2Content = (c2 && c2.v && c2.v.content);
        var useC2InlineStr = (c2 && c2.is && c2.is.rs) || (c2 && c2.is && c2.is.t);

        var res: OpenXml.Cell = {
            cm: c2 && c2.cm ? c2.cm : (c1 ? c1.cm : undefined),
            f: (c2 && c2.f) || (c1 && c1.f) ? {
                content: c2 && c2.f && c2.f.content ? c2.f.content : (c1 && c1.f ? c1.f.content : undefined),
                ref: c2 && c2.f && c2.f.ref ? c2.f.ref : (c1 && c1.f ? c1.f.ref : undefined),
                si: c2 && c2.f && c2.f.si ? c2.f.si : (c1 && c1.f ? c1.f.si : undefined),
                t: c2 && c2.f && c2.f.t ? c2.f.t : (c1 && c1.f ? c1.f.t : undefined),
            } : undefined,
            is: (c2 && c2.is) || (c1 && c1.is) ? {
                rs: useC2InlineStr ? c2.is.rs : (c1 && c1.is ? c1.is.rs : undefined),
                t: useC2InlineStr ? c2.is.t : (c1 && c1.is ? c1.is.t : undefined),
            } : undefined,
            r: c2 && c2.r ? c2.r : (c1 ? c1.r : undefined),
            s: c2 && c2.s ? c2.s : (c1 ? c1.s : undefined),
            t: useC2Content || useC2InlineStr ? c2.t : (c1 ? c1.t : undefined),
            v: (c2 && c2.v) || (c1 && c1.v) ? {
                content: useC2Content ? c2.v.content : (c1 && c1.v ? c1.v.content : undefined),
            } : undefined,
            vm: c2 && c2.vm ? c2.vm : (c1 ? c1.vm : undefined),
        };
        return res;
    }


    /** Given an 'original' and 'new' cell, if the original cell uses shared strings and the new cell uses inline strings, replace the contents of the
     * original cell's shared strings in the shared string table with the new cell's inlined strings and change the new cell's 't', 'v', and 'is' to match.
     * Side effects: newCell is modified, SharedStringTable is modified, any other references to those SharedStringItem indexes now reference new shared strings
     * @param sharedStrings
     * @param origCell
     * @param newCell
     */
    export function _lookupAndOverwriteSharedStrings(sharedStrings: OpenXml.SharedStringTable, origCell: OpenXml.Cell, newCell: OpenXml.Cell): void {
        // if the original cell used shared strings
        if (origCell.v && CellValues.SharedString.xmlValue == origCell.t) {
            var isInlineStr: boolean, isInvalidFormatStr: boolean;
            // if the new cell uses inline strings
            if ((isInlineStr = (newCell.is && CellValues.InlineString.xmlValue == newCell.t)) || (isInvalidFormatStr = (newCell.v && CellValues.String.xmlValue == newCell.t))) {
                var ssIdx = parseInt(origCell.v.content);
                // overwrite the original shared string with the new inline string and use it instead
                var strs = isInlineStr ? SharedStringsUtil.extractText(newCell.is) : (isInvalidFormatStr ? [newCell.v.content] : null);
                SharedStringsUtil.setSharedString(sharedStrings, ssIdx, strs);
                newCell.is = null;
                newCell.v = {
                    content: ssIdx + ''
                };
                newCell.t = CellValues.SharedString.xmlValue;
            }
        }
    }


    /** Convert an array of strings to an array of rich text runs */
    export function _simpleCellDataToRichTextRuns(values: string[], preserveSpace = false): OpenXml.RichTextRun[] {
        var res: OpenXml.RichTextRun[] = [];
        for (var i = 0, size = values.length; i < size; i++) {
            res.push({
                t: {
                    content: values[i],
                    preserveSpace: preserveSpace
                }, rPr: null
            });
        }
        return res;
    }

}

export = WorksheetUtil;