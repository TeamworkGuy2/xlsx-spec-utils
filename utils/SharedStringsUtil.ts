import SharedStringItem = require("../../xlsx-spec-models/types/SharedStringItem");

/** Functions for working with 'OpenXml.SharedStringTable'.  Find, add, and overwrite workbook shared strings.
 * @since 2016-3-31
 */
module SharedStringsUtil {


    /** Try to find a shared string matching the given parameters, return the shared string's index if found, -1 if no match
     */
    export function findSharedString(sharedStrings: OpenXml.SharedStringTable, str: string, preserveSpace?: boolean): number {
        var sis = sharedStrings.sis;
        for (var i = 0, size = sis.length; i < size; i++) {
            var ss = sis[i];
            if (ss.t && ss.t.content === str &&
                    (preserveSpace == null || ss.t.preserveSpace == preserveSpace)) {
                return i;
            }
        }
        return -1;
    }


    /** Create an Open XML SharedStringItem object, add it to the shared strings and return the new shared string's index
     */
    export function createSharedString(sharedStrings: OpenXml.SharedStringTable, str: string, preserveSpace?: boolean): number {
        return addSharedString(sharedStrings, {
            t: {
                content: str,
                preserveSpace: preserveSpace,
            }
        });
    }


    /** Try to find a SharedStringItem matching the given parameters, if one cannot be found, create one, return the index of the shared string found or the index of the newly created shared string
     */
    export function findOrCreateSharedString(sharedStrings: OpenXml.SharedStringTable, str: string, preserveSpace?: boolean): number {
        var idx = findSharedString(sharedStrings, str, preserveSpace);
        if (idx < 0) {
            idx = createSharedString(sharedStrings, str, preserveSpace);
        }
        return idx;
    }


    export function addSharedString(sharedStrings: OpenXml.SharedStringTable, sharedStrItem: OpenXml.SharedStringItem, copy: boolean = true) {
        sharedStrItem = copy ? SharedStringItem.copy(sharedStrItem) : sharedStrItem;
        sharedStrings.sis.push(sharedStrItem);
        return sharedStrings.sis.length - 1;
    }


    /** Given a shared string table, an index, and one or more strings, find the shared string at 'idx' in the shared string table and if it is a plain string, set it's value to 'strs[0]',
     * else the shared string contains multiple sub-strings, loop over each and set each equal to 'strs[i]'
     * @param sharedStrings the 'OpenXml.SharedStringTable' to search
     * @param idx the index of the shared string to modify
     * @param strs the list of strings to use to set the shared string or its sub-strings
     */
    export function setSharedString(sharedStrings: OpenXml.SharedStringTable, idx: number, strs: string[]) {
        var sharedStr = sharedStrings.sis[idx];
        if (!sharedStr) {
            throw new Error("could not find shared string '" + idx + "' there are " + sharedStrings.sis.length + " shared strings");
        }

        if (sharedStr.t) {
            sharedStr.t.content = strs[0];
        }
        else {
            var richStrs = sharedStr.rs;
            for (var i = 0, size = richStrs.length; i < size; i++) {
                richStrs[i].t.content = strs[i];
            }
        }
    }


    /** Create a copy of an existing shared string instance
     * @param sharedStrings
     * @param idx
     * @return the new copy's index in the shared string table
     */
    export function copySharedString(sharedStrings: OpenXml.SharedStringTable, idx: number): number {
        var sharedStr = sharedStrings.sis[idx];
        if (!sharedStr) {
            throw new Error("could not find shared string '" + idx + "' there are " + sharedStrings.sis.length + " shared strings");
        }

        return addSharedString(sharedStrings, sharedStr, true);
    }


    /** Copies and sets a shared string
     * @see copySharedString()
     * @see setSharedString()
     * @param sharedStrings
     * @param idx
     * @param strs
     * @return the new copy's index in the shared string table
     */
    export function copyAndSetSharedString(sharedStrings: OpenXml.SharedStringTable, idx: number, strs: string[]): number {
        var newIdx = copySharedString(sharedStrings, idx);
        setSharedString(sharedStrings, newIdx, strs);
        return newIdx;
    }


    /** Returns an array with a single string containing contents of 't' (Text) or an array with strings from each element in 'rs' (Rich Text Run)
     * @param sharedStr the SharedStringItem or InlineString to extract the 'content' strings from
     */
    export function extractText(sharedStr: OpenXml.SharedStringItem | OpenXml.InlineString): string[] {
        if (sharedStr.t) {
            return [sharedStr.t.content];
        }
        else {
            var res: string[] = [];
            var richStrs = sharedStr.rs;
            for (var i = 0, size = richStrs.length; i < size; i++) {
                res.push(richStrs[i].t.content);
            }
            return res;
        }
    }

}

export = SharedStringsUtil;