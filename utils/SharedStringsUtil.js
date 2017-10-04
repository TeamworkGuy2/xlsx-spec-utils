"use strict";
var SharedStringTable = require("../../xlsx-spec-models/root-types/SharedStringTable");
/** Functions for working with 'OpenXml.SharedStringTable'.  Find, add, and overwrite workbook shared strings.
 * @since 2016-3-31
 */
var SharedStringsUtil;
(function (SharedStringsUtil) {
    /** Try to find a shared string matching the given parameters, return the shared string's index if found, -1 if no match
     */
    function findSharedString(sharedStrings, str, preserveSpace) {
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
    SharedStringsUtil.findSharedString = findSharedString;
    /** Create an Open XML SharedStringItem object, add it to the shared strings and return the new shared string's index
     */
    function createSharedString(sharedStrings, str, preserveSpace) {
        return addSharedString(sharedStrings, {
            t: {
                content: str,
                preserveSpace: preserveSpace,
            }
        });
    }
    SharedStringsUtil.createSharedString = createSharedString;
    /** Try to find a SharedStringItem matching the given parameters, if one cannot be found, create one, return the index of the shared string found or the index of the newly created shared string
     */
    function findOrCreateSharedString(sharedStrings, str, preserveSpace) {
        var idx = findSharedString(sharedStrings, str, preserveSpace);
        if (idx < 0) {
            idx = createSharedString(sharedStrings, str, preserveSpace);
        }
        return idx;
    }
    SharedStringsUtil.findOrCreateSharedString = findOrCreateSharedString;
    function addSharedString(sharedStrings, sharedStrItem, copy) {
        if (copy === void 0) { copy = true; }
        sharedStrItem = copy ? SharedStringTable.SharedStringItem.copy(sharedStrItem) : sharedStrItem;
        sharedStrings.sis.push(sharedStrItem);
        return sharedStrings.sis.length - 1;
    }
    SharedStringsUtil.addSharedString = addSharedString;
    /** Given a shared string table, an index, and one or more strings, find the shared string at 'idx' in the shared string table and if it is a plain string, set it's value to 'strs[0]',
     * else the shared string contains multiple sub-strings, loop over each and set each equal to 'strs[i]'
     * @param sharedStrings the 'OpenXml.SharedStringTable' to search
     * @param idx the index of the shared string to modify
     * @param strs the list of strings to use to set the shared string or its sub-strings
     */
    function setSharedString(sharedStrings, idx, strs) {
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
    SharedStringsUtil.setSharedString = setSharedString;
    /** Create a copy of an existing shared string instance
     * @param sharedStrings
     * @param idx
     * @return the new copy's index in the shared string table
     */
    function copySharedString(sharedStrings, idx) {
        var sharedStr = sharedStrings.sis[idx];
        if (!sharedStr) {
            throw new Error("could not find shared string '" + idx + "' there are " + sharedStrings.sis.length + " shared strings");
        }
        return addSharedString(sharedStrings, sharedStr, true);
    }
    SharedStringsUtil.copySharedString = copySharedString;
    /** Copies and sets a shared string
     * @see copySharedString()
     * @see setSharedString()
     * @param sharedStrings
     * @param idx
     * @param strs
     * @return the new copy's index in the shared string table
     */
    function copyAndSetSharedString(sharedStrings, idx, strs) {
        var newIdx = copySharedString(sharedStrings, idx);
        setSharedString(sharedStrings, newIdx, strs);
        return newIdx;
    }
    SharedStringsUtil.copyAndSetSharedString = copyAndSetSharedString;
    /** Returns an array with a single string containing contents of 't' (Text) or an array with strings from each element in 'rs' (Rich Text Run)
     * @param sharedStr the SharedStringItem or InlineString to extract the 'content' strings from
     */
    function extractText(sharedStr) {
        if (sharedStr.t) {
            return [sharedStr.t.content];
        }
        else {
            var res = [];
            var richStrs = sharedStr.rs;
            for (var i = 0, size = richStrs.length; i < size; i++) {
                res.push(richStrs[i].t.content);
            }
            return res;
        }
    }
    SharedStringsUtil.extractText = extractText;
})(SharedStringsUtil || (SharedStringsUtil = {}));
module.exports = SharedStringsUtil;
