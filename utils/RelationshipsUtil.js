"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.RelationshipsUtil = void 0;
var XlsxFileType_1 = require("../files/XlsxFileType");
/** Utilities for working with ".rels" files
 * @author TeamworkGuy2
 * @since 2017-07-02
 */
var RelationshipsUtil;
(function (RelationshipsUtil) {
    var _id = 0;
    function uniqueId(space) {
        var id = ++_id;
        return space + id;
    }
    RelationshipsUtil.uniqueId = uniqueId;
    function createBaseRels(sheetNum, files) {
        var relationships = [];
        for (var i = 0, size = files.length; i < size; i++) {
            relationships.push({
                id: "rId" + (i + 1),
                target: files[i].schemaUrl,
                type: XlsxFileType_1.XlsxFileType.getXmlFilePath(sheetNum, files[i])
            });
        }
        return { relationships: relationships };
    }
    RelationshipsUtil.createBaseRels = createBaseRels;
})(RelationshipsUtil = exports.RelationshipsUtil || (exports.RelationshipsUtil = {}));
