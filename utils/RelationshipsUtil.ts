import XlsxFileType = require("../files/XlsxFileType");

/** Utilities for working with ".rels" files
 * @author TeamworkGuy2
 * @since 2017-07-02
 */
module RelationshipsUtil {
    var _id = 0;


    export function uniqueId(space: string) {
        var id = ++_id;
        return space + id;
    }


    export function createBaseRels(sheetNum: number, files: OpenXmlIo.XlsxFileType[]): OpenXml.Relationships {
        var relationships: OpenXml.Relationship[] = [];
        for (var i = 0, size = files.length; i < size; i++) {
            relationships.push({
                id: "rId" + (i + 1),
                target: files[i].schemaUrl,
                type: XlsxFileType.getXmlFilePath(sheetNum, files[i])
            });
        }
        return { relationships };
    }

}

export = RelationshipsUtil;