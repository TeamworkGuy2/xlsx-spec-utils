"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.XlsxFileType = void 0;
/** Information about an XLSX file
 * @author TeamworkGuy2
 * @since 2016-5-27
 * @see OpenXmlIo.XlsxFileType
 */
var XlsxFileType = /** @class */ (function () {
    function XlsxFileType(schemaUrl, contentType, schemaTarget, xlsxFilePath, pathIsTemplate, pathTemplateToken) {
        this.schemaUrl = schemaUrl;
        this.contentType = contentType;
        this.schemaTarget = schemaTarget;
        this.xlsxFilePath = xlsxFilePath;
        this.pathIsTemplate = pathIsTemplate;
        this.pathTemplateToken = pathTemplateToken;
    }
    XlsxFileType.getXmlFilePath = function (sheetNum, info) {
        // TODO the path template token may not be a sheet number, could be a resource identifier (i.e. an image or item prop number)
        return (info.pathIsTemplate ? info.xlsxFilePath.split(info.pathTemplateToken).join(sheetNum) : info.xlsxFilePath);
    };
    return XlsxFileType;
}());
exports.XlsxFileType = XlsxFileType;
