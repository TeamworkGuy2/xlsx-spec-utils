"use strict";
/** Information about an XLSX file
 * @author TeamworkGuy2
 * @since 2016-5-27
 * @see OpenXmlIo.XlsxFileType
 */
var XlsxFileTypeImpl = (function () {
    /**
     * @param schemaUrl: the URL of this file's XML DTD schema
     * @param contentType: the content/mime type name of this file
     * @param schemaTarget: the 'target' attribute for this file type used in XLSX files
     * @param xlsxFilePath: the relative path inside an unzipped XLSX file where this file resides (the path can be a template string that needs a specific sheet number or resource identifier to complete)
     * @param pathIsTemplate: whether the 'xlsxFilePath' is a template
     * @param pathTemplateToken: the template token/string to replace in 'xslxFilePath' with a sheet number or resource identifier to make it a valid path
     */
    function XlsxFileTypeImpl(schemaUrl, contentType, schemaTarget, xlsxFilePath, pathIsTemplate, pathTemplateToken) {
        this.schemaUrl = schemaUrl;
        this.contentType = contentType;
        this.schemaTarget = schemaTarget;
        this.xlsxFilePath = xlsxFilePath;
        this.pathIsTemplate = pathIsTemplate;
        this.pathTemplateToken = pathTemplateToken;
    }
    return XlsxFileTypeImpl;
}());
module.exports = XlsxFileTypeImpl;
