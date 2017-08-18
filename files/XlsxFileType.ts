/** Information about an XLSX file
 * @author TeamworkGuy2
 * @since 2016-5-27
 * @see OpenXmlIo.XlsxFileType
 */
class XlsxFileTypeImpl implements OpenXmlIo.XlsxFileType {
    public schemaUrl: string;
    public schemaTarget: string;
    public contentType: string;
    public xlsxFilePath: string;
    public pathIsTemplate: boolean;
    public pathTemplateToken: string;


    /**
     * @param schemaUrl: the URL of this file's XML DTD schema
     * @param contentType: the content/mime type name of this file
     * @param schemaTarget: the 'target' attribute for this file type used in XLSX files
     * @param xlsxFilePath: the relative path inside an unzipped XLSX file where this file resides (the path can be a template string that needs a specific sheet number or resource identifier to complete)
     * @param pathIsTemplate: whether the 'xlsxFilePath' is a template
     * @param pathTemplateToken: the template token/string to replace in 'xslxFilePath' with a sheet number or resource identifier to make it a valid path
     */
    constructor(schemaUrl: string, contentType: string, schemaTarget: string, xlsxFilePath: string, pathIsTemplate: boolean, pathTemplateToken: string) {
        this.schemaUrl = schemaUrl;
        this.contentType = contentType;
        this.schemaTarget = schemaTarget;
        this.xlsxFilePath = xlsxFilePath;
        this.pathIsTemplate = pathIsTemplate;
        this.pathTemplateToken = pathTemplateToken;
    }


    public static getXmlFilePath(sheetNum: number, info: OpenXmlIo.XlsxFileType) {
        // TODO the path template token may not be a sheet number, could be a resource identifier (i.e. an image or item prop number)
        return (info.pathIsTemplate ? info.xlsxFilePath.split(info.pathTemplateToken).join(<string><any>sheetNum) : info.xlsxFilePath);
    }

}

export = XlsxFileTypeImpl;