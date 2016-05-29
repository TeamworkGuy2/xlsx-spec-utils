import XmlFileReadWriter = require("./files/XmlFileReadWriter");
import XlsxFileType = require("./files/XlsxFileType");
import WorksheetUtil = require("./utils/WorksheetUtil");
import CalculationChain = require("../xlsx-spec-models/types/CalculationChain");
import Comments = require("../xlsx-spec-models/types/Comments");
import SharedStringTable = require("../xlsx-spec-models/types/SharedStringTable");
import Stylesheet = require("../xlsx-spec-models/types/Stylesheet");
import Worksheet = require("../xlsx-spec-models/types/Worksheet");

/**
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
module ExcelTemplateLoad {

    /** The hope is to eventually implement all files, but these are the only ones currently supported
     */
    export interface ParsedXlsxFileInst {
        calcChain: OpenXml.CalculationChain;
        sharedStrings: OpenXml.SharedStringTable;
        stylesheet: OpenXml.Stylesheet;
        worksheets: {
            comments: OpenXml.Comments;
            worksheet: OpenXml.Worksheet;
        }[];
    }


    export var RootNamespaceUrl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // XML namespaces and flags for the various sub files inside a zipped Open XML Spreadsheet file
    export var XlsxFileTypes = {
        App: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            "docProps/app.xml", "docProps/app.xml", false, null),
        CalcChain: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain", "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
            "calcChain.xml", "xl/calcChain.xml", false, null),
        Comments: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            "../comments#.xml", "xl/comments#.xml", true, "#"),
        Core: new XlsxFileType("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "application/vnd.openxmlformats-package.core-properties+xml",
            "docProps/core.xml", "docProps/core.xml", false, null),
        Custom: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            "docProps/custom.xml", "docProps/custom.xml", false, null),
        Drawing: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", "application/vnd.openxmlformats-officedocument.drawing+xml",
            "../drawings/drawing#.xml", "xl/drawings/drawing#.xml", true, "#"),
        ItemProps: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
            "itemProps#.xml", "customXml/itemProps#.xml", true, "#"),
        SharedStrings: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            "sharedStrings.xml", "xl/sharedStrings.xml", false, null),
        Styles: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
            "styles.xml", "xl/styles.xml", false, null),
        Theme: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "application/vnd.openxmlformats-officedocument.theme+xml",
            "theme/theme#.xml", "xl/theme/theme#.xml", true, "#"),
        Workbook: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "xl/workbook.xml", "xl/workbook.xml", false, null),
        Worksheet: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            "worksheets/sheet#.xml", "xl/worksheets/sheet#.xml", true, "#"),
    };


    export var XlsxFiles = {
        CalcChain: new XmlFileReadWriter(XlsxFileTypes.CalcChain, CalculationChain, prepCalcChainForWrite),
        Comments: new XmlFileReadWriter(XlsxFileTypes.Comments, Comments, prepCommentsForWrite),
        SharedStrings: new XmlFileReadWriter(XlsxFileTypes.SharedStrings, SharedStringTable, prepSharedStringsForWrite),
        Styles: new XmlFileReadWriter(XlsxFileTypes.Styles, Stylesheet, prepStylesForWrite),
        Worksheet: new XmlFileReadWriter(XlsxFileTypes.Worksheet, Worksheet, prepWorksheetForWrite),
    };


    export function readZip<T>(data: Uint8Array, jszip: new (data: Uint8Array) => T): T {
        var isfile = false;
        var firstByte = data[0];
        if (firstByte !== 0x50) {
            throw new Error("Unsupported file " + firstByte);
        }
        var zip = new jszip(data);
        return zip;
    }


    // ==== prep*ForWrite functions for various XLSX internal files ====
    function prepCalcChainForWrite(xmlDoc: OpenXmlIo.ParsedFile, inst: OpenXml.CalculationChain) {
        var calcChainDom = <HTMLElement>xmlDoc.dom.childNodes[0];
        xmlDoc.domHelper.removeChilds(calcChainDom);
    }


    function prepCommentsForWrite(xmlDoc: OpenXmlIo.ParsedFile, inst: OpenXml.Comments) {
        var commentsDom = <HTMLElement>xmlDoc.dom.childNodes[0];
        xmlDoc.domHelper.removeChilds(commentsDom);
    }


    function prepSharedStringsForWrite(xmlDoc: OpenXmlIo.ParsedFile, inst: OpenXml.SharedStringTable) {
        var dom = xmlDoc.dom;
        var sharedStrings = <HTMLElement>dom.childNodes[0];
        xmlDoc.domHelper.removeNodeAttr(sharedStrings, "count");
        xmlDoc.domHelper.removeNodeAttr(sharedStrings, "uniqueCount");
        xmlDoc.domHelper.removeChilds(sharedStrings);
    }


    function prepStylesForWrite(xmlDoc: OpenXmlIo.ParsedFile, inst: OpenXml.Stylesheet) {
        var commentsDom = <HTMLElement>xmlDoc.dom.childNodes[0];
        xmlDoc.domHelper.removeChilds(commentsDom);
    }


    function prepWorksheetForWrite(xmlDoc: OpenXmlIo.ParsedFile, inst: OpenXml.Worksheet) {
        var worksheet = <HTMLElement>xmlDoc.dom.childNodes[0];
        xmlDoc.domHelper.removeChilds(worksheet);
        //xmlDoc.domHelper.removeChilds(xmlDoc.domHelper.queryOneChild(worksheet, "dimension"));
        //xmlDoc.domHelper.removeChilds(xmlDoc.domHelper.queryOneChild(worksheet, "sheetViews"));
        //xmlDoc.domHelper.removeChilds(xmlDoc.domHelper.queryOneChild(worksheet, "sheetFormatPr"));
        //xmlDoc.domHelper.removeChilds(xmlDoc.domHelper.queryOneChild(worksheet, "cols"));
        //xmlDoc.domHelper.removeChilds(xmlDoc.domHelper.queryOneChild(worksheet, "sheetData"));
        //xmlDoc.domHelper.removeChilds(xmlDoc.domHelper.queryOneChild(worksheet, "pageMargins"));
        //xmlDoc.domHelper.removeChilds(xmlDoc.domHelper.queryOneChild(worksheet, "pageSetup"));
        //xmlDoc.domHelper.removeChilds(xmlDoc.domHelper.queryOneChild(worksheet, "headerFooter"));
        //xmlDoc.domHelper.removeChilds(xmlDoc.domHelper.queryOneChild(worksheet, "drawing"));
        //xmlDoc.domHelper.removeChilds(xmlDoc.domHelper.queryOneChild(worksheet, "legacyDrawing"));

        WorksheetUtil.updateBounds(inst);
    }


    // ==== functions for reading/writing higher level ParsedXlsxFileInst objects to JSZip files ====

    export function loadExcelFileInst(readFileData: (path: string) => string): ExcelTemplateLoad.ParsedXlsxFileInst {
        // TODO load number of sheets from '[Content_Types].xml' or 'xl/workbook.xml', also fix for media/images/itemProps
        var sheetNum = 1;

        var calcChain = loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.CalcChain);
        var comments = loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.Comments);
        var sharedStrings = loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.SharedStrings);
        var worksheet = loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.Worksheet);
        var stylesheet = loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.Styles);

        return {
            calcChain,
            sharedStrings,
            stylesheet,
            worksheets: [{
                comments,
                worksheet,
            }],
        };
    }


    export function saveExcelFileInst(data: ExcelTemplateLoad.ParsedXlsxFileInst, writeFileData: (path: string, data: string) => void) {
        // these 'files' are shared all worksheets in a workbook
        saveXmlFile(null, writeFileData, data.calcChain, ExcelTemplateLoad.XlsxFiles.CalcChain);
        saveXmlFile(null, writeFileData, data.sharedStrings, ExcelTemplateLoad.XlsxFiles.SharedStrings);
        saveXmlFile(null, writeFileData, data.stylesheet, ExcelTemplateLoad.XlsxFiles.Styles);

        for (var i = 0, size = data.worksheets.length; i < size; i++) {
            // TODO load number of sheets from '[Content_Types].xml' or 'xl/workbook.xml', also fix this to work with media, images, itemProps
            var sheetNum = i + 1;
            var worksheet = data.worksheets[i];

            saveXmlFile(sheetNum, writeFileData, worksheet.comments, ExcelTemplateLoad.XlsxFiles.Comments);
            saveXmlFile(sheetNum, writeFileData, worksheet.worksheet, ExcelTemplateLoad.XlsxFiles.Worksheet);
        }
    }


    function loadXmlFile<T>(sheetNum: number, readFileData: (path: string) => string, loader: OpenXmlIo.FileReadWriter<T>): T {
        var info = loader.fileInfo;
        // TODO the path template token may not be a sheet number, could be a resource identifier (i.e. an image or item prop number)
        var path = info.pathIsTemplate ? info.xlsxFilePath.split(info.pathTemplateToken).join(<string><any>sheetNum) : info.xlsxFilePath;
        var data = readFileData(path);
        var inst = loader.read(data);
        return inst;
    }


    function saveXmlFile<T>(sheetNum: number, writeFileData: (path: string, data: string) => void, data: T, writer: OpenXmlIo.FileReadWriter<T>): void {
        var info = writer.fileInfo;
        // TODO the path template token may not be a sheet number, could be a resource identifier (i.e. an image or item prop number)
        var path = info.pathIsTemplate ? info.xlsxFilePath.split(info.pathTemplateToken).join(<string><any>sheetNum) : info.xlsxFilePath;
        var dataStr = writer.write(data);
        writeFileData(path, dataStr);
    }

}

export = ExcelTemplateLoad;