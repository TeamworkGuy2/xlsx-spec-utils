"use strict";
var XmlFileReadWriter = require("./files/XmlFileReadWriter");
var XlsxFileType = require("./files/XlsxFileType");
var WorksheetUtil = require("./utils/WorksheetUtil");
var CalculationChain = require("../xlsx-spec-models/types/CalculationChain");
var Comments = require("../xlsx-spec-models/types/Comments");
var SharedStringTable = require("../xlsx-spec-models/types/SharedStringTable");
var Stylesheet = require("../xlsx-spec-models/types/Stylesheet");
var Worksheet = require("../xlsx-spec-models/types/Worksheet");
var WorksheetDrawing = require("../xlsx-spec-models/types/WorksheetDrawing");
/**
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
var ExcelTemplateLoad;
(function (ExcelTemplateLoad) {
    ExcelTemplateLoad.RootNamespaceUrl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    // XML namespaces and flags for the various sub files inside a zipped Open XML Spreadsheet file
    ExcelTemplateLoad.XlsxFileTypes = {
        App: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "application/vnd.openxmlformats-officedocument.extended-properties+xml", "docProps/app.xml", "docProps/app.xml", false, null),
        CalcChain: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain", "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml", "calcChain.xml", "xl/calcChain.xml", false, null),
        Comments: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", "../comments#.xml", "xl/comments#.xml", true, "#"),
        Core: new XlsxFileType("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "application/vnd.openxmlformats-package.core-properties+xml", "docProps/core.xml", "docProps/core.xml", false, null),
        Custom: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", "application/vnd.openxmlformats-officedocument.custom-properties+xml", "docProps/custom.xml", "docProps/custom.xml", false, null),
        Drawing: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", "application/vnd.openxmlformats-officedocument.drawing+xml", "../drawings/drawing#.xml", "xl/drawings/drawing#.xml", true, "#"),
        ItemProps: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml", "itemProps#.xml", "customXml/itemProps#.xml", true, "#"),
        SharedStrings: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", "sharedStrings.xml", "xl/sharedStrings.xml", false, null),
        Styles: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", "styles.xml", "xl/styles.xml", false, null),
        Theme: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "application/vnd.openxmlformats-officedocument.theme+xml", "theme/theme#.xml", "xl/theme/theme#.xml", true, "#"),
        Workbook: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "xl/workbook.xml", "xl/workbook.xml", false, null),
        Worksheet: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", "worksheets/sheet#.xml", "xl/worksheets/sheet#.xml", true, "#"),
    };
    ExcelTemplateLoad.XlsxFiles = {
        CalcChain: new XmlFileReadWriter(ExcelTemplateLoad.XlsxFileTypes.CalcChain, CalculationChain, prepCalcChainForWrite),
        Comments: new XmlFileReadWriter(ExcelTemplateLoad.XlsxFileTypes.Comments, Comments, prepCommentsForWrite),
        SharedStrings: new XmlFileReadWriter(ExcelTemplateLoad.XlsxFileTypes.SharedStrings, SharedStringTable, prepSharedStringsForWrite),
        Styles: new XmlFileReadWriter(ExcelTemplateLoad.XlsxFileTypes.Styles, Stylesheet, prepStylesForWrite),
        Worksheet: new XmlFileReadWriter(ExcelTemplateLoad.XlsxFileTypes.Worksheet, Worksheet, prepWorksheetForWrite),
        WorksheetDrawing: new XmlFileReadWriter(ExcelTemplateLoad.XlsxFileTypes.Drawing, WorksheetDrawing, prepDrawingsForWrite),
    };
    function readZip(data, jszip) {
        var firstByte = data[0];
        if (firstByte !== 0x50) {
            throw new Error("Unsupported file " + firstByte);
        }
        var zip = new jszip(data);
        return zip;
    }
    ExcelTemplateLoad.readZip = readZip;
    // ==== prep*ForWrite functions for various XLSX internal files ====
    function prepCalcChainForWrite(xmlDoc, inst) {
        var calcChainDom = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(calcChainDom);
    }
    function prepCommentsForWrite(xmlDoc, inst) {
        var commentsDom = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(commentsDom);
    }
    function prepDrawingsForWrite(xmlDoc, inst) {
        var commentsDom = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(commentsDom);
    }
    function prepSharedStringsForWrite(xmlDoc, inst) {
        var dom = xmlDoc.dom;
        var sharedStrings = dom.childNodes[0];
        xmlDoc.removeNodeAttr(sharedStrings, "count");
        xmlDoc.removeNodeAttr(sharedStrings, "uniqueCount");
        xmlDoc.removeChilds(sharedStrings);
    }
    function prepStylesForWrite(xmlDoc, inst) {
        var commentsDom = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(commentsDom);
    }
    function prepWorksheetForWrite(xmlDoc, inst) {
        var worksheet = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(worksheet);
        WorksheetUtil.updateBounds(inst);
    }
    // ==== functions for reading/writing higher level ParsedXlsxFileInst objects to JSZip files ====
    function loadXlsxFile(loadSettings, readFileData) {
        // TODO load number of sheets from '[Content_Types].xml' or 'xl/workbook.xml', also need to add media/images/itemProps parsing
        var sheetNum = 1;
        var calcChain = (loadSettings.calcChain !== false ? loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.CalcChain) : null);
        var comments = (loadSettings.comments !== false ? loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.Comments) : null);
        var sharedStrings = (loadSettings.sharedStrings !== false ? loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.SharedStrings) : null);
        var worksheetDrawing = (loadSettings.worksheetDrawing !== false ? loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.WorksheetDrawing) : null);
        var worksheet = loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.Worksheet);
        var stylesheet = loadXmlFile(sheetNum, readFileData, ExcelTemplateLoad.XlsxFiles.Styles);
        return {
            calcChain: calcChain,
            sharedStrings: sharedStrings,
            stylesheet: stylesheet,
            worksheetDrawing: worksheetDrawing,
            worksheets: [{
                    comments: comments,
                    worksheet: worksheet,
                }],
        };
    }
    ExcelTemplateLoad.loadXlsxFile = loadXlsxFile;
    function saveXlsxFile(data, writeFileData) {
        // these 'files' are shared all worksheets in a workbook
        if (data.calcChain != null) {
            saveXmlFile(null, writeFileData, data.calcChain, ExcelTemplateLoad.XlsxFiles.CalcChain);
        }
        if (data.sharedStrings != null) {
            saveXmlFile(null, writeFileData, data.sharedStrings, ExcelTemplateLoad.XlsxFiles.SharedStrings);
        }
        if (data.worksheetDrawing != null) {
            saveXmlFile(1, writeFileData, data.worksheetDrawing, ExcelTemplateLoad.XlsxFiles.WorksheetDrawing);
        }
        saveXmlFile(null, writeFileData, data.stylesheet, ExcelTemplateLoad.XlsxFiles.Styles);
        for (var i = 0, size = data.worksheets.length; i < size; i++) {
            // TODO load number of sheets from '[Content_Types].xml' or 'xl/workbook.xml', also fix this to work with media, images, itemProps
            var sheetNum = i + 1;
            var worksheet = data.worksheets[i];
            if (worksheet.comments != null) {
                saveXmlFile(sheetNum, writeFileData, worksheet.comments, ExcelTemplateLoad.XlsxFiles.Comments);
            }
            saveXmlFile(sheetNum, writeFileData, worksheet.worksheet, ExcelTemplateLoad.XlsxFiles.Worksheet);
        }
    }
    ExcelTemplateLoad.saveXlsxFile = saveXlsxFile;
    function loadXmlFile(sheetNum, readFileData, loader) {
        var info = loader.fileInfo;
        // TODO the path template token may not be a sheet number, could be a resource identifier (i.e. an image or item prop number)
        var path = info.pathIsTemplate ? info.xlsxFilePath.split(info.pathTemplateToken).join(sheetNum) : info.xlsxFilePath;
        var data = readFileData(path);
        var inst = data != null ? loader.read(data) : null;
        return inst;
    }
    function saveXmlFile(sheetNum, writeFileData, data, writer) {
        var info = writer.fileInfo;
        // TODO the path template token may not be a sheet number, could be a resource identifier (i.e. an image or item prop number)
        var path = info.pathIsTemplate ? info.xlsxFilePath.split(info.pathTemplateToken).join(sheetNum) : info.xlsxFilePath;
        var dataStr = writer.write(data);
        writeFileData(path, dataStr);
    }
})(ExcelTemplateLoad || (ExcelTemplateLoad = {}));
module.exports = ExcelTemplateLoad;
