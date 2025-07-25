"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.XlsxReaderWriter = void 0;
/// <reference path="../xlsx-spec-models/open-xml.d.ts" />
/// <reference path="../xlsx-spec-models/open-xml-io.d.ts" />
var CalcChain_1 = require("xlsx-spec-models/root-types/CalcChain");
var Comments_1 = require("xlsx-spec-models/root-types/Comments");
var ContentTypes_1 = require("xlsx-spec-models/root-types/ContentTypes");
var Relationships_1 = require("xlsx-spec-models/root-types/Relationships");
var SharedStringTable_1 = require("xlsx-spec-models/root-types/SharedStringTable");
var Stylesheet_1 = require("xlsx-spec-models/root-types/Stylesheet");
var Workbook_1 = require("xlsx-spec-models/root-types/Workbook");
var Worksheet_1 = require("xlsx-spec-models/root-types/Worksheet");
var WorksheetDrawing_1 = require("xlsx-spec-models/root-types/WorksheetDrawing");
var XmlFileReadWriter_1 = require("./files/XmlFileReadWriter");
var XlsxFileType_1 = require("./files/XlsxFileType");
var WorksheetUtil_1 = require("./utils/WorksheetUtil");
/**
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
var XlsxReaderWriter;
(function (XlsxReaderWriter) {
    XlsxReaderWriter.RootNamespaceUrl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    // XML namespaces and flags for the various sub files inside a zipped Open XML Spreadsheet file
    XlsxReaderWriter.XlsxFileTypes = {
        App: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "application/vnd.openxmlformats-officedocument.extended-properties+xml", "docProps/app.xml", "docProps/app.xml", false, null),
        CalcChain: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain", "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml", "calcChain.xml", "xl/calcChain.xml", false, null),
        Comments: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", "../comments#.xml", "xl/comments#.xml", true, "#"),
        ContentTypes: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/package/2006/content-types", "application/vnd.openxmlformats-package.content-types+xml", "[Content_Types].xml", "[Content_Types].xml", false, "#"),
        Core: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "application/vnd.openxmlformats-package.core-properties+xml", "docProps/core.xml", "docProps/core.xml", false, null),
        Custom: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", "application/vnd.openxmlformats-officedocument.custom-properties+xml", "docProps/custom.xml", "docProps/custom.xml", false, null),
        Drawing: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", "application/vnd.openxmlformats-officedocument.drawing+xml", "../drawings/drawing#.xml", "xl/drawings/drawing#.xml", true, "#"),
        ItemProps: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml", "itemProps#.xml", "customXml/itemProps#.xml", true, "#"),
        Rels: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/package/2006/relationships", "application/vnd.openxmlformats-package.relationships+xml", "_rels/.rels", "_rels/.rels", false, "#"),
        SharedStrings: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", "sharedStrings.xml", "xl/sharedStrings.xml", false, null),
        Styles: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", "styles.xml", "xl/styles.xml", false, null),
        Theme: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "application/vnd.openxmlformats-officedocument.theme+xml", "theme/theme#.xml", "xl/theme/theme#.xml", true, "#"),
        Workbook: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "xl/workbook.xml", "xl/workbook.xml", false, null),
        WorkbookRels: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/package/2006/relationships", "application/vnd.openxmlformats-package.relationships+xml", "xl/_rels/workbook.xml.rels", "xl/_rels/workbook.xml.rels", false, "#"),
        Worksheet: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", "worksheets/sheet#.xml", "xl/worksheets/sheet#.xml", true, "#"),
        WorksheetRels: new XlsxFileType_1.XlsxFileType("http://schemas.openxmlformats.org/package/2006/relationships", "application/vnd.openxmlformats-package.relationships+xml", "xl/worksheets/_rels/sheet#.xml.rels", "xl/worksheets/_rels/sheet#.xml.rels", true, "#"),
    };
    XlsxReaderWriter.XlsxFiles = {
        CalcChain: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.CalcChain, CalcChain_1.CalcChain, prepCalcChainForWrite),
        Comments: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.Comments, Comments_1.Comments, prepCommentsForWrite),
        ContentTypes: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.ContentTypes, ContentTypes_1.ContentTypes, prepContentTypesForWrite),
        Rels: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.Rels, Relationships_1.Relationships, prepRelsForWrite),
        SharedStrings: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.SharedStrings, SharedStringTable_1.SharedStringTable, prepSharedStringsForWrite),
        Styles: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.Styles, Stylesheet_1.Stylesheet, prepStylesForWrite),
        Workbook: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.Workbook, Workbook_1.Workbook, prepWorkbookForWrite),
        WorkbookRels: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.WorkbookRels, Relationships_1.Relationships, prepRelsForWrite),
        Worksheet: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.Worksheet, Worksheet_1.Worksheet, prepWorksheetForWrite),
        WorksheetRels: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.WorksheetRels, Relationships_1.Relationships, prepRelsForWrite),
        WorksheetDrawing: new XmlFileReadWriter_1.XmlFileReadWriter(XlsxReaderWriter.XlsxFileTypes.Drawing, WorksheetDrawing_1.WorksheetDrawing, prepDrawingsForWrite),
    };
    function readZip(data, jszip) {
        var firstByte = data[0];
        if (firstByte !== 0x50) {
            throw new Error("Unsupported file " + firstByte);
        }
        var zip = new jszip(data);
        return zip;
    }
    XlsxReaderWriter.readZip = readZip;
    // ==== prep*ForWrite functions for various XLSX internal files ====
    function prepCalcChainForWrite(xmlDoc, inst) {
        var calcChainDom = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(calcChainDom);
    }
    function prepCommentsForWrite(xmlDoc, inst) {
        var commentsDom = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(commentsDom);
    }
    function prepContentTypesForWrite(xmlDoc, inst) {
        var contentTypesDom = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(contentTypesDom);
    }
    function prepDrawingsForWrite(xmlDoc, inst) {
        var commentsDom = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(commentsDom);
    }
    function prepRelsForWrite(xmlDoc, inst) {
        var relsDom = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(relsDom);
    }
    function prepSharedStringsForWrite(xmlDoc, inst) {
        var sharedStrings = xmlDoc.dom.childNodes[0];
        xmlDoc.removeAttr(sharedStrings, "count");
        xmlDoc.removeAttr(sharedStrings, "uniqueCount");
        xmlDoc.removeChilds(sharedStrings);
    }
    function prepStylesForWrite(xmlDoc, inst) {
        var commentsDom = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(commentsDom);
    }
    function prepWorkbookForWrite(xmlDoc, inst) {
        var workbook = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(workbook);
    }
    function prepWorksheetForWrite(xmlDoc, inst) {
        var worksheet = xmlDoc.dom.childNodes[0];
        xmlDoc.removeChilds(worksheet);
        WorksheetUtil_1.WorksheetUtil.updateBounds(inst);
    }
    // ==== functions for reading/writing higher level ParsedXlsxFileInst objects to JSZip files ====
    function loadXlsxFile(loadSettings, readFileData) {
        // TODO load number of sheets from '[Content_Types].xml' or 'xl/workbook.xml', also need to add media/images/itemProps parsing
        var sheetNum = 1;
        var rels = (loadSettings.rels !== false ? loadXmlFile(sheetNum, readFileData, XlsxReaderWriter.XlsxFiles.Rels) : null);
        var contentTypes = (loadSettings.contentTypes !== false ? loadXmlFile(sheetNum, readFileData, XlsxReaderWriter.XlsxFiles.ContentTypes) : null);
        var calcChain = (loadSettings.calcChain !== false ? loadXmlFile(sheetNum, readFileData, XlsxReaderWriter.XlsxFiles.CalcChain) : null);
        var sharedStrings = (loadSettings.sharedStrings !== false ? loadXmlFile(sheetNum, readFileData, XlsxReaderWriter.XlsxFiles.SharedStrings) : null);
        var workbook = (loadSettings.workbook !== false ? loadXmlFile(sheetNum, readFileData, XlsxReaderWriter.XlsxFiles.Workbook) : null);
        var workbookRels = (loadSettings.workbookRels !== false ? loadXmlFile(sheetNum, readFileData, XlsxReaderWriter.XlsxFiles.WorkbookRels) : null);
        var worksheetDrawing = (loadSettings.worksheetDrawing !== false ? loadXmlFile(sheetNum, readFileData, XlsxReaderWriter.XlsxFiles.WorksheetDrawing) : null);
        var stylesheet = loadXmlFile(sheetNum, readFileData, XlsxReaderWriter.XlsxFiles.Styles);
        var worksheets = [];
        for (var i = 0, size = loadSettings.sheetCount; i < size; i++) {
            var sheetRels = (loadSettings.worksheetRels !== false ? loadXmlFile(i + 1, readFileData, XlsxReaderWriter.XlsxFiles.WorksheetRels) : null);
            var comments = (loadSettings.comments !== false ? loadXmlFile(i + 1, readFileData, XlsxReaderWriter.XlsxFiles.Comments) : null);
            var worksheet = loadXmlFile(i + 1, readFileData, XlsxReaderWriter.XlsxFiles.Worksheet);
            worksheets.push({
                sheetRels: sheetRels,
                comments: comments,
                worksheet: worksheet,
            });
        }
        return {
            rels: rels,
            contentTypes: contentTypes,
            calcChain: calcChain,
            sharedStrings: sharedStrings,
            stylesheet: stylesheet,
            worksheetDrawing: worksheetDrawing,
            workbook: workbook,
            workbookRels: workbookRels,
            worksheets: worksheets,
        };
    }
    XlsxReaderWriter.loadXlsxFile = loadXlsxFile;
    function saveXlsxFile(data, writeFileData) {
        // these 'files' are shared by all worksheets in a workbook
        if (data.rels != null) {
            saveXmlFile(null, writeFileData, data.rels, XlsxReaderWriter.XlsxFiles.Rels);
        }
        if (data.contentTypes != null) {
            saveXmlFile(null, writeFileData, data.contentTypes, XlsxReaderWriter.XlsxFiles.ContentTypes);
        }
        if (data.calcChain != null) {
            saveXmlFile(null, writeFileData, data.calcChain, XlsxReaderWriter.XlsxFiles.CalcChain);
        }
        if (data.sharedStrings != null) {
            saveXmlFile(null, writeFileData, data.sharedStrings, XlsxReaderWriter.XlsxFiles.SharedStrings);
        }
        if (data.workbook != null) {
            saveXmlFile(null, writeFileData, data.workbook, XlsxReaderWriter.XlsxFiles.Workbook);
        }
        if (data.workbookRels != null) {
            saveXmlFile(null, writeFileData, data.workbookRels, XlsxReaderWriter.XlsxFiles.WorkbookRels);
        }
        if (data.worksheetDrawing != null) {
            saveXmlFile(1, writeFileData, data.worksheetDrawing, XlsxReaderWriter.XlsxFiles.WorksheetDrawing);
        }
        saveXmlFile(null, writeFileData, data.stylesheet, XlsxReaderWriter.XlsxFiles.Styles);
        // worksheet specific files
        for (var i = 0, size = data.worksheets.length; i < size; i++) {
            // TODO load number of sheets from '[Content_Types].xml' or 'xl/workbook.xml', also fix this to work with media, images, itemProps
            var sheetNum = i + 1;
            var worksheet = data.worksheets[i];
            if (worksheet.sheetRels != null) {
                saveXmlFile(sheetNum, writeFileData, worksheet.sheetRels, XlsxReaderWriter.XlsxFiles.WorksheetRels);
            }
            if (worksheet.comments != null) {
                saveXmlFile(sheetNum, writeFileData, worksheet.comments, XlsxReaderWriter.XlsxFiles.Comments);
            }
            saveXmlFile(sheetNum, writeFileData, worksheet.worksheet, XlsxReaderWriter.XlsxFiles.Worksheet);
        }
    }
    XlsxReaderWriter.saveXlsxFile = saveXlsxFile;
    // TODO finish implementing
    function defaultFileCreator(path) {
        var workbookId = "rId1";
        var sheetId = "rId50";
        var rels = {
            relationships: [
                { id: workbookId, target: "xl/workbook.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" }
            ]
        };
        var contentTypes = {
            defaults: [
                { contentType: "application/vnd.openxmlformats-package.relationships+xml", extension: "rels" },
                { contentType: "application/xml", extension: "xml" }
            ],
            overrides: [
                { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", partName: "/xl/workbook.xml" },
                { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", partName: "/xl/sharedStrings.xml" },
                { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", partName: "/xl/styles.xml" },
                { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", partName: "/xl/worksheets/sheet1.xml" }
            ]
        };
        var sharedStrings = { count: 0, uniqueCount: 0, sis: [] };
        var stylesheet = createDefaultStylesheet();
        var workbook = {
            sheets: {
                sheets: [{ id: sheetId, sheetId: 1, name: "Sheet 1" }]
            }
        };
        var workbookRels = {
            relationships: [
                { id: "rId45", target: "sharedStrings.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" },
                { id: "rId46", target: "styles.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" },
                { id: sheetId, target: "worksheets/sheet1.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" }
            ]
        };
        var worksheets = [{
                comments: null,
                sheetRels: { relationships: [] },
                worksheet: {
                    cols: [{ cols: [{ max: 1, min: 1, width: 9.140625 }] }],
                    dimension: { ref: "A1:A1" },
                    sheetData: { rows: [] }
                }
            }];
        return {
            rels: rels,
            contentTypes: contentTypes,
            calcChain: null,
            sharedStrings: sharedStrings,
            stylesheet: stylesheet,
            worksheetDrawing: null,
            workbook: workbook,
            workbookRels: workbookRels,
            worksheets: worksheets
        };
    }
    function loadXmlFile(sheetNum, readFileData, loader) {
        var path = XlsxFileType_1.XlsxFileType.getXmlFilePath(sheetNum, loader.fileInfo);
        var data = readFileData(path);
        var inst = data != null ? loader.read(data) : null;
        return inst;
    }
    function saveXmlFile(sheetNum, writeFileData, data, writer) {
        var path = XlsxFileType_1.XlsxFileType.getXmlFilePath(sheetNum, writer.fileInfo);
        var dataStr = writer.write(data);
        writeFileData(path, dataStr);
    }
    function createDefaultStylesheet() {
        return {
            borders: {
                count: 1,
                borders: [{ left: {}, right: {}, top: {}, bottom: {}, diagonal: {} }]
            },
            cellStyleXfs: {
                count: 1,
                xfs: [{ borderId: 0, fillId: 0, fontId: 0, numFmtId: 0 }]
            },
            cellXfs: {
                count: 1,
                xfs: [{ borderId: 0, fillId: 0, fontId: 0, numFmtId: 0 }]
            },
            cellStyles: {
                count: 1,
                cellStyles: [{ builtinId: 0, xfId: 0, name: "Normal" }]
            },
            dxfs: {
                count: 1,
                dxfs: [{}]
            },
            fills: {
                count: 2,
                fills: [{}, { patternFill: { patternType: "gray125", bgColor: { rgb: "FF333333" }, fgColor: { rgb: "FF333333" } } }]
            },
            fonts: {
                count: 1,
                fonts: [{}]
            },
            numFmts: {
                count: 0,
                numFmts: []
            }
        };
    }
})(XlsxReaderWriter = exports.XlsxReaderWriter || (exports.XlsxReaderWriter = {}));
