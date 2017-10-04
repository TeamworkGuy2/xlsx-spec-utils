/// <reference path="../xlsx-spec-models/open-xml.d.ts" />
/// <reference path="../xlsx-spec-models/open-xml-io.d.ts" />

import XmlFileReadWriter = require("./files/XmlFileReadWriter");
import XlsxFileType = require("./files/XlsxFileType");
import WorksheetUtil = require("./utils/WorksheetUtil");
import CalcChain = require("../xlsx-spec-models/root-types/CalcChain");
import Comments = require("../xlsx-spec-models/root-types/Comments");
import ContentTypes = require("../xlsx-spec-models/root-types/ContentTypes");
import Relationships = require("../xlsx-spec-models/root-types/Relationships");
import SharedStringTable = require("../xlsx-spec-models/root-types/SharedStringTable");
import Stylesheet = require("../xlsx-spec-models/root-types/Stylesheet");
import Workbook = require("../xlsx-spec-models/root-types/Workbook");
import Worksheet = require("../xlsx-spec-models/root-types/Worksheet");
import WorksheetDrawing = require("../xlsx-spec-models/root-types/WorksheetDrawing");

/**
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
module XlsxReaderWriter {

    export interface LoadSettings {
        sheetCount: number;
        calcChain?: boolean;
        comments?: boolean;
        contentTypes?: boolean;
        rels?: boolean;
        sharedStrings?: boolean;
        workbook?: boolean;
        workbookRels?: boolean;
        worksheetDrawing?: boolean;
        worksheetRels?: boolean;
    }


    export interface ParsedWorksheet {
        sheetRels: OpenXml.Relationships;
        comments: OpenXml.Comments;
        worksheet: OpenXml.Worksheet;
    }


    /** The hope is to eventually implement all files, but these are the only ones currently supported
     */
    export interface ParsedXlsxFileInst {
        rels: OpenXml.Relationships;
        contentTypes: OpenXml.ContentTypes;
        workbookRels: OpenXml.Relationships;
        calcChain: OpenXml.CalculationChain;
        sharedStrings: OpenXml.SharedStringTable;
        stylesheet: OpenXml.Stylesheet;
        workbook: OpenXml.Workbook;
        worksheetDrawing: OpenXml.WorksheetDrawing;
        worksheets: ParsedWorksheet[];
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
        ContentTypes: new XlsxFileType("http://schemas.openxmlformats.org/package/2006/content-types", "application/vnd.openxmlformats-package.content-types+xml",
            "[Content_Types].xml", "[Content_Types].xml", false, "#"),
        Core: new XlsxFileType("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "application/vnd.openxmlformats-package.core-properties+xml",
            "docProps/core.xml", "docProps/core.xml", false, null),
        Custom: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            "docProps/custom.xml", "docProps/custom.xml", false, null),
        Drawing: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", "application/vnd.openxmlformats-officedocument.drawing+xml",
            "../drawings/drawing#.xml", "xl/drawings/drawing#.xml", true, "#"),
        ItemProps: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
            "itemProps#.xml", "customXml/itemProps#.xml", true, "#"),
        Rels: new XlsxFileType("http://schemas.openxmlformats.org/package/2006/relationships", "application/vnd.openxmlformats-package.relationships+xml",
            "_rels/.rels", "_rels/.rels", false, "#"),
        SharedStrings: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            "sharedStrings.xml", "xl/sharedStrings.xml", false, null),
        Styles: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
            "styles.xml", "xl/styles.xml", false, null),
        Theme: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "application/vnd.openxmlformats-officedocument.theme+xml",
            "theme/theme#.xml", "xl/theme/theme#.xml", true, "#"),
        Workbook: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "xl/workbook.xml", "xl/workbook.xml", false, null),
        WorkbookRels: new XlsxFileType("http://schemas.openxmlformats.org/package/2006/relationships", "application/vnd.openxmlformats-package.relationships+xml",
            "xl/_rels/workbook.xml.rels", "xl/_rels/workbook.xml.rels", false, "#"),
        Worksheet: new XlsxFileType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            "worksheets/sheet#.xml", "xl/worksheets/sheet#.xml", true, "#"),
        WorksheetRels: new XlsxFileType("http://schemas.openxmlformats.org/package/2006/relationships", "application/vnd.openxmlformats-package.relationships+xml",
            "xl/worksheets/_rels/sheet#.xml.rels", "xl/worksheets/_rels/sheet#.xml.rels", true, "#"),
    };


    export var XlsxFiles = {
        CalcChain: new XmlFileReadWriter(XlsxFileTypes.CalcChain, CalcChain.CalcChain, prepCalcChainForWrite),
        Comments: new XmlFileReadWriter(XlsxFileTypes.Comments, Comments.Comments, prepCommentsForWrite),
        ContentTypes: new XmlFileReadWriter(XlsxFileTypes.ContentTypes, ContentTypes.ContentTypes, prepContentTypesForWrite),
        Rels: new XmlFileReadWriter(XlsxFileTypes.Rels, Relationships.Relationships, prepRelsForWrite),
        SharedStrings: new XmlFileReadWriter(XlsxFileTypes.SharedStrings, SharedStringTable.SharedStringTable, prepSharedStringsForWrite),
        Styles: new XmlFileReadWriter(XlsxFileTypes.Styles, Stylesheet.Stylesheet, prepStylesForWrite),
        Workbook: new XmlFileReadWriter(XlsxFileTypes.Workbook, Workbook.Workbook, prepWorkbookForWrite),
        WorkbookRels: new XmlFileReadWriter(XlsxFileTypes.WorkbookRels, Relationships.Relationships, prepRelsForWrite),
        Worksheet: new XmlFileReadWriter(XlsxFileTypes.Worksheet, Worksheet.Worksheet, prepWorksheetForWrite),
        WorksheetRels: new XmlFileReadWriter(XlsxFileTypes.WorksheetRels, Relationships.Relationships, prepRelsForWrite),
        WorksheetDrawing: new XmlFileReadWriter(XlsxFileTypes.Drawing, WorksheetDrawing.WorksheetDrawing, prepDrawingsForWrite),
    };


    export function readZip<T>(data: Uint8Array, jszip: new (data: Uint8Array) => T): T {
        var firstByte = data[0];
        if (firstByte !== 0x50) {
            throw new Error("Unsupported file " + firstByte);
        }
        var zip = new jszip(data);
        return zip;
    }


    // ==== prep*ForWrite functions for various XLSX internal files ====
    function prepCalcChainForWrite(xmlDoc: OpenXmlIo.WriterContext, inst: OpenXml.CalculationChain) {
        var calcChainDom = <HTMLElement>(<Document>xmlDoc.dom).childNodes[0];
        xmlDoc.removeChilds(calcChainDom);
    }


    function prepCommentsForWrite(xmlDoc: OpenXmlIo.WriterContext, inst: OpenXml.Comments) {
        var commentsDom = <HTMLElement>(<Document>xmlDoc.dom).childNodes[0];
        xmlDoc.removeChilds(commentsDom);
    }


    function prepContentTypesForWrite(xmlDoc: OpenXmlIo.WriterContext, inst: OpenXml.ContentTypes) {
        var contentTypesDom = <HTMLElement>(<Document>xmlDoc.dom).childNodes[0];
        xmlDoc.removeChilds(contentTypesDom);
    }


    function prepDrawingsForWrite(xmlDoc: OpenXmlIo.WriterContext, inst: OpenXml.WorksheetDrawing) {
        var commentsDom = <HTMLElement>(<Document>xmlDoc.dom).childNodes[0];
        xmlDoc.removeChilds(commentsDom);
    }


    function prepRelsForWrite(xmlDoc: OpenXmlIo.WriterContext, inst: OpenXml.Relationships) {
        var relsDom = <HTMLElement>(<Document>xmlDoc.dom).childNodes[0];
        xmlDoc.removeChilds(relsDom);
    }


    function prepSharedStringsForWrite(xmlDoc: OpenXmlIo.WriterContext, inst: OpenXml.SharedStringTable) {
        var sharedStrings = <HTMLElement>(<Document>xmlDoc.dom).childNodes[0];
        xmlDoc.removeNodeAttr(sharedStrings, "count");
        xmlDoc.removeNodeAttr(sharedStrings, "uniqueCount");
        xmlDoc.removeChilds(sharedStrings);
    }


    function prepStylesForWrite(xmlDoc: OpenXmlIo.WriterContext, inst: OpenXml.Stylesheet) {
        var commentsDom = <HTMLElement>(<Document>xmlDoc.dom).childNodes[0];
        xmlDoc.removeChilds(commentsDom);
    }


    function prepWorkbookForWrite(xmlDoc: OpenXmlIo.WriterContext, inst: OpenXml.Workbook) {
        var workbook = <HTMLElement>(<Document>xmlDoc.dom).childNodes[0];
        xmlDoc.removeChilds(workbook);
    }


    function prepWorksheetForWrite(xmlDoc: OpenXmlIo.WriterContext, inst: OpenXml.Worksheet) {
        var worksheet = <HTMLElement>(<Document>xmlDoc.dom).childNodes[0];
        xmlDoc.removeChilds(worksheet);

        WorksheetUtil.updateBounds(inst);
    }


    // ==== functions for reading/writing higher level ParsedXlsxFileInst objects to JSZip files ====

    export function loadXlsxFile(loadSettings: LoadSettings, readFileData: (path: string) => string): ParsedXlsxFileInst {
        // TODO load number of sheets from '[Content_Types].xml' or 'xl/workbook.xml', also need to add media/images/itemProps parsing
        var sheetNum = 1;

        var rels = (loadSettings.rels !== false ? loadXmlFile(sheetNum, readFileData, XlsxFiles.Rels) : null);
        var contentTypes = (loadSettings.contentTypes !== false ? loadXmlFile(sheetNum, readFileData, XlsxFiles.ContentTypes) : null);
        var calcChain = (loadSettings.calcChain !== false ? loadXmlFile(sheetNum, readFileData, XlsxFiles.CalcChain) : null);
        var sharedStrings = (loadSettings.sharedStrings !== false ? loadXmlFile(sheetNum, readFileData, XlsxFiles.SharedStrings) : null);
        var workbook = (loadSettings.workbook !== false ? loadXmlFile(sheetNum, readFileData, XlsxFiles.Workbook) : null);
        var workbookRels = (loadSettings.workbookRels !== false ? loadXmlFile(sheetNum, readFileData, XlsxFiles.WorkbookRels) : null);
        var worksheetDrawing = (loadSettings.worksheetDrawing !== false ? loadXmlFile(sheetNum, readFileData, XlsxFiles.WorksheetDrawing) : null);
        var stylesheet = loadXmlFile(sheetNum, readFileData, XlsxFiles.Styles);

        var worksheets: ParsedWorksheet[] = [];
        for (var i = 0, size = loadSettings.sheetCount; i < size; i++) {
            var sheetRels = (loadSettings.worksheetRels !== false ? loadXmlFile(i + 1, readFileData, XlsxFiles.WorksheetRels) : null);
            var comments = (loadSettings.comments !== false ? loadXmlFile(i + 1, readFileData, XlsxFiles.Comments) : null);
            var worksheet = loadXmlFile(i + 1, readFileData, XlsxFiles.Worksheet);

            worksheets.push({
                sheetRels,
                comments,
                worksheet,
            });
        }

        return {
            rels,
            contentTypes,
            calcChain,
            sharedStrings,
            stylesheet,
            worksheetDrawing,
            workbook,
            workbookRels,
            worksheets,
        };
    }


    export function saveXlsxFile(data: ParsedXlsxFileInst, writeFileData: (path: string, data: string) => void) {
        // these 'files' are shared by all worksheets in a workbook
        if (data.rels != null) { saveXmlFile(null, writeFileData, data.rels, XlsxFiles.Rels); }
        if (data.contentTypes != null) { saveXmlFile(null, writeFileData, data.contentTypes, XlsxFiles.ContentTypes); }
        if (data.calcChain != null) { saveXmlFile(null, writeFileData, data.calcChain, XlsxFiles.CalcChain); }
        if (data.sharedStrings != null) { saveXmlFile(null, writeFileData, data.sharedStrings, XlsxFiles.SharedStrings); }
        if (data.workbook != null) { saveXmlFile(null, writeFileData, data.workbook, XlsxFiles.Workbook); }
        if (data.workbookRels != null) { saveXmlFile(null, writeFileData, data.workbookRels, XlsxFiles.WorkbookRels); }
        if (data.worksheetDrawing != null) { saveXmlFile(1, writeFileData, data.worksheetDrawing, XlsxFiles.WorksheetDrawing); }
        saveXmlFile(null, writeFileData, data.stylesheet, XlsxFiles.Styles);

        // worksheet specific files
        for (var i = 0, size = data.worksheets.length; i < size; i++) {
            // TODO load number of sheets from '[Content_Types].xml' or 'xl/workbook.xml', also fix this to work with media, images, itemProps
            var sheetNum = i + 1;
            var worksheet = data.worksheets[i];

            if (worksheet.sheetRels != null) { saveXmlFile(sheetNum, writeFileData, worksheet.sheetRels, XlsxFiles.WorksheetRels); }
            if (worksheet.comments != null) { saveXmlFile(sheetNum, writeFileData, worksheet.comments, XlsxFiles.Comments); }
            saveXmlFile(sheetNum, writeFileData, worksheet.worksheet, XlsxFiles.Worksheet);
        }
    }


    // TODO finish implementing
    function defaultFileCreator(path: string): ParsedXlsxFileInst {
        var workbookId = "rId1";
        var sheetId = "rId50";
        var rels: OpenXml.Relationships = {
            relationships: [
                { id: workbookId, target: "xl/workbook.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" }
            ]
        };
        var contentTypes: OpenXml.ContentTypes = {
            defaults: [
                { contentType: "application/vnd.openxmlformats-package.relationships+xml", extension: "rels" },
                { contentType: "application/xml", extension: "xml"}
            ],
            overrides: [
                { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", partName: "/xl/workbook.xml" },
                { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", partName: "/xl/sharedStrings.xml" },
                { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", partName: "/xl/styles.xml" },
                { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", partName: "/xl/worksheets/sheet1.xml" }
            ]
        };
        var sharedStrings: OpenXml.SharedStringTable = { count: 0, uniqueCount: 0, sis: [] };
        var stylesheet = createDefaultStylesheet();
        var workbook: OpenXml.Workbook = {
            sheets: {
                sheets: [{ id: sheetId, sheetId: 1, name: "Sheet 1" }]
            }
        };
        var workbookRels: OpenXml.Relationships = {
            relationships: [
                { id: "rId45", target: "sharedStrings.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" },
                { id: "rId46", target: "styles.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" },
                { id: sheetId, target: "worksheets/sheet1.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" }
            ]
        };
        var worksheets: ParsedWorksheet[] = [{
            comments: null,
            sheetRels: { relationships: [] },
            worksheet: {
                cols: [{ cols: [{ max: 1, min: 1, width: 9.140625 }] }],
                dimension: { ref: "A1:A1" },
                sheetData: { rows: [] }
            }
        }];

        return {
            rels,
            contentTypes,
            calcChain: null,
            sharedStrings,
            stylesheet,
            worksheetDrawing: null,
            workbook,
            workbookRels,
            worksheets
        };
    }


    function loadXmlFile<T>(sheetNum: number, readFileData: (path: string) => string, loader: OpenXmlIo.FileReadWriter<T>): T {
        var path = XlsxFileType.getXmlFilePath(sheetNum, loader.fileInfo);
        var data = readFileData(path);
        var inst = data != null ? loader.read(data) : null;
        return inst;
    }


    function saveXmlFile<T>(sheetNum: number, writeFileData: (path: string, data: string) => void, data: T, writer: OpenXmlIo.FileReadWriter<T>): void {
        var path = XlsxFileType.getXmlFilePath(sheetNum, writer.fileInfo);
        var dataStr = writer.write(data);
        writeFileData(path, dataStr);
    }


    function createDefaultStylesheet(): OpenXml.Stylesheet {
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
}

export = XlsxReaderWriter;