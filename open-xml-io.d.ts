/// <reference path="../ts-dom-builder/dom/dom-builder.d.ts" />
/// <reference path="../xlsx-spec-models/open-xml.d.ts" />

/** Read/write interfaces for the 'OpenXml' definition
 * @since 2016-5-27
 */
declare module OpenXmlIo {

    /** Information about a XLSX file, schema url, MIME content type, relative path within the XLSX file zip folder, etc.
     */
    export interface XlsxFileType {
        /** the URL of this file's XML DTD schema */
        schemaUrl: string;
        /** the 'target' attribute for this file type used in XLSX files */
        schemaTarget: string;
        /* the content/mime type name of this file */
        contentType: string;
        /** the relative path inside an unzipped XLSX file where this file resides (the path can be a template string that needs a specific sheet number or resource identifier to complete) */
        xlsxFilePath: string;
        /** refers to 'xlsxFilePath' field, whether the 'xlsxFilePath' is a template string or not */
        pathIsTemplate: boolean;
        /** a string to find/replace in 'xlsxFilePath' with a worksheet number or resource identifier (e.g. 'drawing1.xml', 'drawing2.xml', etc. names can be created using a template string) */
        pathTemplateToken: string;
    }


    /** An instance of a parsed XML file with utilities to help manipulate the resulting XMLDocument
     */
    export interface ParsedFile {
        /** this XML file's parsed DOM */
        dom: XMLDocument;
        /** a DOM builder for this XML document */
        domBldr: DomBuilderFactory;
        /** an XLSX DOM manipulation utility */
        domHelper: DomBuilderHelper;
        /** utlities for reading XML child elements */
        readOpenXml: ReadOpenXmlElement;
        /** utlities for writing XML child elements */
        writeOpenXml: WriteOpenXmlElement;
        /** a validator for XLSX DOM elements */
        validator: DomValidate;
    }


    /** Read and write OpenXml files of a specific type into objects or back to an XML string
     */
    export interface FileReadWriter<T> {
        fileInfo: XlsxFileType;

        read(xmlContentStr: string): T;

        write(inst: T): string;

        // alternatives using existing Documents
        loadFromDom(dom: Document): T;

        saveToDom(inst: T): Document;
    }


    export interface ReadOpenXmlElement {
        readMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, reader: ReadFunc<T>, elems: HTMLElement[]): T[];
        readMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, reader: ReadFuncNamed<T>, elems: HTMLElement[], expectedElemName: string): T[];
    }


    export interface WriteOpenXmlElement {
        writeMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, writer: WriteFunc<T>, insts: T[] | { [id: string]: T }, keys?: string[]): HTMLElement[];
        writeMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, writer: WriteFuncNamed<T>, insts: T[], expectedElemName: string): HTMLElement[];
    }


    export interface Read<T> {
        read: ReadFunc<T>;
    }

    export interface ReadNamed<T> extends Read<T> {
        read: ReadFuncNamed<T>;
    }


    export interface ReadFunc<T> {
        (xmlDoc: OpenXmlIo.ParsedFile, elem: HTMLElement, expectedElemName?: string): T;
    }

    export interface ReadFuncNamed<T> {
        (xmlDoc: OpenXmlIo.ParsedFile, elem: HTMLElement, expectedElemName: string): T;
    }


    export interface Write<T> {
        write: WriteFunc<T>;
    }

    export interface WriteNamed<T> extends Write<T> {
        write: WriteFuncNamed<T>;
    }


    export interface WriteFunc<T> {
        (xmlDoc: OpenXmlIo.ParsedFile, inst: T, expectedElemName?: string): HTMLElement;
    }

    export interface WriteFuncNamed<T> {
        (xmlDoc: OpenXmlIo.ParsedFile, inst: T, expectedElemName: string): HTMLElement;
    }


    export interface ReadWrite<T> extends OpenXmlIo.Read<T>, OpenXmlIo.Write<T> {
    }

    export interface ReadWriteNamed<T> extends OpenXmlIo.ReadNamed<T>, OpenXmlIo.WriteNamed<T> {
    }

}
