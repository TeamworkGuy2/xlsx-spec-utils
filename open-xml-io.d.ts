/// <reference path="../dom-builder/dom/dom-builder.d.ts" />
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

        write(data: T): string;

        // alternatives using existing Documents
        loadFromDom(dom: Document): T;

        saveToDom(data: T): Document;
    }


    /** Helper interface for parsing HTMLElement arrays using a 'reader' function which accepts individual HTMLElements
     */
    export interface ReadOpenXmlElement {

        /** Given a 'reader' function and an array of HTML elements, run the reader against each of the elements and return the results as an array.
         * @return an array of results in the same order as the 'elems' array
         */
        readMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, reader: /*ReadFunc<T>*/(xmlDoc: OpenXmlIo.ParsedFile, elem: HTMLElement, expectedElemName?: string) => T, elems: HTMLElement[]): T[];

        /** Given a 'reader' function and an array of HTML elements, run the reader against each of the elements and return the results as an array.
         * @param expectedElemName the expected nodeName of each of the 'elems', throw an error if any mismatch
         * @return an array of results in the same order as the 'elems' array
         */
        readMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, reader: /*ReadFuncNamed<T>*/(xmlDoc: OpenXmlIo.ParsedFile, elem: HTMLElement, expectedElemName: string) => T, elems: HTMLElement[], expectedElemName: string): T[];
    }


    /** Helper interface for serializing an array of data to HTMLElements using a 'writer' function which accepts individual data items
     */
    export interface WriteOpenXmlElement {

        /** Given a 'writer' function and an array of data objects, run the writer against each of the objects and return the results as an array of HTMLElements.
         * @return an array of HTMLElements in the same order as the 'insts' array
         */
        writeMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, writer: /*WriteFunc<T>*/(xmlDoc: OpenXmlIo.ParsedFile, data: T, expectedElemName?: string) => HTMLElement, insts: T[] | { [id: string]: T }, keys?: string[]): HTMLElement[];

        /** Given a 'writer' function and an array of data objects, run the writer against each of the objects and return the results as an array of HTMLElements.
         * @param expectedElemName the expected nodeName of each of the HTMLElements produced by the writer, throw an error if any mismatch
         * @return an array of HTMLElements in the same order as the 'insts' array
         */
        writeMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, writer: /*WriteFuncNamed<T>*/(xmlDoc: OpenXmlIo.ParsedFile, data: T, expectedElemName: string) => HTMLElement, insts: T[], expectedElemName: string): HTMLElement[];
    }


    export interface Read<T> {
        read(xmlDoc: OpenXmlIo.ParsedFile, elem: HTMLElement, expectedElemName?: string): T; /*ReadFunc<T>*/
    }

    export interface ReadNamed<T> extends Read<T> {
        read(xmlDoc: OpenXmlIo.ParsedFile, elem: HTMLElement, expectedElemName: string): T; /*ReadFuncNamed<T>*/
    }


    export interface ReadFunc<T> {
        (xmlDoc: OpenXmlIo.ParsedFile, elem: HTMLElement, expectedElemName?: string): T;
    }

    export interface ReadFuncNamed<T> {
        (xmlDoc: OpenXmlIo.ParsedFile, elem: HTMLElement, expectedElemName: string): T;
    }


    export interface Write<T> {
        write(xmlDoc: OpenXmlIo.ParsedFile, data: T, expectedElemName?: string): HTMLElement; /*WriteFunc<T>*/
    }

    export interface WriteNamed<T> extends Write<T> {
        write(xmlDoc: OpenXmlIo.ParsedFile, data: T, expectedElemName: string): HTMLElement; /*WriteFuncNamed<T>*/
    }


    export interface WriteFunc<T> {
        (xmlDoc: OpenXmlIo.ParsedFile, data: T, expectedElemName?: string): HTMLElement;
    }

    export interface WriteFuncNamed<T> {
        (xmlDoc: OpenXmlIo.ParsedFile, data: T, expectedElemName: string): HTMLElement;
    }


    export interface ReadWrite<T> extends OpenXmlIo.Read<T>, OpenXmlIo.Write<T> {
    }

    export interface ReadWriteNamed<T> extends OpenXmlIo.ReadNamed<T>, OpenXmlIo.WriteNamed<T> {
    }

}
