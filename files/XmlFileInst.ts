﻿import { DomBuilderFactory } from "@twg2/dom-builder/dom/DomBuilderFactory";
import { DomBuilderHelper } from "@twg2/dom-builder/dom/DomBuilderHelper";
import { XlsxDomErrorsImpl } from "../errors/XlsxDomErrorsImpl";

/** Implementation of OpenXmlIo.ParsedFile, contains:
 * - An XMLDocument containing the file data
 * - A DomBuilderFactory for creating new DOM elements when writing data back to the file
 * - An XlsxDom utility object with methods to make reading/writing data to DOM elements easier
 * - An XlsxDomValidator for checking that DOM elements match expected formats and throwing errors when not
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
export module XmlFileInst {

    export class XmlDocFile<D extends DocumentLike = DocumentLike>
        extends DomBuilderHelper
        implements OpenXmlIo.ReaderContext, OpenXmlIo.WriterContext
    {
        /** this XML file's parsed DOM */
        public dom: D;
        /** a DOM builder for this XML document */
        public domBldr: DomBuilderFactory<D>;
        /** read/write XLSX DOM element utility functions */
        public readMulti: OpenXmlIo.ElementsReader;
        public writeMulti: OpenXmlIo.ElementsWriter;
        /** a validator for XLSX DOM elements */
        public validator: DomValidate;

        constructor(dom: D) {
            super(dom, XlsxDomErrorsImpl);
            this.dom = dom;
            this.domBldr = new DomBuilderFactory(dom);
            this.validator = XlsxDomErrorsImpl;
            this.readMulti = <T>(reader: OpenXmlIo.ReadFunc<T> | OpenXmlIo.ReadFuncNamed<T>, elems: HTMLElement[], expectedElemName?: string): T[] => XmlFileInst.readMulti(this, reader, elems, expectedElemName);
            this.writeMulti = <T>(writer: OpenXmlIo.WriteFunc<T> | OpenXmlIo.WriteFuncNamed<T>, insts: T[] | { [id: string]: T }, keysOrExpectedElemName?: string | string[]): ElementLike[] => XmlFileInst.writeMulti(this, writer, insts, keysOrExpectedElemName);
        }

    }


    export function newInst(dom: XMLDocument): XmlDocFile<XMLDocument>;
    export function newInst<D extends DocumentLike>(dom: D): XmlDocFile<DocumentLike>;
    export function newInst<D extends DocumentLike>(dom: D) {
        return new XmlDocFile<D>(dom);
    }


    /** Provides generic logic for reading/writing an array of OpenXml elements using a reader/writer for a single element of the same type
     */

    export function readMulti<T>(xmlDoc: OpenXmlIo.ReaderContext, reader: OpenXmlIo.ReadFunc<T> | OpenXmlIo.ReadFuncNamed<T>, elems: HTMLElement[], expectedElemName?: string): T[] {
        var res: T[] = [];
        for (var i = 0, size = elems.length; i < size; i++) {
            var elem = elems[i];
            res.push((<OpenXmlIo.ReadFunc<T>>reader)(xmlDoc, elem, expectedElemName));
        }
        return res;
    }


    export function writeMulti<T>(xmlDoc: OpenXmlIo.WriterContext, writer: OpenXmlIo.WriteFuncNamed<T>, insts: T[], expectedElemName?: string): ElementLike[];
    export function writeMulti<T>(xmlDoc: OpenXmlIo.WriterContext, writer: OpenXmlIo.WriteFunc<T>, insts: { [id: string]: T }, keys?: string[]): ElementLike[];
    export function writeMulti<T>(xmlDoc: OpenXmlIo.WriterContext, writer: OpenXmlIo.WriteFunc<T> | OpenXmlIo.WriteFuncNamed<T>, insts: T[] | { [id: string]: T }, keysOrExpectedElemName?: string | string[]): ElementLike[];
    export function writeMulti<T>(xmlDoc: OpenXmlIo.WriterContext, writer: OpenXmlIo.WriteFunc<T> | OpenXmlIo.WriteFuncNamed<T>, insts: T[] | { [id: string]: T }, keysOrExpectedElemName?: string | string[]): ElementLike[] {
        var res: ElementLike[] = [];
        if (Array.isArray(keysOrExpectedElemName)) {
            var keys = keysOrExpectedElemName;
            for (var i = 0, size = keys.length || (<T[]>insts).length; i < size; i++) {
                var inst = <T>(<any>insts)[keys[i]];
                res.push((<OpenXmlIo.WriteFunc<T>>writer)(xmlDoc, inst));
            }
        }
        else {
            var expectedElemName = keysOrExpectedElemName;
            for (var i = 0, size = (<T[]>insts).length; i < size; i++) {
                var inst = <T>(<any>insts)[i];
                res.push((<OpenXmlIo.WriteFunc<T>>writer)(xmlDoc, inst, expectedElemName));
            }
        }
        return res;
    }

}