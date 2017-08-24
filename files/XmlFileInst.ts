import DomBuilderFactory = require("../../dom-builder/dom/DomBuilderFactory");
import DomBuilderHelper = require("../../dom-builder/dom/DomBuilderHelper");
import XlsxDomErrorsImpl = require("../errors/XlsxDomErrorsImpl");

/** Implementation of OpenXmlIo.ParsedFile, contains:
 * - An XMLDocument containing the file data
 * - A DomBuilderFactory for creating new DOM elements when writing data back to the file
 * - An XlsxDom utility object with methods to make reading/writing data to DOM elements easier
 * - An XlsxDomValidator for checking that DOM elements match expected formats and throwing errors when not
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
module XmlFileInst {

    export class DocLikeFile extends DomBuilderHelper implements OpenXmlIo.ReaderContext, OpenXmlIo.WriterContext {
        /** this XML file's parsed DOM */
        public dom: DocumentLike;
        /** a DOM builder for this XML document */
        public domBldr: DomBuilderFactory<DocumentLike>;
        /** read/write XLSX DOM element utility functions */
        public readMulti: OpenXmlIo.ElementsReader;
        public writeMulti: OpenXmlIo.ElementsWriter;
        /** a validator for XLSX DOM elements */
        public validator: DomValidate;

        constructor(dom: DocumentLike) {
            super(dom, XlsxDomErrorsImpl);
        }

    }


    export class XmlDocFile extends DomBuilderHelper implements OpenXmlIo.ReaderContext, OpenXmlIo.WriterContext {
        /** this XML file's parsed DOM */
        public dom: XMLDocument;
        /** a DOM builder for this XML document */
        public domBldr: DomBuilderFactory<XMLDocument>;
        /** read/write XLSX DOM element utility functions */
        public readMulti: OpenXmlIo.ElementsReader;
        public writeMulti: OpenXmlIo.ElementsWriter;
        /** a validator for XLSX DOM elements */
        public validator: DomValidate;

        constructor(dom: XMLDocument) {
            super(dom, XlsxDomErrorsImpl);
        }

    }


    export function newInst(dom: XMLDocument): XmlDocFile;
    export function newInst<D extends DocumentLike>(dom: D): DocLikeFile;
    export function newInst<D extends DocumentLike>(dom: D) {
        var inst = ((<any>dom).childNodes != null ? new XmlDocFile(<XMLDocument><any>dom) : new DocLikeFile(dom));
        inst.dom = dom;
        inst.domBldr = new DomBuilderFactory<D>(dom);

        inst.readMulti = function readMulti<T>(this: OpenXmlIo.ReaderContext, reader: (xmlDoc: OpenXmlIo.ReaderContext, elem: HTMLElement, expectedElemName?: string) => T, elems: HTMLElement[], expectedElemName?: string) {
            return XmlFileInst.readMulti(this, reader, elems, expectedElemName);
        };
        inst.writeMulti = function writeMulti<T, E extends ElementLike>(this: OpenXmlIo.WriterContext, writer: (xmlDoc: OpenXmlIo.WriterContext, data: T, expectedElemName?: string) => E, insts: T[] | { [id: string]: T }, keysOrExpectedElemName?: string | string[]) {
            return XmlFileInst.writeMulti(this, writer, <any>insts, <any>keysOrExpectedElemName);
        };

        inst.validator = XlsxDomErrorsImpl;
        return inst;
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
    export function writeMulti<T>(xmlDoc: OpenXmlIo.WriterContext, writer: OpenXmlIo.WriteFunc<T>, insts: { [id: string]: T }, keys?: string[]): ElementLike[]
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

export = XmlFileInst;