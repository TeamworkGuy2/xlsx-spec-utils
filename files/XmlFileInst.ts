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
class XmlFileInst implements OpenXmlIo.ParsedFile {
    /** this XML file's parsed DOM */
    public dom: XMLDocument;
    /** a DOM builder for this XML document */
    public domBldr: DomBuilderFactory;
    /** a XLSX DOM manipulation utility */
    public domHelper: DomBuilderHelper;
    /** utlities for reading XML child elements */
    public readOpenXml: OpenXmlIo.ReadOpenXmlElement;
    /** utlities for writing XML child elements */
    public writeOpenXml: OpenXmlIo.WriteOpenXmlElement;
    /** a validator for XLSX DOM elements */
    public validator: DomValidate;


    constructor(dom: XMLDocument) {
        this.dom = dom;
        this.domBldr = new DomBuilderFactory(dom);
        this.domHelper = new DomBuilderHelper(dom, XlsxDomErrorsImpl);

        var readWriteElements = new XmlFileInst.ReadWriteOpenXmlElementImpl();
        this.readOpenXml = readWriteElements;
        this.writeOpenXml = readWriteElements;

        this.validator = XlsxDomErrorsImpl;
    }


    public static newInst(dom: XMLDocument) {
        return new XmlFileInst(dom);
    }

}

module XmlFileInst {


    /** Provides generic logic for reading/writing an array of OpenXml elements using a reader/writer for a single element of the same type
     */
    export class ReadWriteOpenXmlElementImpl implements OpenXmlIo.ReadOpenXmlElement, OpenXmlIo.WriteOpenXmlElement {

        public readMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, reader: OpenXmlIo.ReadFunc<T> | OpenXmlIo.ReadFuncNamed<T>, elems: HTMLElement[], expectedElemName?: string): T[] {
            var res: T[] = [];
            for (var i = 0, size = elems.length; i < size; i++) {
                var elem = elems[i];
                res.push((<OpenXmlIo.ReadFunc<T>>reader)(xmlDoc, elem, expectedElemName));
            }
            return res;
        }


        public writeMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, writer: OpenXmlIo.WriteFuncNamed<T>, insts: T[], expectedElemName?: string): HTMLElement[];
        public writeMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, writer: OpenXmlIo.WriteFunc<T>, insts: { [id: string]: T }, keys?: string[]): HTMLElement[]
        public writeMulti<T>(xmlDoc: OpenXmlIo.ParsedFile, writer: OpenXmlIo.WriteFunc<T> | OpenXmlIo.WriteFuncNamed<T>, insts: T[] | { [id: string]: T }, keysOrExpectedElemName?: string | string[]): HTMLElement[] {
            var res: HTMLElement[] = [];
            if (Array.isArray(keysOrExpectedElemName)) {
                var keys = keysOrExpectedElemName;
                for (var i = 0, size = keys.length || (<T[]>insts).length; i < size; i++) {
                    var inst = <T>insts[keys[i]];
                    res.push((<OpenXmlIo.WriteFunc<T>>writer)(xmlDoc, inst));
                }
            }
            else {
                var expectedElemName = keysOrExpectedElemName;
                for (var i = 0, size = (<T[]>insts).length; i < size; i++) {
                    var inst = <T>insts[i];
                    res.push((<OpenXmlIo.WriteFunc<T>>writer)(xmlDoc, inst, expectedElemName));
                }
            }
            return res;
        }

    }

}

export = XmlFileInst;