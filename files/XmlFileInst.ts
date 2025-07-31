import { DomBuilderFactory } from "@twg2/dom-builder/dom/DomBuilderFactory";
import { DomBuilderHelper } from "@twg2/dom-builder/dom/DomBuilderHelper";
import { DomLite } from "@twg2/dom-builder/dom/DomLite";
import { XlsxDomErrorsImpl } from "../errors/XlsxDomErrorsImpl";
import * as XlsxNamespace from "./XlsxNamespace";

/** Implementation of {@link OpenXmlIo.ReaderContext} and {@link OpenXmlIo.WriterContext}, contains:
 * - An XMLDocument containing the file data
 * - A DomBuilderFactory for creating new DOM elements when writing data back to the file
 * - An XlsxDom utility object with methods to make reading/writing data to DOM elements easier
 * - An XlsxDomValidator for checking that DOM elements match expected formats and throwing errors when not
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
export class XmlFileInst<D extends DocumentLike = DocumentLike>
    extends DomBuilderHelper
    implements OpenXmlIo.ReaderContext, OpenXmlIo.WriterContext
{
    /** this XML file's parsed DOM */
    public dom: D;
    /** a DOM builder for this XML document */
    public domBldr: DomBuilderFactory;
    /** read/write XLSX DOM element utility functions */
    public readMulti: OpenXmlIo.ElementsReader;
    public writeMulti: OpenXmlIo.ElementsWriter;
    /** a validator for XLSX DOM elements */
    public validator: DomValidate;

    /**
     * Create a XML file instance backed by a {@link XMLDocument}
     * @param dom a {@link DocumentLike} object that this XML file will read from and write to
     * @param namespaceURI optional namespace to assign to elements created by this instance's
     * `domBldr` property. If provided, elements will be created using `dom.createElementNS()`.
     */
    constructor(
        dom: D,
        namespaceURI?: string | null,
        attributeNamespaceHandler?: ((elem: ElementLike, qualifiedName: string) => string | null) | undefined
    ) {
        super(dom, XlsxDomErrorsImpl);
        this.dom = dom;
        // custom handling for attribute namespaces in OpenXML files
        this.domBldr = new DomBuilderFactory(dom, namespaceURI, attributeNamespaceHandler);
        this.validator = XlsxDomErrorsImpl;
        this.readMulti = <T>(reader: OpenXmlIo.ReadFunc<T> | OpenXmlIo.ReadFuncNamed<T>, elems: HTMLElement[], expectedElemName?: string): T[] => XmlFileInst.readMulti(this, reader, elems, expectedElemName);
        this.writeMulti = <T>(writer: OpenXmlIo.WriteFunc<T> | OpenXmlIo.WriteFuncNamed<T>, insts: T[] | { [id: string]: T }, keysOrExpectedElemName?: string | string[]): ElementLike[] => XmlFileInst.writeMulti(this, writer, insts, keysOrExpectedElemName);
    }

    /**
     * Create a XML file instance backed by a {@link XMLDocument}
     * @param dom a {@link DocumentLike} object that this XML file will read from and write to
     * @param namespaceURI optional namespace to assign to elements created by the instance's
     * `domBldr` property. If provided, elements will be created using `dom.createElementNS()`.
     */
    public static newInst(dom: XMLDocument, namespaceURI?: string | null): XmlFileInst<XMLDocument>;
    public static newInst<D extends DocumentLike>(dom: D, namespaceURI?: string | null): XmlFileInst<DocumentLike>;
    public static newInst<D extends DocumentLike>(dom: D, namespaceURI?: string | null) {
        return new XmlFileInst<D>(dom, namespaceURI, (elem, name) => XmlFileInst.lookupAndAddNamespace(dom, elem, name));
    }

    /** Logic for reading/writing an array of OpenXml elements using a reader/writer for a single element of the same type
     */

    public static readMulti<T>(xmlDoc: OpenXmlIo.ReaderContext, reader: OpenXmlIo.ReadFunc<T> | OpenXmlIo.ReadFuncNamed<T>, elems: HTMLElement[], expectedElemName?: string): T[] {
        var res: T[] = [];
        for (var i = 0, size = elems.length; i < size; i++) {
            var elem = elems[i];
            res.push((<OpenXmlIo.ReadFunc<T>>reader)(xmlDoc, elem, expectedElemName));
        }
        return res;
    }

    public static writeMulti<T>(xmlDoc: OpenXmlIo.WriterContext, writer: OpenXmlIo.WriteFuncNamed<T>, insts: T[], expectedElemName?: string): ElementLike[];
    public static writeMulti<T>(xmlDoc: OpenXmlIo.WriterContext, writer: OpenXmlIo.WriteFunc<T>, insts: { [id: string]: T }, keys?: string[]): ElementLike[];
    public static writeMulti<T>(xmlDoc: OpenXmlIo.WriterContext, writer: OpenXmlIo.WriteFunc<T> | OpenXmlIo.WriteFuncNamed<T>, insts: T[] | { [id: string]: T }, keysOrExpectedElemName?: string | string[]): ElementLike[];
    public static writeMulti<T>(xmlDoc: OpenXmlIo.WriterContext, writer: OpenXmlIo.WriteFunc<T> | OpenXmlIo.WriteFuncNamed<T>, insts: T[] | { [id: string]: T }, keysOrExpectedElemName?: string | string[]): ElementLike[] {
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

    /**
     * Logic to pick a namespaceURI for an attribute based on its qualifying name.
     * And add that namespace to the root of the document if necessary.
     * The prefix is examined and 'xml:' returns the default XML namespace URI.
     * The DOM's `lookupNamespaceURI(prefix)` is called and if it returns a namespace, that is used.
     * Recognized Open XML prefixes like 'r' and 'x14ac' return the corresponding schema URI.
     * @see XlsxNamespace
     * @param element the element that the attribute is being set on
     * @param qualifiedName the qualified name of the attribute, including prefix, such as 'r:id' or 'xml:space'
     * @returns the corresponding namespace URI or null if no match is found
     */
    public static lookupAndAddNamespace(document: DocumentLike, element: ElementLike, qualifiedName: string): string | null {
        const colonIdx = qualifiedName.indexOf(':');
        let namespaceUri: string | null = null;
        if (qualifiedName.startsWith('xml:')) {
            namespaceUri = DomLite.XML_NAMESPACE;
        }
        if (namespaceUri == null) {
            namespaceUri = document.lookupNamespaceURI(qualifiedName.substring(0, colonIdx));
        }
        const documentElement = document.documentElement;
        const prefix = qualifiedName.substring(0, colonIdx);
        if (namespaceUri == null && prefix != null) {
            namespaceUri = XlsxNamespace.openxmlNamespaces[prefix];
            // If an OpenXML 'additional' namespace is used for an attribute, add it to the root of the document
            if (namespaceUri != null) {
                documentElement.setAttributeNS(DomLite.XMLNS_NAMESPACE, `xmlns:${prefix}`, namespaceUri);
            }
        }
        if (namespaceUri == null && prefix != null) {
            namespaceUri = XlsxNamespace.xlsxAdditionalNamespaces[prefix];
            // If an OpenXML 'additional' namespace is used for an attribute, add it and ignore it on the
            // root of the document since this is how OpenXML files handle namespaces
            if (namespaceUri != null) {
                documentElement.setAttributeNS(DomLite.XMLNS_NAMESPACE, `xmlns:${prefix}`, namespaceUri);
                // set the 'mc' namespace
                const ignorableNsUri = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
                documentElement.setAttributeNS(DomLite.XMLNS_NAMESPACE, 'xmlns:mc', ignorableNsUri);
                // add the new ignorable prefi to the existing 'mc:Ignorable' prefixes list (if present)
                const rootAttrs = documentElement.attributes;
                const prefixesAttr = rootAttrs.getNamedItemNS(ignorableNsUri, 'Ignorable');
                const prefixes = `${prefix}${prefixesAttr?.value ? ` ${prefixesAttr.value}` : ''}`;
                documentElement.setAttributeNS(ignorableNsUri, 'mc:Ignorable', prefixes);
            }
        }
        if (namespaceUri == null) {
            namespaceUri = element.namespaceURI ?? null;
        }
        return namespaceUri;
    }
}