"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.XmlFileReadWriter = void 0;
var DomBuilderHelper_1 = require("@twg2/dom-builder/dom/DomBuilderHelper");
var XmlFileInst_1 = require("./XmlFileInst");
/**
 * An {@link OpenXmlIo.FileReadWriter} implementation with a configurable pre-write callback and
 * a cache containing the last read()/loadFromDom() result.
 * Internally this uses {@link DomBuilderHelper}'s `getParser()` and `getSerializer()` to
 * read and write XML.
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
var XmlFileReadWriter = /** @class */ (function () {
    /** Create an XML file reader/writer
     * @param fileInfo information about the OpenXML XLSX file type that this reader/writer handles
     * @param rootReadWriter the reader/writer that performs the serialization and parsing of DOM elements
     * @param prepForWrite a function which is called with the previous read()/loadFromDom() result before new DOM nodes are added by write()/saveToDom().
     * A common use for the function is to clear the DOM and add default schema/root meta-data elements and attributes about the XLSX file type.
     */
    function XmlFileReadWriter(fileInfo, rootReadWriter, prepForWrite) {
        this.fileInfo = fileInfo;
        this.rootReadWriter = rootReadWriter;
        this.prepForWrite = prepForWrite;
    }
    /** Parses an XML string and passes the resulting Document to {@link loadFromDom} which creates
     * a new ParsedFile instance with the DOM data and parses it into an object
     * @param xmlContentStr the XML string to parse
     * @return the data object returned by rootReadWriter.read() given the DOM parsed from 'xmlContentStr'
     */
    XmlFileReadWriter.prototype.read = function (xmlContentStr) {
        var dom = XmlFileReadWriter.xmlTextToDom(xmlContentStr);
        return this.loadFromDom(dom);
    };
    /** Calls {@link saveToDom} to write 'data' to the last DOM loaded by read()/loadFromDom()
     * and then serializes the DOM to XML.
     * @param data the data to write
     * @return XML document with XML declaration serialized from the `rootReadWriter.write()` {@link Element}(s)
     */
    XmlFileReadWriter.prototype.write = function (data) {
        var dom = this.saveToDom(data);
        return XmlFileReadWriter.domToXmlText(dom, true, this.fileInfo);
    };
    /** Creates a new OpenXmlIo.ParsedFile instance to hold DOM data, calls the `rootReadWriter` read() method and returns the result
     * @param dom the Document to read
     * @param namespaceURI optional namespace to assign to elements created using this instance's
     * `lastReadXmlDoc.domBldr` property. This allows callers to control the default namespace of
     * all elements created when `saveToDom()` is called and the `rootReadWriter.write()` uses
     * `domBldr.create()` to create new elements.
     * If provided, elements will be created using `dom.createElementNS()` instead of `dom.createElement()`.
     * @return the data object returned by rootReadWriter.read() given the 'dom' parameter
     */
    XmlFileReadWriter.prototype.loadFromDom = function (dom, namespaceURI) {
        var xmlDoc = XmlFileInst_1.XmlFileInst.newInst(dom, namespaceURI);
        this.lastReadXmlDoc = xmlDoc;
        var domRoot = xmlDoc.dom.childNodes[0];
        return this.rootReadWriter.read(xmlDoc, domRoot);
    };
    /** Write data into the last DOM result loaded by read()/loadFromDom().
     * The `prepForWrite` function is called first with the last DOM result.
     * Then the `rootReadWriter` write() method is called to write the 'data' parameter to an DOM {@link Document}.
     * Finally the child elements of the write() result are added to the last DOM result and the last DOM result is returned
     * @param data the data to convert to {@link Element}(s)
     * @return the last DOM result loaded by read()/loadFromDom() with additional elements created by rootReadWriter based on the 'data' parameter provided
     */
    XmlFileReadWriter.prototype.saveToDom = function (data) {
        // TODO: should create a new XML Document rather than reusing last read Document,
        // New document should include an XML declaration via `doc.createProcessingInstruction('xml', 'version="1.0" encoding="UTF-8"')`
        // https://stackoverflow.com/questions/68801002/add-xml-declaration-to-xml-document-programmatically
        var xmlDoc = this.lastReadXmlDoc;
        if (xmlDoc == null) {
            throw new Error("Must call loadFromDom() before saveToDom()");
        }
        this.prepForWrite(xmlDoc, data);
        var elem = this.rootReadWriter.write(xmlDoc, data);
        var elemDom = xmlDoc.dom.childNodes[0];
        xmlDoc.addChilds(elemDom, xmlDoc.getChildNodes(elem));
        return xmlDoc.dom;
    };
    /** Convert a DOM node (can be an entire document or a subtree) to a string.
     * This adds an XML declaration if the 'dom' has none.
     * Uses {@link DomBuilderHelper}'s `getSerializer()` to serialize the DOM node.
     * @param dom the DOM node to serialize
     * @param includeXmlDeclaration optional flag, if false no XML declaration, i.e. `<?xml version="1.0" encoding="UTF-8" [standalone="yes"]?>`
     * will be included at the beginning of the returned XML.
     * If undefined or true, an XML declaration will be added if missing.
     * @param fileType optional file type information. If the `fileInfo.xlsxFilePath` ends with '.rels' then
     * a `standalone="yes"` attribute is included in the optional XML declaration.
     * @returns the serialized 'dom' XML
     */
    XmlFileReadWriter.domToXmlText = function (dom, includeXmlDeclaration, fileType) {
        var _a;
        var xml = DomBuilderHelper_1.DomBuilderHelper.getSerializer().serializeToString(dom);
        if (includeXmlDeclaration !== false && !xml.startsWith("<?xml")) {
            xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"".concat(fileType == null || !((_a = fileType === null || fileType === void 0 ? void 0 : fileType.xlsxFilePath) === null || _a === void 0 ? void 0 : _a.endsWith(".rels")) ? " standalone=\"yes\"" : '', "?>") + xml;
        }
        return xml;
    };
    /** Convert an XML string into a DOM document.
     * Uses {@link DomBuilderHelper}'s `getParser()` to parse the xml string.
     * @param xmlStr the XML string to parse
     */
    XmlFileReadWriter.xmlTextToDom = function (xmlStr) {
        var dom = DomBuilderHelper_1.DomBuilderHelper.getParser().parseFromString(xmlStr, "application/xml");
        return dom;
    };
    return XmlFileReadWriter;
}());
exports.XmlFileReadWriter = XmlFileReadWriter;
