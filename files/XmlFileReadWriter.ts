import DomBuilderHelper = require("@twg2/dom-builder/dom/DomBuilderHelper");
import XmlFileInst = require("./XmlFileInst");

/** An OpenXmlIo FileReadWriter implementation with a configurable pre-write callback and a cache containing the last read()/loadFromDom() result
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
class XmlFileReadWriter<T> implements OpenXmlIo.FileReadWriter<T> {
    public fileInfo: OpenXmlIo.XlsxFileType;
    private prepForWrite: (xmlDoc: OpenXmlIo.WriterContext, inst: T) => void;
    private rootReadWriter: OpenXmlIo.ReadWrite<T>;
    private lastReadXmlDoc: XmlFileInst.XmlDocFile | undefined;


    /** Create an XML file reader/writer
     * @param fileInfo information about the OpenXML XLSX file type that this reader/writer handles
     * @param rootReadWriter the reader/writer that performs the serialization and parsing of DOM elements
     * @param prepForWrite a function which is called with the previous read()/loadFromDom() result before new DOM nodes are added by write()/saveToDom().
     * A common use for the function is to clear the DOM and add default schema/root meta-data elements and attributes about the XLSX file type.
     */
    constructor(fileInfo: OpenXmlIo.XlsxFileType, rootReadWriter: OpenXmlIo.ReadWrite<T>, prepForWrite: (xmlDoc: OpenXmlIo.WriterContext, inst: T) => void) {
        this.fileInfo = fileInfo;
        this.rootReadWriter = rootReadWriter;
        this.prepForWrite = prepForWrite;
    }


    /** Parses an XML string and passes the resulting Document to loadFromDom() which creates a new ParsedFile instance with the DOM data and parses it into an object
     * @param xmlContentStr the XML string to parse
     * @return the data object returned by rootReadWriter.read() given the DOM parsed from 'xmlContentStr'
     */
    public read(xmlContentStr: string): T {
        var dom = XmlFileReadWriter.xmlTextToDom(xmlContentStr);
        return this.loadFromDom(dom);
    }


    /** Calls saveToDom(data) which writes 'data' to the last DOM result loaded by read()/loadFromDom() and then serializes the resulting DOM to text.
     * @param data the data to write
     * @return the XML string from serializing the HTMLElements created by rootReadWriter.write()
     */
    public write(data: T): string {
        var dom = this.saveToDom(data);
        return XmlFileReadWriter.domToXmlText(dom);
    }


    /** Creates a new OpenXmlIo.ParsedFile instance to hold DOM data, calls the 'rootReadWriter' read() method and returns the result
     * @param dom the Document to read
     * @return the data object returned by rootReadWriter.read() given the 'dom' parameter
     */
    public loadFromDom(dom: Document): T {
        var xmlDoc = XmlFileInst.newInst(dom);
        this.lastReadXmlDoc = xmlDoc;

        var domRoot = <HTMLElement>xmlDoc.dom.childNodes[0];
        return this.rootReadWriter.read(xmlDoc, domRoot);
    }


    /** Write data into the last DOM result loaded by read()/loadFromDom().
     * The 'prepForWrite' function is called first with the last DOM result.
     * Then the 'rootReadWriter' write() method is called to write the 'data' parameter to an HTMLElement subtree.
     * Finally the child elements of the write() result are added to the last DOM result and the last DOM result is returned
     * @param data the data to convert to HTMLElement(s)
     * @return the last DOM result loaded by read()/loadFromDom() with additional elements created by rootReadWriter based on the 'data' parameter provided
     */
    public saveToDom(data: T): Document {
        var xmlDoc = this.lastReadXmlDoc;
        if (xmlDoc == null) {
            throw new Error("Must call loadFromDom() before saveToDom()");
        }
        this.prepForWrite(xmlDoc, data);
        var elem = this.rootReadWriter.write(xmlDoc, data);

        var elemDom = <HTMLElement>(<Document>xmlDoc.dom).childNodes[0];
        xmlDoc.addChilds(elemDom, xmlDoc.getChildNodes(elem));

        return xmlDoc.dom;
    }


    /** Convert a DOM node (can be an entire document or a subtree) to a string
     * @param dom the DOM node to serialize
     */
    public static domToXmlText(dom: Node): string {
        return DomBuilderHelper.getSerializer().serializeToString(dom);
    }


    /** Convert an XML string into a DOM document
     * @param xmlStr the XML string to parse
     */
    public static xmlTextToDom(xmlStr: string): Document {
        var dom = DomBuilderHelper.getParser().parseFromString(xmlStr, "application/xml");
        return dom;
    }

}

export = XmlFileReadWriter;