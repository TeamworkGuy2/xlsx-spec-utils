import DomBuilderHelper = require("../../ts-dom-builder/dom/DomBuilderHelper");
import XmlFileInst = require("./XmlFileInst");

/** This object implementation instantiates a new factory every time read() function called, usage: read()/write() pairs
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
class XmlFileReadWriter<T> implements OpenXmlIo.FileReadWriter<T> {
    public fileInfo: OpenXmlIo.XlsxFileType;
    private prepForWrite: (xmlDoc: OpenXmlIo.ParsedFile, inst: T) => void;
    private rootReadWriter: OpenXmlIo.ReadWrite<T>;
    private lastReadXmlDoc: OpenXmlIo.ParsedFile;


    constructor(fileInfo: OpenXmlIo.XlsxFileType, rootReadWriter: OpenXmlIo.ReadWrite<T>, prepForWrite: (xmlDoc: OpenXmlIo.ParsedFile, inst: T) => void) {
        this.fileInfo = fileInfo;
        this.rootReadWriter = rootReadWriter;
        this.prepForWrite = prepForWrite;
    }


    public read(xmlContentStr: string): T {
        var dom = XmlFileReadWriter.xmlTextToDom(xmlContentStr);
        return this.loadFromDom(dom);
    }


    public write(inst: T): string {
        var dom = this.saveToDom(inst);
        return XmlFileReadWriter.domToXmlText(dom);
    }


    public loadFromDom(dom: Document): T {
        var xmlDoc = XmlFileInst.newInst(dom);
        this.lastReadXmlDoc = xmlDoc;

        var domRoot = <HTMLElement>xmlDoc.dom.childNodes[0];
        return this.rootReadWriter.read(xmlDoc, domRoot);
    }


    public saveToDom(inst: T): Document {
        var xmlDoc = this.lastReadXmlDoc;
        this.prepForWrite(xmlDoc, inst);
        var elemDom = <HTMLElement>xmlDoc.dom.childNodes[0];

        var elem = this.rootReadWriter.write(xmlDoc, inst);
        xmlDoc.domHelper.addChilds(elemDom, xmlDoc.domHelper.getChilds(elem));

        return xmlDoc.dom;
    }


    public static domToXmlText(dom: Document): string {
        return DomBuilderHelper.getSerializer().serializeToString(dom);
    }

    
    public static xmlTextToDom(xmlStr: string): Document {
        var dom = DomBuilderHelper.getParser().parseFromString(xmlStr, "application/xml");
        return dom;
    }

}

export = XmlFileReadWriter;