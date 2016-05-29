"use strict";
var DomBuilderHelper = require("../../ts-dom-builder/dom/DomBuilderHelper");
var XmlFileInst = require("./XmlFileInst");
/** This object implementation instantiates a new factory every time read() function called, usage: read()/write() pairs
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
var XmlFileReadWriter = (function () {
    function XmlFileReadWriter(fileInfo, rootReadWriter, prepForWrite) {
        this.fileInfo = fileInfo;
        this.rootReadWriter = rootReadWriter;
        this.prepForWrite = prepForWrite;
    }
    XmlFileReadWriter.prototype.read = function (xmlContentStr) {
        var dom = XmlFileReadWriter.xmlTextToDom(xmlContentStr);
        return this.loadFromDom(dom);
    };
    XmlFileReadWriter.prototype.write = function (inst) {
        var dom = this.saveToDom(inst);
        return XmlFileReadWriter.domToXmlText(dom);
    };
    XmlFileReadWriter.prototype.loadFromDom = function (dom) {
        var xmlDoc = XmlFileInst.newInst(dom);
        this.lastReadXmlDoc = xmlDoc;
        var domRoot = xmlDoc.dom.childNodes[0];
        return this.rootReadWriter.read(xmlDoc, domRoot);
    };
    XmlFileReadWriter.prototype.saveToDom = function (inst) {
        var xmlDoc = this.lastReadXmlDoc;
        this.prepForWrite(xmlDoc, inst);
        var elemDom = xmlDoc.dom.childNodes[0];
        var elem = this.rootReadWriter.write(xmlDoc, inst);
        xmlDoc.domHelper.addChilds(elemDom, xmlDoc.domHelper.getChilds(elem));
        return xmlDoc.dom;
    };
    XmlFileReadWriter.domToXmlText = function (dom) {
        return DomBuilderHelper.getSerializer().serializeToString(dom);
    };
    XmlFileReadWriter.xmlTextToDom = function (xmlStr) {
        var dom = DomBuilderHelper.getParser().parseFromString(xmlStr, "application/xml");
        return dom;
    };
    return XmlFileReadWriter;
}());
module.exports = XmlFileReadWriter;
