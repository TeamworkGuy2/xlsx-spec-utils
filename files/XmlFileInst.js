"use strict";
var DomBuilderFactory = require("../../ts-dom-builder/dom/DomBuilderFactory");
var DomBuilderHelper = require("../..//ts-dom-builder/dom/DomBuilderHelper");
var XlsxDomErrorsImpl = require("../errors/XlsxDomErrorsImpl");
/** Implementation of OpenXmlIo.ParsedFile, contains:
 * - An XMLDocument containing the file data
 * - A DomBuilderFactory for creating new DOM elements when writing data back to the file
 * - An XlsxDom utility object with methods to make reading/writing data to DOM elements easier
 * - An XlsxDomValidator for checking that DOM elements match expected formats and throwing errors when not
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
var XmlFileInst = (function () {
    function XmlFileInst(dom) {
        this.dom = dom;
        this.domBldr = new DomBuilderFactory(dom);
        this.domHelper = new DomBuilderHelper(dom, XlsxDomErrorsImpl);
        var readWriteElements = new XmlFileInst.ReadWriteOpenXmlElementImpl();
        this.readOpenXml = readWriteElements;
        this.writeOpenXml = readWriteElements;
        this.validator = XlsxDomErrorsImpl;
    }
    XmlFileInst.newInst = function (dom) {
        return new XmlFileInst(dom);
    };
    return XmlFileInst;
}());
var XmlFileInst;
(function (XmlFileInst) {
    /** Provides generic logic for reading/writing an array of OpenXml elements using a reader/writer for a single element of the same type
     */
    var ReadWriteOpenXmlElementImpl = (function () {
        function ReadWriteOpenXmlElementImpl() {
        }
        ReadWriteOpenXmlElementImpl.prototype.readMulti = function (xmlDoc, reader, elems, expectedElemName) {
            var res = [];
            for (var i = 0, size = elems.length; i < size; i++) {
                var elem = elems[i];
                res.push(reader(xmlDoc, elem, expectedElemName));
            }
            return res;
        };
        ReadWriteOpenXmlElementImpl.prototype.writeMulti = function (xmlDoc, writer, insts, keysOrExpectedElemName) {
            var res = [];
            if (Array.isArray(keysOrExpectedElemName)) {
                for (var i = 0, size = keysOrExpectedElemName.length || insts.length; i < size; i++) {
                    var inst = insts[keysOrExpectedElemName[i]];
                    res.push(writer(xmlDoc, inst));
                }
            }
            else {
                for (var i = 0, size = insts.length; i < size; i++) {
                    var inst = insts[i];
                    res.push(writer(xmlDoc, inst, keysOrExpectedElemName));
                }
            }
            return res;
        };
        return ReadWriteOpenXmlElementImpl;
    }());
    XmlFileInst.ReadWriteOpenXmlElementImpl = ReadWriteOpenXmlElementImpl;
})(XmlFileInst || (XmlFileInst = {}));
module.exports = XmlFileInst;
