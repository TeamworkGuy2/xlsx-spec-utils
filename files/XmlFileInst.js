"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.XmlFileInst = void 0;
var DomBuilderFactory_1 = require("@twg2/dom-builder/dom/DomBuilderFactory");
var DomBuilderHelper_1 = require("@twg2/dom-builder/dom/DomBuilderHelper");
var XlsxDomErrorsImpl_1 = require("../errors/XlsxDomErrorsImpl");
/** Implementation of OpenXmlIo.ParsedFile, contains:
 * - An XMLDocument containing the file data
 * - A DomBuilderFactory for creating new DOM elements when writing data back to the file
 * - An XlsxDom utility object with methods to make reading/writing data to DOM elements easier
 * - An XlsxDomValidator for checking that DOM elements match expected formats and throwing errors when not
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
var XmlFileInst;
(function (XmlFileInst) {
    var XmlDocFile = /** @class */ (function (_super) {
        __extends(XmlDocFile, _super);
        function XmlDocFile(dom) {
            var _this = _super.call(this, dom, XlsxDomErrorsImpl_1.XlsxDomErrorsImpl) || this;
            _this.dom = dom;
            _this.domBldr = new DomBuilderFactory_1.DomBuilderFactory(dom);
            _this.validator = XlsxDomErrorsImpl_1.XlsxDomErrorsImpl;
            _this.readMulti = function (reader, elems, expectedElemName) { return XmlFileInst.readMulti(_this, reader, elems, expectedElemName); };
            _this.writeMulti = function (writer, insts, keysOrExpectedElemName) { return XmlFileInst.writeMulti(_this, writer, insts, keysOrExpectedElemName); };
            return _this;
        }
        return XmlDocFile;
    }(DomBuilderHelper_1.DomBuilderHelper));
    XmlFileInst.XmlDocFile = XmlDocFile;
    function newInst(dom) {
        return new XmlDocFile(dom);
    }
    XmlFileInst.newInst = newInst;
    /** Provides generic logic for reading/writing an array of OpenXml elements using a reader/writer for a single element of the same type
     */
    function readMulti(xmlDoc, reader, elems, expectedElemName) {
        var res = [];
        for (var i = 0, size = elems.length; i < size; i++) {
            var elem = elems[i];
            res.push(reader(xmlDoc, elem, expectedElemName));
        }
        return res;
    }
    XmlFileInst.readMulti = readMulti;
    function writeMulti(xmlDoc, writer, insts, keysOrExpectedElemName) {
        var res = [];
        if (Array.isArray(keysOrExpectedElemName)) {
            var keys = keysOrExpectedElemName;
            for (var i = 0, size = keys.length || insts.length; i < size; i++) {
                var inst = insts[keys[i]];
                res.push(writer(xmlDoc, inst));
            }
        }
        else {
            var expectedElemName = keysOrExpectedElemName;
            for (var i = 0, size = insts.length; i < size; i++) {
                var inst = insts[i];
                res.push(writer(xmlDoc, inst, expectedElemName));
            }
        }
        return res;
    }
    XmlFileInst.writeMulti = writeMulti;
})(XmlFileInst = exports.XmlFileInst || (exports.XmlFileInst = {}));
