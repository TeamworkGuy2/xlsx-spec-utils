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
var DomLite_1 = require("@twg2/dom-builder/dom/DomLite");
var XlsxDomErrorsImpl_1 = require("../errors/XlsxDomErrorsImpl");
var XlsxNamespace = require("./XlsxNamespace");
/** Implementation of {@link OpenXmlIo.ReaderContext} and {@link OpenXmlIo.WriterContext}, contains:
 * - An XMLDocument containing the file data
 * - A DomBuilderFactory for creating new DOM elements when writing data back to the file
 * - An XlsxDom utility object with methods to make reading/writing data to DOM elements easier
 * - An XlsxDomValidator for checking that DOM elements match expected formats and throwing errors when not
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
var XmlFileInst = /** @class */ (function (_super) {
    __extends(XmlFileInst, _super);
    /**
     * Create a XML file instance backed by a {@link XMLDocument}
     * @param dom a {@link DocumentLike} object that this XML file will read from and write to
     * @param namespaceURI optional namespace to assign to elements created by this instance's
     * `domBldr` property. If provided, elements will be created using `dom.createElementNS()`.
     */
    function XmlFileInst(dom, namespaceURI, attributeNamespaceHandler) {
        var _this = _super.call(this, dom, XlsxDomErrorsImpl_1.XlsxDomErrorsImpl) || this;
        _this.dom = dom;
        // custom handling for attribute namespaces in OpenXML files
        _this.domBldr = new DomBuilderFactory_1.DomBuilderFactory(dom, namespaceURI, attributeNamespaceHandler);
        _this.validator = XlsxDomErrorsImpl_1.XlsxDomErrorsImpl;
        _this.readMulti = function (reader, elems, expectedElemName) { return XmlFileInst.readMulti(_this, reader, elems, expectedElemName); };
        _this.writeMulti = function (writer, insts, keysOrExpectedElemName) { return XmlFileInst.writeMulti(_this, writer, insts, keysOrExpectedElemName); };
        return _this;
    }
    XmlFileInst.newInst = function (dom, namespaceURI) {
        return new XmlFileInst(dom, namespaceURI, function (elem, name) { return XmlFileInst.lookupAndAddNamespace(dom, elem, name); });
    };
    /** Logic for reading/writing an array of OpenXml elements using a reader/writer for a single element of the same type
     */
    XmlFileInst.readMulti = function (xmlDoc, reader, elems, expectedElemName) {
        var res = [];
        for (var i = 0, size = elems.length; i < size; i++) {
            var elem = elems[i];
            res.push(reader(xmlDoc, elem, expectedElemName));
        }
        return res;
    };
    XmlFileInst.writeMulti = function (xmlDoc, writer, insts, keysOrExpectedElemName) {
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
    };
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
    XmlFileInst.lookupAndAddNamespace = function (document, element, qualifiedName) {
        var _a;
        var colonIdx = qualifiedName.indexOf(':');
        var namespaceUri = null;
        if (qualifiedName.startsWith('xml:')) {
            namespaceUri = DomLite_1.DomLite.XML_NAMESPACE;
        }
        if (namespaceUri == null) {
            namespaceUri = document.lookupNamespaceURI(qualifiedName.substring(0, colonIdx));
        }
        var documentElement = document.documentElement;
        var prefix = qualifiedName.substring(0, colonIdx);
        if (namespaceUri == null && prefix != null) {
            namespaceUri = XlsxNamespace.openxmlNamespaces[prefix];
            // If an OpenXML 'additional' namespace is used for an attribute, add it to the root of the document
            if (namespaceUri != null) {
                documentElement.setAttributeNS(DomLite_1.DomLite.XMLNS_NAMESPACE, "xmlns:".concat(prefix), namespaceUri);
            }
        }
        if (namespaceUri == null && prefix != null) {
            namespaceUri = XlsxNamespace.xlsxAdditionalNamespaces[prefix];
            // If an OpenXML 'additional' namespace is used for an attribute, add it and ignore it on the
            // root of the document since this is how OpenXML files handle namespaces
            if (namespaceUri != null) {
                documentElement.setAttributeNS(DomLite_1.DomLite.XMLNS_NAMESPACE, "xmlns:".concat(prefix), namespaceUri);
                // set the 'mc' namespace
                var ignorableNsUri = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
                documentElement.setAttributeNS(DomLite_1.DomLite.XMLNS_NAMESPACE, 'xmlns:mc', ignorableNsUri);
                // add the new ignorable prefi to the existing 'mc:Ignorable' prefixes list (if present)
                var rootAttrs = documentElement.attributes;
                var prefixesAttr = rootAttrs.getNamedItemNS(ignorableNsUri, 'Ignorable');
                var prefixes = "".concat(prefix).concat((prefixesAttr === null || prefixesAttr === void 0 ? void 0 : prefixesAttr.value) ? " ".concat(prefixesAttr.value) : '');
                documentElement.setAttributeNS(ignorableNsUri, 'mc:Ignorable', prefixes);
            }
        }
        if (namespaceUri == null) {
            namespaceUri = (_a = element.namespaceURI) !== null && _a !== void 0 ? _a : null;
        }
        return namespaceUri;
    };
    return XmlFileInst;
}(DomBuilderHelper_1.DomBuilderHelper));
exports.XmlFileInst = XmlFileInst;
