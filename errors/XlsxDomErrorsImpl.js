"use strict";
/** Common XLSX XML file DOM error checking/throwing functions
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
var XlsxDomErrorsImpl;
(function (XlsxDomErrorsImpl) {
    var typeCheck = XlsxDomErrorsImpl; // TODO type-checker
    function missingNode(nodeName, parent) {
        return new Error("Error reading XLSX, missing required node '" + nodeName + "'");
    }
    XlsxDomErrorsImpl.missingNode = missingNode;
    function missingAttribute(attributeName, parent) {
        return new Error("Error reading XLSX, missing required attribute '" + attributeName + "'");
    }
    XlsxDomErrorsImpl.missingAttribute = missingAttribute;
    function expectNode(node, expectedNodeName, parentNodeName, idx, size) {
        if (node.tagName !== expectedNodeName) {
            throw unexpectedNode(node.tagName, expectedNodeName, parentNodeName, idx, size);
        }
    }
    XlsxDomErrorsImpl.expectNode = expectNode;
    function unexpectedNode(badNodeName, expectedNodeName, parentNodeName, idx, size) {
        return new Error("Error reading XLSXtemplate, unexpected node '" + badNodeName + "'" +
            (expectedNodeName ? ", expected only '" + expectedNodeName + "' nodes" : "") +
            (parentNodeName ? ", of parent node '" + parentNodeName + "'" : "") +
            (idx || size ? (idx ? ", index=" + idx : "") + (size ? ", size=" + size : "") : ""));
    }
    XlsxDomErrorsImpl.unexpectedNode = unexpectedNode;
})(XlsxDomErrorsImpl || (XlsxDomErrorsImpl = {}));
module.exports = XlsxDomErrorsImpl;
