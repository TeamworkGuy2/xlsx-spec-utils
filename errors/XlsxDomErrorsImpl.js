"use strict";
/** Common Excel file XML DOM error checking/throwing functions
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
var XlsxDomErrorsImpl;
(function (XlsxDomErrorsImpl) {
    var typeCheck = XlsxDomErrorsImpl; // TODO type-checker
    function missingNode(nodeName) {
        return new Error("Error reading Excel template, missing required '" + nodeName + "' node");
    }
    XlsxDomErrorsImpl.missingNode = missingNode;
    function expectNode(node, expectedNodeName, parentNodeName, idx, size) {
        if (node.tagName !== expectedNodeName) {
            throw unexpectedNode(node.tagName, expectedNodeName, parentNodeName, idx, size);
        }
    }
    XlsxDomErrorsImpl.expectNode = expectNode;
    function unexpectedNode(badNodeName, expectedNodeName, parentNodeName, idx, size) {
        return new Error("Error reading Excel template, unexpected '" + badNodeName + "' node" +
            (expectedNodeName ? ", expected only '" + expectedNodeName + "' nodes" : "") +
            (parentNodeName ? ", of parent node '" + parentNodeName + "'" : "") +
            (idx || size ? (idx ? ", index=" + idx : "") + (size ? ", size=" + size : "") : ""));
    }
    XlsxDomErrorsImpl.unexpectedNode = unexpectedNode;
})(XlsxDomErrorsImpl || (XlsxDomErrorsImpl = {}));
module.exports = XlsxDomErrorsImpl;
