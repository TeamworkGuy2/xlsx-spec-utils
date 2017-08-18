/** Common XLSX XML file DOM error checking/throwing functions
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
module XlsxDomErrorsImpl {
    var typeCheck: DomValidate = XlsxDomErrorsImpl; // TODO type-checker


    export function missingNode(nodeName: string) {
        return new Error("Error reading XLSX template, missing required '" + nodeName + "' node");
    }


    export function expectNode(node: { tagName: string }, expectedNodeName: string, parentNodeName: string, idx?: number, size?: number) {
        if (node.tagName !== expectedNodeName) {
            throw unexpectedNode(node.tagName, expectedNodeName, parentNodeName, idx, size);
        }
    }


    export function unexpectedNode(badNodeName: string, expectedNodeName?: string, parentNodeName?: string, idx?: number, size?: number) {
        return new Error("Error reading XLSX template, unexpected '" + badNodeName + "' node" +
            (expectedNodeName ? ", expected only '" + expectedNodeName + "' nodes" : "") +
            (parentNodeName ? ", of parent node '" + parentNodeName + "'" : "") +
            (idx || size ? (idx ? ", index=" + idx : "") + (size ? ", size=" + size : "") : "")
        );
    }

}

export = XlsxDomErrorsImpl;