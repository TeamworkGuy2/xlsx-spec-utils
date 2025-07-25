/** Common XLSX XML file DOM error checking/throwing functions
 * @author TeamworkGuy2
 * @since 2016-5-27
 */
export module XlsxDomErrorsImpl {
    var typeCheck: DomValidate = XlsxDomErrorsImpl; // TODO type-checker


    export function missingNode(nodeName: string, parent?: any | null): Error {
        return new Error("Error reading XLSX, missing required node '" + nodeName + "'");
    }


    export function missingAttribute(attributeName: string, parent?: any | null): Error {
        return new Error("Error reading XLSX, missing required attribute '" + attributeName + "'");
    }


    export function expectNode(node: { tagName: string }, expectedNodeName: string, parentNodeName: string, idx?: number, size?: number) {
        if (node.tagName !== expectedNodeName) {
            throw unexpectedNode(node.tagName, expectedNodeName, parentNodeName, idx, size);
        }
    }


    export function unexpectedNode(badNodeName: string, expectedNodeName?: string, parentNodeName?: string, idx?: number, size?: number): Error {
        return new Error("Error reading XLSX template, unexpected node '" + badNodeName + "'" +
            (expectedNodeName ? ", expected only '" + expectedNodeName + "' nodes" : "") +
            (parentNodeName ? ", of parent node '" + parentNodeName + "'" : "") +
            (idx || size ? (idx ? ", index=" + idx : "") + (size ? ", size=" + size : "") : "")
        );
    }

}