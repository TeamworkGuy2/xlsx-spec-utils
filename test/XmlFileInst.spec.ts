import * as chai from "chai";
import * as JSDom from "jsdom";
import { DomBuilderFactory } from "@twg2/dom-builder/dom/DomBuilderFactory";
import { XmlFileInst } from "../files/XmlFileInst";

var asr = chai.assert;

suite("DomBuilder", function domBuilder() {

    test("create w/ lookupAndAddNamespace", function createWithNamespacesTest() {
        const dom = new JSDom.JSDOM(
            "<?xml version=\"1.0\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></sst>",
            { contentType: "text/xml" }
        ).window.document;
        const domBldr = new DomBuilderFactory(dom, null, (elem, name) => XmlFileInst.lookupAndAddNamespace(dom, elem, name));
        // element with multiple attributes with namespaces to test the DomBuilderFactory's lookupAndAddNamespace()
        const elem = domBldr.create('s')
            .attrBool('xml:space', true, true, '1', '0')
            .attr('x14ac:dyDescent', '0.25')
            .attr('xr:test', '123')
            .attr('r:id', 'test')
            .element;
        dom.documentElement.appendChild(elem);

        asr.equal(dom.documentElement.outerHTML.replace(/ xmlns=""/g, ""),
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr x14ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                '<s xml:space="1" x14ac:dyDescent="0.25" xr:test="123" r:id="test"/>' +
            '</sst>'
        );
    });
});