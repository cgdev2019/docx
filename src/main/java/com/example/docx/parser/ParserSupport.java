package com.example.docx.parser;

import com.example.docx.DocxException;
import org.w3c.dom.Element;

final class ParserSupport {

    private ParserSupport() {
    }

    static DocxException unknownElement(String context, Element element) {
        String namespace = element.getNamespaceURI();
        String localName = element.getLocalName();
        String prefix = element.getPrefix();
        String qualified = prefix != null ? prefix + ":" + localName : localName;
        return new DocxException("Unsupported " + context + " element <" + qualified + "> in namespace " + namespace);
    }
}
