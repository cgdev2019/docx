package com.example.docx.parser;

import com.example.docx.model.document.WordDocument;
import org.w3c.dom.Element;

/**
 * Helpers for structured document tags.
 */
final class SdtParser {

    private SdtParser() {
    }

    static WordDocument.SdtProperties parseProperties(Element sdtElement) {
        Element sdtPr = XmlUtils.firstChild(sdtElement, Namespaces.WORD_MAIN, "sdtPr").orElse(null);
        if (sdtPr == null) {
            return new WordDocument.SdtProperties(null, null, null, null);
        }
        String tag = XmlUtils.firstChild(sdtPr, Namespaces.WORD_MAIN, "tag")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .orElse(null);
        String alias = XmlUtils.firstChild(sdtPr, Namespaces.WORD_MAIN, "alias")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .orElse(null);
        String id = XmlUtils.firstChild(sdtPr, Namespaces.WORD_MAIN, "id")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .orElse(null);
        return new WordDocument.SdtProperties(tag, alias, id, sdtPr);
    }
}
