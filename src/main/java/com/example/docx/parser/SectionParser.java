package com.example.docx.parser;

import com.example.docx.model.document.WordDocument;
import org.w3c.dom.Element;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Parses {@code w:sectPr} elements.
 */
final class SectionParser {

    private SectionParser() {
    }

    static WordDocument.SectionProperties parseSectionProperties(Element element) {
        Element pgSz = XmlUtils.firstChild(element, Namespaces.WORD_MAIN, "pgSz").orElse(null);
        WordDocument.PageDimensions dimensions = null;
        if (pgSz != null) {
            int w = XmlUtils.intAttribute(pgSz, "w:w") != null ? XmlUtils.intAttribute(pgSz, "w:w") : 0;
            int h = XmlUtils.intAttribute(pgSz, "w:h") != null ? XmlUtils.intAttribute(pgSz, "w:h") : 0;
            String orientation = pgSz.getAttributeNS(Namespaces.WORD_MAIN, "orient");
            orientation = orientation.isEmpty() ? null : orientation;
            dimensions = new WordDocument.PageDimensions(w, h, orientation);
        }
        Element pgMar = XmlUtils.firstChild(element, Namespaces.WORD_MAIN, "pgMar").orElse(null);
        WordDocument.PageMargins margins = null;
        if (pgMar != null) {
            margins = new WordDocument.PageMargins(
                    valueOrZero(pgMar, "top"),
                    valueOrZero(pgMar, "right"),
                    valueOrZero(pgMar, "bottom"),
                    valueOrZero(pgMar, "left"),
                    valueOrZero(pgMar, "header"),
                    valueOrZero(pgMar, "footer"),
                    valueOrZero(pgMar, "gutter"));
        }
        Map<String, String> headerRefs = new LinkedHashMap<>();
        Map<String, String> footerRefs = new LinkedHashMap<>();
        for (Element ref : XmlUtils.childElements(element)) {
            if (!Namespaces.WORD_MAIN.equals(ref.getNamespaceURI())) {
                continue;
            }
            switch (ref.getLocalName()) {
                case "headerReference" -> headerRefs.put(ref.getAttributeNS(Namespaces.WORD_MAIN, "type"),
                        ref.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"));
                case "footerReference" -> footerRefs.put(ref.getAttributeNS(Namespaces.WORD_MAIN, "type"),
                        ref.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"));
                default -> {
                }
            }
        }
        WordDocument.HeaderFooterReference headerFooter = headerRefs.isEmpty() && footerRefs.isEmpty()
                ? null
                : new WordDocument.HeaderFooterReference(headerRefs, footerRefs);
        String sectionType = XmlUtils.firstChild(element, Namespaces.WORD_MAIN, "type")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .orElse(null);
        return new WordDocument.SectionProperties(dimensions, margins, headerFooter, sectionType, element);
    }

    private static int valueOrZero(Element parent, String attr) {
        String value = parent.getAttributeNS(Namespaces.WORD_MAIN, attr);
        if (value == null || value.isEmpty()) {
            return 0;
        }
        try {
            return Integer.parseInt(value);
        } catch (NumberFormatException e) {
            return 0;
        }
    }
}
