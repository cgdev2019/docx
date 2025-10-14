package com.example.docx.parser;

import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilderFactory;

import static org.junit.jupiter.api.Assertions.*;

class XmlUtilsTest {

    @Test
    void booleanElementTreatsMissingValAsTrue() throws Exception {
        Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        Element parent = doc.createElementNS(Namespaces.WORD_MAIN, "w:rPr");
        doc.appendChild(parent);

        Element bold = doc.createElementNS(Namespaces.WORD_MAIN, "w:b");
        parent.appendChild(bold);

        assertTrue(XmlUtils.booleanElement(parent, Namespaces.WORD_MAIN, "b"));

        bold.setAttributeNS(Namespaces.WORD_MAIN, "w:val", "");
        assertTrue(XmlUtils.booleanElement(parent, Namespaces.WORD_MAIN, "b"));

        bold.setAttributeNS(Namespaces.WORD_MAIN, "w:val", "0");
        assertFalse(XmlUtils.booleanElement(parent, Namespaces.WORD_MAIN, "b"));

        bold.setAttributeNS(Namespaces.WORD_MAIN, "w:val", "false");
        assertFalse(XmlUtils.booleanElement(parent, Namespaces.WORD_MAIN, "b"));

        bold.setAttributeNS(Namespaces.WORD_MAIN, "w:val", "1");
        assertTrue(XmlUtils.booleanElement(parent, Namespaces.WORD_MAIN, "b"));

        bold.setAttributeNS(Namespaces.WORD_MAIN, "w:val", "true");
        assertTrue(XmlUtils.booleanElement(parent, Namespaces.WORD_MAIN, "b"));
    }
}
