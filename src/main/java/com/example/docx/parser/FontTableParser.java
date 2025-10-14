package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.support.FontTable;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Parses {@code word/fontTable.xml}.
 */
public final class FontTableParser {

    public FontTable parse(DocxArchive archive) throws IOException {
        if (!archive.exists("word/fontTable.xml")) {
            return null;
        }
        try (InputStream input = archive.open("word/fontTable.xml")) {
            Document document = XmlUtils.parse(input);
            Element root = document.getDocumentElement();
            Map<String, FontTable.Font> fonts = new LinkedHashMap<>();
            for (Element fontEl : XmlUtils.children(root, Namespaces.WORD_MAIN, "font")) {
                String name = fontEl.getAttributeNS(Namespaces.WORD_MAIN, "name");
                Map<String, String> props = new LinkedHashMap<>();
                for (Element child : XmlUtils.childElements(fontEl)) {
                    if (!Namespaces.WORD_MAIN.equals(child.getNamespaceURI())) {
                        continue;
                    }
                    String val = child.getAttributeNS(Namespaces.WORD_MAIN, "val");
                    if (!val.isEmpty()) {
                        props.put(child.getLocalName(), val);
                    }
                }
                fonts.put(name, new FontTable.Font(name, props, fontEl));
            }
            return new FontTable(fonts);
        }
    }
}
