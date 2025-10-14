package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.document.WordDocument;
import com.example.docx.model.relationship.RelationshipSet;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.IOException;
import java.io.InputStream;

/**
 * Parses the main document part ({@code word/document.xml}).
 */
public final class MainDocumentParser {

    private final ParsingContext context;

    public MainDocumentParser(ParsingContext context) {
        this.context = context;
    }

    public WordDocument parse(DocxArchive archive, RelationshipSet relationships) throws IOException {
        if (!archive.exists("word/document.xml")) {
            return null;
        }
        try (InputStream input = archive.open("word/document.xml")) {
            Document document = XmlUtils.parse(input);
            Element root = document.getDocumentElement();
            Element body = XmlUtils.firstChild(root, Namespaces.WORD_MAIN, "body")
                    .orElseThrow(() -> new IOException("Missing <w:body> element"));
            WordDocument.Builder builder = WordDocument.builder();
            for (Element child : XmlUtils.childElements(body)) {
                if (!Namespaces.WORD_MAIN.equals(child.getNamespaceURI())) {
                    throw ParserSupport.unknownElement("body", child);
                }
                if ("sectPr".equals(child.getLocalName())) {
                    builder.sectionProperties(SectionParser.parseSectionProperties(child));
                } else {
                    builder.addBlock(context.blockParser.parse(child, relationships));
                }
            }
            return builder.build();
        }
    }
}
