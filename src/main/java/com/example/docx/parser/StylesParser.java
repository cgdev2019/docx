package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.DocxPackage;
import com.example.docx.model.document.WordDocument;
import com.example.docx.model.relationship.RelationshipSet;
import com.example.docx.model.styles.StyleDefinitions;
import com.example.docx.model.styles.StyleDefinitions.Style;
import com.example.docx.model.styles.StyleDefinitions.TableStyleProperties;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Parses {@code word/styles.xml}.
 */
public final class StylesParser {

    private final RelationshipsParser relationshipsParser = new RelationshipsParser();
    private final ParsingContext context;

    public StylesParser(ParsingContext context) {
        this.context = context;
    }

    public StyleDefinitions parse(DocxArchive archive, DocxPackage.Builder builder) throws IOException {
        if (!archive.exists("word/styles.xml")) {
            return null;
        }
        RelationshipSet rels = relationshipsParser.parse(archive, "word/_rels/styles.xml.rels");
        if (rels != null) {
            builder.relationshipForPart("word/styles.xml", rels);
        }
        try (InputStream input = archive.open("word/styles.xml")) {
            Document document = XmlUtils.parse(input);
            Element root = document.getDocumentElement();
            Map<String, Style> styles = new LinkedHashMap<>();
            Element docDefaults = XmlUtils.firstChild(root, Namespaces.WORD_MAIN, "docDefaults").orElse(null);
            List<String> defaultHierarchy = new ArrayList<>();
            for (Element styleEl : XmlUtils.children(root, Namespaces.WORD_MAIN, "style")) {
                Style style = parseStyle(document, styleEl);
                styles.put(style.styleId(), style);
                if (style.defaultStyle()) {
                    defaultHierarchy.add(style.styleId());
                }
            }
            return new StyleDefinitions(styles, defaultHierarchy, docDefaults);
        }
    }

    private Style parseStyle(Document document, Element styleEl) {
        String styleId = styleEl.getAttributeNS(Namespaces.WORD_MAIN, "styleId");
        String type = styleEl.getAttributeNS(Namespaces.WORD_MAIN, "type");
        String name = XmlUtils.firstChild(styleEl, Namespaces.WORD_MAIN, "name")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .orElse(null);
        String basedOn = XmlUtils.firstChild(styleEl, Namespaces.WORD_MAIN, "basedOn")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .orElse(null);
        String next = XmlUtils.firstChild(styleEl, Namespaces.WORD_MAIN, "next")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .orElse(null);
        String link = XmlUtils.firstChild(styleEl, Namespaces.WORD_MAIN, "link")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .orElse(null);
        boolean defaultStyle = "1".equals(styleEl.getAttributeNS(Namespaces.WORD_MAIN, "default"));
        boolean customStyle = "1".equals(styleEl.getAttributeNS(Namespaces.WORD_MAIN, "customStyle"));

        WordDocument.ParagraphProperties paragraphProperties = null;
        WordDocument.RunProperties runProperties = null;
        TableStyleProperties tableProperties = null;

        Element pPr = XmlUtils.firstChild(styleEl, Namespaces.WORD_MAIN, "pPr").orElse(null);
        if (pPr != null) {
            Element paragraph = document.createElementNS(Namespaces.WORD_MAIN, "w:p");
            paragraph.appendChild(pPr.cloneNode(true));
            paragraphProperties = context.paragraphParser.parseParagraph(paragraph, null).properties();
        }
        Element rPr = XmlUtils.firstChild(styleEl, Namespaces.WORD_MAIN, "rPr").orElse(null);
        if (rPr != null) {
            Element run = document.createElementNS(Namespaces.WORD_MAIN, "w:r");
            run.appendChild(rPr.cloneNode(true));
            runProperties = context.runParser.parse(run).properties();
        }
        Element tblPr = XmlUtils.firstChild(styleEl, Namespaces.WORD_MAIN, "tblPr").orElse(null);
        if (tblPr != null) {
            tableProperties = new TableStyleProperties(tblPr);
        }
        return new Style(styleId, type, name, basedOn, next, link, defaultStyle, customStyle,
                paragraphProperties, runProperties, tableProperties, styleEl);
    }
}
