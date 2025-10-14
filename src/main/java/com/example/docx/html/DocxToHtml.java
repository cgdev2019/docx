package com.example.docx.html;

import com.example.docx.model.DocxPackage;
import com.example.docx.model.document.WordDocument;
import com.example.docx.model.relationship.RelationshipSet;
import com.example.docx.model.styles.StyleDefinitions;
import com.example.docx.model.support.Theme;
import com.example.docx.parser.Namespaces;
import com.example.docx.parser.XmlUtils;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;

public final class DocxToHtml {

    static final String DOCUMENT_PART = "word/document.xml";
    static final WordDocument.RunProperties EMPTY_RUN_PROPERTIES =
            new WordDocument.RunProperties(null, false, false, false, null,
                    false, false, false, false, false, null, null, null,
                    null, null, Map.of(), null);
    static final WordDocument.ParagraphProperties EMPTY_PARAGRAPH_PROPERTIES =
            new WordDocument.ParagraphProperties(null, null, null, null, null,
                    null, false, false, false, List.of(), null);

    private final String language;

    public DocxToHtml() {
        this("fr");
    }

    public DocxToHtml(String language) {
        this.language = language == null || language.isBlank() ? "fr" : language;
    }

    public String convert(DocxPackage docxPackage) {
        Objects.requireNonNull(docxPackage, "docxPackage");
        WordDocument document = docxPackage.document().orElse(null);
        StyleDefinitions styles = docxPackage.styles().orElse(StyleDefinitions.empty());
        Map<String, String> themeColors = extractThemeColors(docxPackage);
        StyleResolver styleResolver = new StyleResolver(styles, themeColors);
        RelationshipSet relationships = docxPackage.relationshipsByPart().get(DOCUMENT_PART);
        HyperlinkResolver hyperlinkResolver = new HyperlinkResolver(relationships);
        StyleResolver.ResolvedParagraph baseParagraph = styleResolver.resolveParagraph(null, List.of());
        StyleResolver.ResolvedRun baseRun = styleResolver.resolveRun(EMPTY_RUN_PROPERTIES, baseParagraph);
        StyleRegistry registry = new StyleRegistry(ParagraphCss.from(baseParagraph, baseRun), RunCss.from(baseRun));

        RenderContext renderContext = new RenderContext(styleResolver, hyperlinkResolver, registry, themeColors);
        BlockRenderer blockRenderer = new BlockRenderer(renderContext);

        String bodyContent;
        if (document == null || document.bodyElements().isEmpty()) {
            bodyContent = "<p class=\"docx-paragraph docx-empty\">Document vide</p>";
        } else {
            bodyContent = blockRenderer.renderBlocks(document.bodyElements());
        }

        return buildHtml(document, registry, bodyContent);
    }

    private static Map<String, String> extractThemeColors(DocxPackage docxPackage) {
        if (docxPackage == null) {
            return Map.of();
        }
        return docxPackage.theme()
                .flatMap(Theme::document)
                .map(DocxToHtml::readThemeColorScheme)
                .orElse(Map.of());
    }

    private static Map<String, String> readThemeColorScheme(Document document) {
        if (document == null) {
            return Map.of();
        }
        Element root = document.getDocumentElement();
        if (root == null) {
            return Map.of();
        }
        Element themeElements = XmlUtils.firstChild(root, Namespaces.DRAWINGML_MAIN, "themeElements").orElse(null);
        if (themeElements == null) {
            return Map.of();
        }
        Element clrScheme = XmlUtils.firstChild(themeElements, Namespaces.DRAWINGML_MAIN, "clrScheme").orElse(null);
        if (clrScheme == null) {
            return Map.of();
        }
        Map<String, String> colors = new LinkedHashMap<>();
        for (Element entry : XmlUtils.childElements(clrScheme)) {
            if (entry == null || entry.getLocalName() == null) {
                continue;
            }
            if (!Namespaces.DRAWINGML_MAIN.equals(entry.getNamespaceURI())) {
                continue;
            }
            String key = entry.getLocalName().toLowerCase(Locale.ROOT);
            String value = readSchemeColor(entry);
            if (value != null) {
                colors.put(key, value);
            }
        }
        return Map.copyOf(colors);
    }

    private static String readSchemeColor(Element entry) {
        for (Element child : XmlUtils.childElements(entry)) {
            if (child == null || !Namespaces.DRAWINGML_MAIN.equals(child.getNamespaceURI())) {
                continue;
            }
            String localName = child.getLocalName();
            if ("srgbClr".equals(localName)) {
                String val = child.getAttribute("val");
                if (val != null && !val.isBlank()) {
                    return '#' + val.toLowerCase(Locale.ROOT);
                }
            }
        }
        return null;
    }

    private String buildHtml(WordDocument document, StyleRegistry registry, String bodyContent) {
        StringBuilder builder = new StringBuilder();
        builder.append("<!DOCTYPE html>\n");
        builder.append("<html lang=\"").append(DocxHtmlUtils.escapeHtmlAttribute(language)).append("\">\n");
        builder.append("<head>\n<meta charset=\"utf-8\">\n<style>\n");
        builder.append(registry.buildCss(document));
        builder.append("</style>\n</head>\n<body class=\"docx-body\">\n");
        builder.append(bodyContent);
        builder.append("\n</body>\n</html>");
        return builder.toString();
    }
}
