package com.example.docx.parser;

import com.example.docx.model.document.WordDocument;
import com.example.docx.model.relationship.RelationshipSet;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.util.ArrayList;
import java.util.List;

/**
 * Handles parsing of paragraph elements.
 */
final class ParagraphParser {

    private final ParsingContext context;

    ParagraphParser(ParsingContext context) {
        this.context = context;
    }

    WordDocument.Paragraph parseParagraph(Element paragraph, RelationshipSet relationships) {
        WordDocument.ParagraphProperties properties = parseProperties(paragraph);
        List<WordDocument.ParagraphContent> contents = new ArrayList<>();
        for (Element child : XmlUtils.childElements(paragraph)) {
            if (!Namespaces.WORD_MAIN.equals(child.getNamespaceURI())) {
                throw ParserSupport.unknownElement("paragraph content", child);
            }
            switch (child.getLocalName()) {
                case "pPr" -> {
                    // already handled when parsing properties
                }
                case "r" -> contents.add(context.runParser.parse(child));
                case "hyperlink" -> contents.add(parseHyperlink(child, relationships));
                case "bookmarkStart" -> contents.add(new WordDocument.BookmarkStart(
                        child.getAttributeNS(Namespaces.WORD_MAIN, "id"),
                        child.getAttributeNS(Namespaces.WORD_MAIN, "name")));
                case "bookmarkEnd" -> contents.add(new WordDocument.BookmarkEnd(
                        child.getAttributeNS(Namespaces.WORD_MAIN, "id")));
                case "fldSimple" -> contents.add(parseSimpleField(child));
                case "sdt" -> contents.add(parseStructuredDocumentTagRun(child, relationships));
                default -> throw ParserSupport.unknownElement("paragraph content", child);
            }
        }
        return new WordDocument.Paragraph(properties, contents);
    }

    private WordDocument.StructuredDocumentTagRun parseStructuredDocumentTagRun(Element element, RelationshipSet relationships) {
        WordDocument.SdtProperties props = SdtParser.parseProperties(element);
        Element content = XmlUtils.firstChild(element, Namespaces.WORD_MAIN, "sdtContent").orElse(null);
        List<WordDocument.ParagraphContent> items = new ArrayList<>();
        if (content != null) {
            for (Element child : XmlUtils.childElements(content)) {
                items.add(parseParagraphContent(child, relationships));
            }
        }
        return new WordDocument.StructuredDocumentTagRun(props, items);
    }

    private WordDocument.ParagraphContent parseParagraphContent(Element element, RelationshipSet relationships) {
        if (!Namespaces.WORD_MAIN.equals(element.getNamespaceURI())) {
            throw ParserSupport.unknownElement("structured document content", element);
        }
        return switch (element.getLocalName()) {
            case "r" -> context.runParser.parse(element);
            case "hyperlink" -> parseHyperlink(element, relationships);
            case "bookmarkStart" -> new WordDocument.BookmarkStart(
                    element.getAttributeNS(Namespaces.WORD_MAIN, "id"),
                    element.getAttributeNS(Namespaces.WORD_MAIN, "name"));
            case "bookmarkEnd" -> new WordDocument.BookmarkEnd(
                    element.getAttributeNS(Namespaces.WORD_MAIN, "id"));
            case "fldSimple" -> parseSimpleField(element);
            case "sdt" -> parseStructuredDocumentTagRun(element, relationships);
            default -> throw ParserSupport.unknownElement("structured document content", element);
        };
    }

    private WordDocument.Hyperlink parseHyperlink(Element element, RelationshipSet relationships) {
        List<WordDocument.Run> runs = new ArrayList<>();
        for (Element child : XmlUtils.children(element, Namespaces.WORD_MAIN, "r")) {
            runs.add(context.runParser.parse(child));
        }
        String relId = element.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
        relId = relId.isEmpty() ? null : relId;
        String anchor = element.getAttributeNS(Namespaces.WORD_MAIN, "anchor");
        anchor = anchor.isEmpty() ? null : anchor;
        return new WordDocument.Hyperlink(relId, anchor, runs);
    }

    private WordDocument.Field parseSimpleField(Element element) {
        String instruction = element.getAttributeNS(Namespaces.WORD_MAIN, "instr");
        WordDocument.Run instrRun = context.runParser.parse(createInstructionRun(element.getOwnerDocument(), instruction));
        List<WordDocument.Run> result = new ArrayList<>();
        for (Element child : XmlUtils.children(element, Namespaces.WORD_MAIN, "r")) {
            result.add(context.runParser.parse(child));
        }
        return new WordDocument.Field(List.of(instrRun), result);
    }

    private Element createInstructionRun(Document document, String instruction) {
        Element run = document.createElementNS(Namespaces.WORD_MAIN, "w:r");
        Element text = document.createElementNS(Namespaces.WORD_MAIN, "w:instrText");
        text.setTextContent(instruction);
        run.appendChild(text);
        return run;
    }

    private WordDocument.ParagraphProperties parseProperties(Element paragraph) {
        Element pPr = XmlUtils.firstChild(paragraph, Namespaces.WORD_MAIN, "pPr").orElse(null);
        if (pPr == null) {
            return new WordDocument.ParagraphProperties(null, null, null, null, null,
                    null, false, false, false, List.of(), null);
        }
        String styleId = XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "pStyle")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .filter(val -> !val.isEmpty())
                .orElse(null);
        WordDocument.NumberingReference numbering = parseNumbering(pPr);
        WordDocument.Alignment alignment = parseAlignment(pPr);
        WordDocument.Indentation indentation = parseIndentation(pPr);
        WordDocument.Spacing spacing = parseSpacing(pPr);
        Integer outlineLevel = XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "outlineLvl")
                .map(el -> XmlUtils.intAttribute(el, "w:val"))
                .orElse(null);
        boolean keepTogether = XmlUtils.booleanElement(pPr, Namespaces.WORD_MAIN, "keepLines");
        boolean keepWithNext = XmlUtils.booleanElement(pPr, Namespaces.WORD_MAIN, "keepNext");
        boolean pageBreakBefore = XmlUtils.booleanElement(pPr, Namespaces.WORD_MAIN, "pageBreakBefore");
        List<WordDocument.TabStop> tabs = parseTabs(pPr);
        return new WordDocument.ParagraphProperties(styleId, numbering, alignment, indentation, spacing,
                outlineLevel, keepTogether, keepWithNext, pageBreakBefore, tabs, pPr);
    }

    private WordDocument.NumberingReference parseNumbering(Element pPr) {
        Element numPr = XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "numPr").orElse(null);
        if (numPr == null) {
            return null;
        }
        Integer numId = XmlUtils.firstChild(numPr, Namespaces.WORD_MAIN, "numId")
                .map(el -> XmlUtils.intAttribute(el, "w:val"))
                .orElse(null);
        Integer level = XmlUtils.firstChild(numPr, Namespaces.WORD_MAIN, "ilvl")
                .map(el -> XmlUtils.intAttribute(el, "w:val"))
                .orElse(null);
        if (numId == null || level == null) {
            return null;
        }
        return new WordDocument.NumberingReference(numId, level);
    }

    private WordDocument.Alignment parseAlignment(Element pPr) {
        return XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "jc")
                .map(el -> switch (el.getAttributeNS(Namespaces.WORD_MAIN, "val")) {
                    case "left" -> WordDocument.Alignment.LEFT;
                    case "center" -> WordDocument.Alignment.CENTER;
                    case "right" -> WordDocument.Alignment.RIGHT;
                    case "both", "justify" -> WordDocument.Alignment.JUSTIFIED;
                    case "distribute" -> WordDocument.Alignment.DISTRIBUTE;
                    case "thaiDistribute" -> WordDocument.Alignment.THAI_DISTRIBUTED;
                    case "justLow" -> WordDocument.Alignment.JUSTIFY_LOW;
                    default -> null;
                })
                .orElse(null);
    }

    private WordDocument.Indentation parseIndentation(Element pPr) {
        Element ind = XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "ind").orElse(null);
        if (ind == null) {
            return null;
        }
        Integer left = XmlUtils.intAttribute(ind, "w:left");
        Integer right = XmlUtils.intAttribute(ind, "w:right");
        Integer firstLine = XmlUtils.intAttribute(ind, "w:firstLine");
        Integer hanging = XmlUtils.intAttribute(ind, "w:hanging");
        return new WordDocument.Indentation(left, right, firstLine, hanging);
    }

    private WordDocument.Spacing parseSpacing(Element pPr) {
        Element spacing = XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "spacing").orElse(null);
        if (spacing == null) {
            return null;
        }
        Integer before = XmlUtils.intAttribute(spacing, "w:before");
        Integer after = XmlUtils.intAttribute(spacing, "w:after");
        Integer line = XmlUtils.intAttribute(spacing, "w:line");
        String rule = spacing.getAttributeNS(Namespaces.WORD_MAIN, "lineRule");
        rule = rule.isEmpty() ? null : rule;
        return new WordDocument.Spacing(before, after, line, rule);
    }

    private List<WordDocument.TabStop> parseTabs(Element pPr) {
        Element tabsEl = XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "tabs").orElse(null);
        if (tabsEl == null) {
            return List.of();
        }
        List<WordDocument.TabStop> result = new ArrayList<>();
        for (Element tab : XmlUtils.children(tabsEl, Namespaces.WORD_MAIN, "tab")) {
            String val = tab.getAttributeNS(Namespaces.WORD_MAIN, "val");
            Integer pos = XmlUtils.intAttribute(tab, "w:pos");
            result.add(new WordDocument.TabStop(val, pos));
        }
        return result;
    }
}
