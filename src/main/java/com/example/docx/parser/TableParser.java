package com.example.docx.parser;

import com.example.docx.model.document.WordDocument;
import com.example.docx.model.relationship.RelationshipSet;
import org.w3c.dom.Element;

import java.util.ArrayList;
import java.util.List;

/**
 * Parses {@code w:tbl} structures.
 */
final class TableParser {

    private final ParsingContext context;

    TableParser(ParsingContext context) {
        this.context = context;
    }

    WordDocument.Table parseTable(Element tableElement, RelationshipSet relationships) {
        WordDocument.TableProperties properties = parseTableProperties(tableElement);
        List<WordDocument.TableRow> rows = new ArrayList<>();
        for (Element child : XmlUtils.childElements(tableElement)) {
            if (!Namespaces.WORD_MAIN.equals(child.getNamespaceURI())) {
                throw ParserSupport.unknownElement("table", child);
            }
            switch (child.getLocalName()) {
                case "tblPr", "tblGrid", "tblPrEx" -> {
                    // handled elsewhere or not required for in-memory model yet
                }
                case "tr" -> rows.add(parseTableRow(child, relationships));
                default -> throw ParserSupport.unknownElement("table", child);
            }
        }
        return new WordDocument.Table(properties, rows);
    }

    private WordDocument.TableRow parseTableRow(Element rowElement, RelationshipSet relationships) {
        WordDocument.TableRowProperties properties = parseTableRowProperties(rowElement);
        List<WordDocument.TableCell> cells = new ArrayList<>();
        for (Element child : XmlUtils.childElements(rowElement)) {
            if (!Namespaces.WORD_MAIN.equals(child.getNamespaceURI())) {
                throw ParserSupport.unknownElement("table row", child);
            }
            switch (child.getLocalName()) {
                case "trPr" -> {
                    // already extracted in properties
                }
                case "tc" -> cells.add(parseTableCell(child, relationships));
                default -> throw ParserSupport.unknownElement("table row", child);
            }
        }
        return new WordDocument.TableRow(properties, cells);
    }

    private WordDocument.TableCell parseTableCell(Element cellElement, RelationshipSet relationships) {
        WordDocument.TableCellProperties properties = parseTableCellProperties(cellElement);
        List<WordDocument.Block> blocks = new ArrayList<>();
        for (Element child : XmlUtils.childElements(cellElement)) {
            if (!Namespaces.WORD_MAIN.equals(child.getNamespaceURI())) {
                throw ParserSupport.unknownElement("table cell", child);
            }
            switch (child.getLocalName()) {
                case "tcPr" -> {
                    // already consumed in properties
                }
                case "p" -> blocks.add(context.paragraphParser.parseParagraph(child, relationships));
                case "tbl" -> blocks.add(parseTable(child, relationships));
                case "sdt" -> blocks.add(context.blockParser.parseStructuredDocumentTag(child, relationships));
                default -> throw ParserSupport.unknownElement("table cell", child);
            }
        }
        return new WordDocument.TableCell(properties, blocks);
    }

    private WordDocument.TableProperties parseTableProperties(Element element) {
        Element tblPr = XmlUtils.firstChild(element, Namespaces.WORD_MAIN, "tblPr").orElse(null);
        if (tblPr == null) {
            return new WordDocument.TableProperties(null, null, null, null, null);
        }
        String styleId = XmlUtils.firstChild(tblPr, Namespaces.WORD_MAIN, "tblStyle")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .filter(val -> !val.isEmpty())
                .orElse(null);
        Element tblW = XmlUtils.firstChild(tblPr, Namespaces.WORD_MAIN, "tblW").orElse(null);
        Integer width = tblW != null ? XmlUtils.intAttribute(tblW, "w:w") : null;
        String widthType = tblW != null ? tblW.getAttributeNS(Namespaces.WORD_MAIN, "type") : null;
        Element tblLook = XmlUtils.firstChild(tblPr, Namespaces.WORD_MAIN, "tblLook").orElse(null);
        Integer look = null;
        if (tblLook != null) {
            String val = tblLook.getAttributeNS(Namespaces.WORD_MAIN, "val");
            if (!val.isEmpty()) {
                try {
                    look = Integer.parseInt(val, 16);
                } catch (NumberFormatException e) {
                    look = null;
                }
            }
        }
        return new WordDocument.TableProperties(styleId, width, widthType, look, tblPr);
    }

    private WordDocument.TableRowProperties parseTableRowProperties(Element element) {
        Element trPr = XmlUtils.firstChild(element, Namespaces.WORD_MAIN, "trPr").orElse(null);
        if (trPr == null) {
            return new WordDocument.TableRowProperties(false, null, null, null, null);
        }
        boolean cantSplit = XmlUtils.booleanElement(trPr, Namespaces.WORD_MAIN, "cantSplit");
        Integer gridAfter = XmlUtils.firstChild(trPr, Namespaces.WORD_MAIN, "gridAfter")
                .map(el -> XmlUtils.intAttribute(el, "w:val"))
                .orElse(null);
        Integer gridBefore = XmlUtils.firstChild(trPr, Namespaces.WORD_MAIN, "gridBefore")
                .map(el -> XmlUtils.intAttribute(el, "w:val"))
                .orElse(null);
        Integer heightTwip = XmlUtils.firstChild(trPr, Namespaces.WORD_MAIN, "trHeight")
                .map(el -> XmlUtils.intAttribute(el, "w:val"))
                .orElse(null);
        return new WordDocument.TableRowProperties(cantSplit, gridAfter, gridBefore, heightTwip, trPr);
    }

    private WordDocument.TableCellProperties parseTableCellProperties(Element element) {
        Element tcPr = XmlUtils.firstChild(element, Namespaces.WORD_MAIN, "tcPr").orElse(null);
        if (tcPr == null) {
            return new WordDocument.TableCellProperties(null, null, null, null, false, null);
        }
        Integer gridSpan = XmlUtils.firstChild(tcPr, Namespaces.WORD_MAIN, "gridSpan")
                .map(el -> XmlUtils.intAttribute(el, "w:val"))
                .orElse(null);
        Element widthEl = XmlUtils.firstChild(tcPr, Namespaces.WORD_MAIN, "tcW").orElse(null);
        Integer width = widthEl != null ? XmlUtils.intAttribute(widthEl, "w:w") : null;
        String widthType = widthEl != null ? widthEl.getAttributeNS(Namespaces.WORD_MAIN, "type") : null;
        String vAlign = XmlUtils.firstChild(tcPr, Namespaces.WORD_MAIN, "vAlign")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .orElse(null);
        boolean vMerge = XmlUtils.firstChild(tcPr, Namespaces.WORD_MAIN, "vMerge").isPresent();
        return new WordDocument.TableCellProperties(gridSpan, width, widthType, vAlign, vMerge, tcPr);
    }
}
