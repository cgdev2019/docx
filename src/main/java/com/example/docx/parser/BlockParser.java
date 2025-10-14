package com.example.docx.parser;

import com.example.docx.model.document.WordDocument;
import com.example.docx.model.relationship.RelationshipSet;
import org.w3c.dom.Element;

import java.util.ArrayList;
import java.util.List;

/**
 * Utility for parsing block-level elements.
 */
final class BlockParser {

    private final ParsingContext context;

    BlockParser(ParsingContext context) {
        this.context = context;
    }

    WordDocument.Block parse(Element element, RelationshipSet relationships) {
        if (!Namespaces.WORD_MAIN.equals(element.getNamespaceURI())) {
            throw ParserSupport.unknownElement("block", element);
        }
        return switch (element.getLocalName()) {
            case "p" -> context.paragraphParser.parseParagraph(element, relationships);
            case "tbl" -> context.tableParser.parseTable(element, relationships);
            case "sdt" -> parseStructuredDocumentTag(element, relationships);
            case "sectPr" -> new WordDocument.SectionBreak(SectionParser.parseSectionProperties(element));
            case "bookmarkStart" -> parseBookmark(element, WordDocument.Bookmark.Kind.START);
            case "bookmarkEnd" -> parseBookmark(element, WordDocument.Bookmark.Kind.END);
            default -> throw ParserSupport.unknownElement("block", element);
        };
    }

    WordDocument.StructuredDocumentTag parseStructuredDocumentTag(Element element, RelationshipSet relationships) {
        WordDocument.SdtProperties props = SdtParser.parseProperties(element);
        Element content = XmlUtils.firstChild(element, Namespaces.WORD_MAIN, "sdtContent").orElse(null);
        List<WordDocument.Block> blocks = content != null
                ? parseChildBlocks(content, relationships)
                : List.of();
        return new WordDocument.StructuredDocumentTag(props, blocks);
    }

    List<WordDocument.Block> parseChildBlocks(Element parent, RelationshipSet relationships) {
        List<WordDocument.Block> blocks = new ArrayList<>();
        for (Element child : XmlUtils.childElements(parent)) {
            blocks.add(parse(child, relationships));
        }
        return blocks;
    }

    private WordDocument.Bookmark parseBookmark(Element element, WordDocument.Bookmark.Kind kind) {
        String id = element.getAttributeNS(Namespaces.WORD_MAIN, "id");
        String name = element.hasAttributeNS(Namespaces.WORD_MAIN, "name")
                ? element.getAttributeNS(Namespaces.WORD_MAIN, "name")
                : null;
        return new WordDocument.Bookmark(kind, id, name);
    }
}
