package com.example.docx.parser;

import com.example.docx.model.document.WordDocument;
import org.w3c.dom.Element;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Parses run-level content ({@code w:r}).
 */
final class RunParser {

    private final ParsingContext context;

    RunParser(ParsingContext context) {
        this.context = context;
    }

    WordDocument.Run parse(Element runElement) {
        WordDocument.RunProperties properties = parseRunProperties(runElement);
        List<WordDocument.Inline> inlines = new ArrayList<>();
        for (Element child : XmlUtils.childElements(runElement)) {
            if (!Namespaces.WORD_MAIN.equals(child.getNamespaceURI())) {
                throw ParserSupport.unknownElement("run content", child);
            }
            switch (child.getLocalName()) {
                case "rPr" -> {
                    // already handled
                }
                case "t" -> inlines.add(new WordDocument.Text(child.getTextContent(),
                        "preserve".equals(child.getAttributeNS("http://www.w3.org/XML/1998/namespace", "space"))));
                case "tab" -> inlines.add(new WordDocument.Tab());
                case "br", "cr" -> inlines.add(parseBreak(child));
                case "noBreakHyphen" -> inlines.add(new WordDocument.NoBreakHyphen());
                case "softHyphen" -> inlines.add(new WordDocument.SoftHyphen());
                case "separator" -> inlines.add(new WordDocument.Separator(WordDocument.Separator.Kind.FOOTNOTE));
                case "continuationSeparator" -> inlines.add(new WordDocument.Separator(WordDocument.Separator.Kind.CONTINUATION));
                case "sym" -> inlines.add(parseSymbol(child));
                case "footnoteReference" -> inlines.add(new WordDocument.FootnoteReference(
                        Integer.parseInt(child.getAttributeNS(Namespaces.WORD_MAIN, "id"))));
                case "footnoteRef" -> inlines.add(new WordDocument.ReferenceMark(WordDocument.ReferenceMark.Kind.FOOTNOTE));
                case "endnoteReference" -> inlines.add(new WordDocument.EndnoteReference(
                        Integer.parseInt(child.getAttributeNS(Namespaces.WORD_MAIN, "id"))));
                case "endnoteRef" -> inlines.add(new WordDocument.ReferenceMark(WordDocument.ReferenceMark.Kind.ENDNOTE));
                case "commentReference" -> inlines.add(new WordDocument.CommentReference(
                        Integer.parseInt(child.getAttributeNS(Namespaces.WORD_MAIN, "id"))));
                case "fldChar" -> inlines.add(parseFieldChar(child));
                case "instrText" -> inlines.add(new WordDocument.FieldInstruction(child.getTextContent()));
                case "drawing" -> inlines.add(parseDrawing(child));
                case "pict" -> {
                    WordDocument.Inline inline = parsePict(child);
                    if (inline != null) {
                        inlines.add(inline);
                    }
                }
                default -> throw ParserSupport.unknownElement("run content", child);
            }
        }
        return new WordDocument.Run(properties, inlines);
    }

    private WordDocument.RunProperties parseRunProperties(Element runElement) {
        Element rPr = XmlUtils.firstChild(runElement, Namespaces.WORD_MAIN, "rPr").orElse(null);
        if (rPr == null) {
            return new WordDocument.RunProperties(null, false, false, false, null,
                    false, false, false, false, false, null, null, null,
                    null, null, Map.of(), null);
        }
        String styleId = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "rStyle")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .filter(val -> !val.isEmpty())
                .orElse(null);
        boolean bold = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "b");
        boolean italic = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "i");
        boolean underline = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "u").isPresent();
        String underlineType = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "u")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .filter(val -> !val.isEmpty())
                .orElse(null);
        boolean strike = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "strike");
        boolean doubleStrike = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "dstrike");
        boolean smallCaps = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "smallCaps");
        boolean allCaps = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "caps");
        boolean vanish = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "vanish");
        String color = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "color")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .filter(val -> !val.isEmpty())
                .orElse(null);
        String highlight = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "highlight")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .filter(val -> !val.isEmpty())
                .orElse(null);
        String vertAlign = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "vertAlign")
                .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                .filter(val -> !val.isEmpty())
                .orElse(null);
        Integer size = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "sz")
                .map(el -> XmlUtils.intAttribute(el, "w:val"))
                .orElse(null);
        Integer sizeCs = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "szCs")
                .map(el -> XmlUtils.intAttribute(el, "w:val"))
                .orElse(null);
        Map<String, String> fonts = new LinkedHashMap<>();
        XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "rFonts").ifPresent(fontsEl -> {
            for (String attr : List.of("ascii", "hAnsi", "eastAsia", "cs")) {
                String val = fontsEl.getAttributeNS(Namespaces.WORD_MAIN, attr);
                if (val != null && !val.isEmpty()) {
                    fonts.put(attr, val);
                }
            }
        });
        return new WordDocument.RunProperties(styleId, bold, italic, underline, underlineType, strike,
                doubleStrike, smallCaps, allCaps, vanish, color, highlight, vertAlign,
                size, sizeCs, fonts, rPr);
    }

    private WordDocument.Inline parseSymbol(Element element) {
        String font = element.getAttributeNS(Namespaces.WORD_MAIN, "font");
        if (font.isEmpty()) {
            String fallback = element.getAttribute("w:font");
            if (!fallback.isEmpty()) {
                font = fallback.replace("w:", "");
            }
        }
        String charCode = element.getAttributeNS(Namespaces.WORD_MAIN, "char");
        if (charCode.isEmpty()) {
            charCode = element.getAttribute("w:char");
        }
        return new WordDocument.Symbol(font.isEmpty() ? null : font, charCode);
    }

    private WordDocument.Inline parseBreak(Element element) {
        String type = element.getAttributeNS(Namespaces.WORD_MAIN, "type");
        WordDocument.Break.Type breakType = switch (type) {
            case "page" -> WordDocument.Break.Type.PAGE;
            case "column" -> WordDocument.Break.Type.COLUMN;
            default -> WordDocument.Break.Type.LINE;
        };
        String clear = element.getAttributeNS(Namespaces.WORD_MAIN, "clear");
        clear = clear.isEmpty() ? null : clear;
        return new WordDocument.Break(breakType, clear);
    }

    private WordDocument.Inline parseFieldChar(Element element) {
        String type = element.getAttributeNS(Namespaces.WORD_MAIN, "fldCharType");
        WordDocument.FieldCharacter.CharacterType charType = switch (type) {
            case "begin" -> WordDocument.FieldCharacter.CharacterType.BEGIN;
            case "separate" -> WordDocument.FieldCharacter.CharacterType.SEPARATE;
            case "end" -> WordDocument.FieldCharacter.CharacterType.END;
            default -> WordDocument.FieldCharacter.CharacterType.BEGIN;
        };
        return new WordDocument.FieldCharacter(charType);
    }

    private WordDocument.Inline parseDrawing(Element element) {
        Element drawing = XmlUtils.childElements(element).stream().findFirst().orElse(null);
        if (drawing == null) {
            return new WordDocument.Drawing(null, null, 0, 0, true);
        }
        boolean inline = "inline".equals(drawing.getLocalName());
        Element extent = XmlUtils.firstChild(drawing, Namespaces.DRAWINGML_WORDPROCESSING, "extent").orElse(null);
        long width = extent != null ? parseLong(extent.getAttribute("cx")) : 0;
        long height = extent != null ? parseLong(extent.getAttribute("cy")) : 0;
        Element docPr = XmlUtils.firstChild(drawing, Namespaces.DRAWINGML_WORDPROCESSING, "docPr").orElse(null);
        String descr = docPr != null ? docPr.getAttribute("descr") : null;
        Element graphic = XmlUtils.firstChild(drawing, Namespaces.DRAWINGML_MAIN, "graphic").orElse(null);
        String relId = null;
        if (graphic != null) {
            Element graphicData = XmlUtils.firstChild(graphic, Namespaces.DRAWINGML_MAIN, "graphicData").orElse(null);
            if (graphicData != null) {
                Element pic = XmlUtils.firstChild(graphicData, Namespaces.DRAWINGML_PIC, "pic").orElse(null);
                if (pic != null) {
                    Element blipFill = XmlUtils.firstChild(pic, Namespaces.DRAWINGML_PIC, "blipFill").orElse(null);
                    if (blipFill != null) {
                        Element blip = XmlUtils.firstChild(blipFill, Namespaces.DRAWINGML_MAIN, "blip").orElse(null);
                        if (blip != null) {
                            relId = blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
                        }
                    }
                }
            }
        }
        return new WordDocument.Drawing(relId, descr, width, height, inline);
    }

    private WordDocument.Inline parsePict(Element element) {
        for (Element child : XmlUtils.childElements(element)) {
            if (Namespaces.VML.equals(child.getNamespaceURI()) && "shape".equals(child.getLocalName())) {
                for (Element pictChild : XmlUtils.childElements(child)) {
                    if (Namespaces.VML.equals(pictChild.getNamespaceURI()) && "imagedata".equals(pictChild.getLocalName())) {
                        String relId = pictChild.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
                        String descr = pictChild.getAttribute("o:title");
                        return new WordDocument.Drawing(relId, descr, 0, 0, true);
                    }
                }
            }
        }
        return null;
    }

    private long parseLong(String value) {
        if (value == null || value.isEmpty()) {
            return 0;
        }
        try {
            return Long.parseLong(value);
        } catch (NumberFormatException e) {
            return 0;
        }
    }
}
