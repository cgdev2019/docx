package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.numbering.NumberingDefinitions;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Parses {@code word/numbering.xml}.
 */
public final class NumberingParser {

    public NumberingDefinitions parse(DocxArchive archive) throws IOException {
        if (!archive.exists("word/numbering.xml")) {
            return null;
        }
        try (InputStream input = archive.open("word/numbering.xml")) {
            Document document = XmlUtils.parse(input);
            Element root = document.getDocumentElement();
            NumberingDefinitions.Builder builder = NumberingDefinitions.builder();
            for (Element abstractNum : XmlUtils.children(root, Namespaces.WORD_MAIN, "abstractNum")) {
                int id = Integer.parseInt(abstractNum.getAttributeNS(Namespaces.WORD_MAIN, "abstractNumId"));
                Map<Integer, NumberingDefinitions.Level> levels = new LinkedHashMap<>();
                for (Element levelEl : XmlUtils.children(abstractNum, Namespaces.WORD_MAIN, "lvl")) {
                    int level = Integer.parseInt(levelEl.getAttributeNS(Namespaces.WORD_MAIN, "ilvl"));
                    String fmt = XmlUtils.firstChild(levelEl, Namespaces.WORD_MAIN, "numFmt")
                            .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                            .orElse(null);
                    String text = XmlUtils.firstChild(levelEl, Namespaces.WORD_MAIN, "lvlText")
                            .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                            .orElse(null);
                    Integer start = XmlUtils.firstChild(levelEl, Namespaces.WORD_MAIN, "start")
                            .map(el -> XmlUtils.intAttribute(el, "w:val"))
                            .orElse(null);
                    Integer restart = XmlUtils.firstChild(levelEl, Namespaces.WORD_MAIN, "lvlRestart")
                            .map(el -> XmlUtils.intAttribute(el, "w:val"))
                            .orElse(null);
                    levels.put(level, new NumberingDefinitions.Level(level, fmt, text, start, restart, levelEl));
                }
                builder.addAbstractNumbering(new NumberingDefinitions.AbstractNumbering(id, levels, abstractNum));
            }
            for (Element num : XmlUtils.children(root, Namespaces.WORD_MAIN, "num")) {
                int id = Integer.parseInt(num.getAttributeNS(Namespaces.WORD_MAIN, "numId"));
                int abstractId = XmlUtils.firstChild(num, Namespaces.WORD_MAIN, "abstractNumId")
                        .map(el -> Integer.parseInt(el.getAttributeNS(Namespaces.WORD_MAIN, "val")))
                        .orElse(-1);
                Map<Integer, NumberingDefinitions.LevelOverride> overrides = new LinkedHashMap<>();
                for (Element overrideEl : XmlUtils.children(num, Namespaces.WORD_MAIN, "lvlOverride")) {
                    int level = Integer.parseInt(overrideEl.getAttributeNS(Namespaces.WORD_MAIN, "ilvl"));
                    Integer startOverride = XmlUtils.firstChild(overrideEl, Namespaces.WORD_MAIN, "startOverride")
                            .map(el -> XmlUtils.intAttribute(el, "w:val"))
                            .orElse(null);
                    Element levelEl = XmlUtils.firstChild(overrideEl, Namespaces.WORD_MAIN, "lvl").orElse(null);
                    NumberingDefinitions.Level definition = null;
                    if (levelEl != null) {
                        definition = new NumberingDefinitions.Level(level,
                                XmlUtils.firstChild(levelEl, Namespaces.WORD_MAIN, "numFmt")
                                        .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                                        .orElse(null),
                                XmlUtils.firstChild(levelEl, Namespaces.WORD_MAIN, "lvlText")
                                        .map(el -> el.getAttributeNS(Namespaces.WORD_MAIN, "val"))
                                        .orElse(null),
                                XmlUtils.firstChild(levelEl, Namespaces.WORD_MAIN, "start")
                                        .map(el -> XmlUtils.intAttribute(el, "w:val"))
                                        .orElse(null),
                                XmlUtils.firstChild(levelEl, Namespaces.WORD_MAIN, "lvlRestart")
                                        .map(el -> XmlUtils.intAttribute(el, "w:val"))
                                        .orElse(null),
                                levelEl);
                    }
                    overrides.put(level, new NumberingDefinitions.LevelOverride(level, startOverride, definition));
                }
                builder.addNumberingInstance(new NumberingDefinitions.NumberingInstance(id, abstractId, overrides, num));
            }
            return builder.build();
        }
    }
}
