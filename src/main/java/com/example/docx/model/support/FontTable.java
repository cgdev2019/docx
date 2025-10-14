package com.example.docx.model.support;

import org.w3c.dom.Element;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

/**
 * Fonts declared in {@code word/fontTable.xml}.
 */
public final class FontTable {

    private final Map<String, Font> fonts;

    public FontTable(Map<String, Font> fonts) {
        this.fonts = Collections.unmodifiableMap(new LinkedHashMap<>(fonts));
    }

    public Map<String, Font> fonts() {
        return fonts;
    }

    public Optional<Font> font(String name) {
        return Optional.ofNullable(fonts.get(name));
    }

    public static final class Font {
        private final String name;
        private final Map<String, String> charsetMapping;
        private final Element raw;

        public Font(String name, Map<String, String> charsetMapping, Element raw) {
            this.name = name;
            this.charsetMapping = Map.copyOf(charsetMapping);
            this.raw = raw;
        }

        public String name() {
            return name;
        }

        public Map<String, String> charsetMapping() {
            return charsetMapping;
        }

        public Optional<Element> raw() {
            return Optional.ofNullable(raw);
        }
    }
}
