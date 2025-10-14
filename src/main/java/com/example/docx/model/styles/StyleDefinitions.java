package com.example.docx.model.styles;

import com.example.docx.model.document.WordDocument;
import org.w3c.dom.Element;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * Style definitions contained in {@code word/styles.xml}.
 */
public final class StyleDefinitions {

    private final Map<String, Style> styles;
    private final List<String> defaultParagraphStyleHierarchy;
    private final Element rawDocumentDefaults;

    public StyleDefinitions(Map<String, Style> styles,
                            List<String> defaultParagraphStyleHierarchy,
                            Element rawDocumentDefaults) {
        this.styles = Collections.unmodifiableMap(new LinkedHashMap<>(styles));
        this.defaultParagraphStyleHierarchy = List.copyOf(defaultParagraphStyleHierarchy);
        this.rawDocumentDefaults = rawDocumentDefaults;
    }

    public Map<String, Style> styles() {
        return styles;
    }

    public Optional<Style> styleById(String styleId) {
        return Optional.ofNullable(styles.get(styleId));
    }

    public List<String> defaultParagraphStyleHierarchy() {
        return defaultParagraphStyleHierarchy;
    }

    public Optional<Element> rawDocumentDefaults() {
        return Optional.ofNullable(rawDocumentDefaults);
    }

    public static StyleDefinitions empty() {
        return new StyleDefinitions(Map.of(), List.of(), null);
    }

    public static final class Style {
        private final String styleId;
        private final String type;
        private final String name;
        private final String basedOn;
        private final String next;
        private final String link;
        private final boolean defaultStyle;
        private final boolean customStyle;
        private final WordDocument.ParagraphProperties paragraphProperties;
        private final WordDocument.RunProperties runProperties;
        private final TableStyleProperties tableProperties;
        private final Element rawStyle;

        public Style(String styleId,
                     String type,
                     String name,
                     String basedOn,
                     String next,
                     String link,
                     boolean defaultStyle,
                     boolean customStyle,
                     WordDocument.ParagraphProperties paragraphProperties,
                     WordDocument.RunProperties runProperties,
                     TableStyleProperties tableProperties,
                     Element rawStyle) {
            this.styleId = styleId;
            this.type = type;
            this.name = name;
            this.basedOn = basedOn;
            this.next = next;
            this.link = link;
            this.defaultStyle = defaultStyle;
            this.customStyle = customStyle;
            this.paragraphProperties = paragraphProperties;
            this.runProperties = runProperties;
            this.tableProperties = tableProperties;
            this.rawStyle = rawStyle;
        }

        public String styleId() {
            return styleId;
        }

        public String type() {
            return type;
        }

        public Optional<String> name() {
            return Optional.ofNullable(name);
        }

        public Optional<String> basedOn() {
            return Optional.ofNullable(basedOn);
        }

        public Optional<String> next() {
            return Optional.ofNullable(next);
        }

        public Optional<String> link() {
            return Optional.ofNullable(link);
        }

        public boolean defaultStyle() {
            return defaultStyle;
        }

        public boolean customStyle() {
            return customStyle;
        }

        public Optional<WordDocument.ParagraphProperties> paragraphProperties() {
            return Optional.ofNullable(paragraphProperties);
        }

        public Optional<WordDocument.RunProperties> runProperties() {
            return Optional.ofNullable(runProperties);
        }

        public Optional<TableStyleProperties> tableProperties() {
            return Optional.ofNullable(tableProperties);
        }

        public Optional<Element> rawStyle() {
            return Optional.ofNullable(rawStyle);
        }
    }

    public static final class TableStyleProperties {
        private final Element rawProperties;

        public TableStyleProperties(Element rawProperties) {
            this.rawProperties = rawProperties;
        }

        public Optional<Element> rawProperties() {
            return Optional.ofNullable(rawProperties);
        }
    }
}
