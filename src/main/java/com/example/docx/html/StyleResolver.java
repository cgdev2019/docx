package com.example.docx.html;

import com.example.docx.model.document.WordDocument;
import com.example.docx.model.styles.StyleDefinitions;
import com.example.docx.parser.Namespaces;
import com.example.docx.parser.XmlUtils;
import org.w3c.dom.Element;

import java.util.*;
import java.util.stream.Collectors;

public final class StyleResolver {
    private final StyleDefinitions definitions;
    private final Map<String, StyleDefinitions.Style> stylesById;
    private final List<WordDocument.RunProperties> defaultCharacterRunProperties;
    private final WordDocument.ParagraphProperties docDefaultParagraphProperties;
    private final WordDocument.RunProperties docDefaultRunProperties;
    private final Map<String, String> themeColors;

    StyleResolver(StyleDefinitions definitions, Map<String, String> themeColors) {
        this.definitions = definitions == null ? StyleDefinitions.empty() : definitions;
        this.themeColors = themeColors == null ? Map.of() : Map.copyOf(themeColors);
        this.stylesById = new LinkedHashMap<>(this.definitions.styles());
        this.defaultCharacterRunProperties = this.definitions.styles().values().stream()
                .filter(style -> "character".equals(style.type()))
                .filter(StyleDefinitions.Style::defaultStyle)
                .map(style -> style.runProperties().orElse(null))
                .filter(Objects::nonNull)
                .collect(Collectors.toList());
        Element docDefaults = this.definitions.rawDocumentDefaults().orElse(null);
        this.docDefaultParagraphProperties = docDefaults == null ? null : parseParagraphDefaults(docDefaults);
        this.docDefaultRunProperties = docDefaults == null ? null : parseRunDefaults(docDefaults);
    }

    ResolvedParagraph resolveParagraph(WordDocument.ParagraphProperties properties,
                                       List<WordDocument.RunProperties> extraRunFallbacks) {
        WordDocument.ParagraphProperties effective = properties == null ? DocxToHtml.EMPTY_PARAGRAPH_PROPERTIES : properties;
        String effectiveStyleId = effective.styleId().orElse(null);
        List<StyleDefinitions.Style> paragraphChain = styleChain(effectiveStyleId);
        LinkedHashSet<String> visited = paragraphChain.stream()
                .map(StyleDefinitions.Style::styleId)
                .collect(Collectors.toCollection(LinkedHashSet::new));
        for (String defaultId : definitions.defaultParagraphStyleHierarchy()) {
            if (visited.add(defaultId)) {
                StyleDefinitions.Style style = stylesById.get(defaultId);
                if (style != null) {
                    paragraphChain.add(style);
                }
            }
        }

        List<WordDocument.ParagraphProperties> paragraphFallbacks = new ArrayList<>();
        List<WordDocument.RunProperties> runFallbacks = new ArrayList<>();
        List<WordDocument.RunProperties> additionalFallbacks = extraRunFallbacks == null ? List.of() : extraRunFallbacks;

        paragraphRunProperties(effective).ifPresent(runFallbacks::add);

        for (StyleDefinitions.Style style : paragraphChain) {
            style.paragraphProperties().ifPresent(paragraphFallbacks::add);
            style.runProperties().ifPresent(runFallbacks::add);
            style.paragraphProperties()
                    .flatMap(StyleResolver::paragraphRunProperties)
                    .ifPresent(runFallbacks::add);
            style.link().ifPresent(linkId -> {
                StyleDefinitions.Style linked = stylesById.get(linkId);
                if (linked != null) {
                    linked.runProperties().ifPresent(runFallbacks::add);
                }
            });
        }
        if (docDefaultParagraphProperties != null) {
            paragraphFallbacks.add(docDefaultParagraphProperties);
            paragraphRunProperties(docDefaultParagraphProperties).ifPresent(runFallbacks::add);
        }
        if (!additionalFallbacks.isEmpty()) {
            runFallbacks.addAll(additionalFallbacks);
        }
        runFallbacks.addAll(defaultCharacterRunProperties);
        if (docDefaultRunProperties != null) {
            runFallbacks.add(docDefaultRunProperties);
        }
        if (effective.styleId().filter("Title"::equals).isPresent()) {
        }

        WordDocument.Alignment alignment = effective.alignment().orElseGet(() ->
                paragraphFallbacks.stream()
                        .map(p -> p.alignment().orElse(null))
                        .filter(Objects::nonNull)
                        .findFirst()
                        .orElse(null));
        WordDocument.Indentation indentation = effective.indentation().orElseGet(() ->
                paragraphFallbacks.stream()
                        .map(p -> p.indentation().orElse(null))
                        .filter(Objects::nonNull)
                        .findFirst()
                        .orElse(null));
        WordDocument.Spacing spacing = effective.spacing().orElseGet(() ->
                paragraphFallbacks.stream()
                        .map(p -> p.spacing().orElse(null))
                        .filter(Objects::nonNull)
                        .findFirst()
                        .orElse(null));
        boolean keepTogether = effective.keepTogether() || paragraphFallbacks.stream().anyMatch(WordDocument.ParagraphProperties::keepTogether);
        boolean keepWithNext = effective.keepWithNext() || paragraphFallbacks.stream().anyMatch(WordDocument.ParagraphProperties::keepWithNext);
        boolean pageBreakBefore = effective.pageBreakBefore() || paragraphFallbacks.stream().anyMatch(WordDocument.ParagraphProperties::pageBreakBefore);
        String shading = DocxToHtml.paragraphShadingColor(effective, themeColors);
        if (shading == null) {
            for (WordDocument.ParagraphProperties fallback : paragraphFallbacks) {
                shading = DocxToHtml.paragraphShadingColor(fallback, themeColors);
                if (shading != null) {
                    break;
                }
            }
        }

        return new ResolvedParagraph(alignment, indentation, spacing, keepTogether, keepWithNext, pageBreakBefore, shading, List.copyOf(runFallbacks));
    }

    ResolvedRun resolveRun(WordDocument.RunProperties properties, ResolvedParagraph paragraph) {
        WordDocument.RunProperties effective = properties == null ? DocxToHtml.EMPTY_RUN_PROPERTIES : properties;
        List<WordDocument.RunProperties> cascade = new ArrayList<>();
        cascade.add(effective);
        effective.styleId().ifPresent(styleId -> {
            for (StyleDefinitions.Style style : styleChain(styleId)) {
                if ("character".equals(style.type())) {
                    style.runProperties().ifPresent(cascade::add);
                }
            }
        });
        cascade.addAll(paragraph.runFallbacks());

        boolean bold = false;
        boolean italic = false;
        boolean underline = false;
        String underlineType = null;
        boolean strike = false;
        boolean doubleStrike = false;
        boolean smallCaps = false;
        boolean allCaps = false;
        boolean vanish = false;
        String color = null;
        String highlight = null;
        String verticalAlign = null;
        Integer size = null;
        Integer complexSize = null;
        Map<String, String> fonts = new LinkedHashMap<>();

        for (WordDocument.RunProperties runProperties : cascade) {
            if (runProperties.bold()) {
                bold = true;
            }
            if (runProperties.italic()) {
                italic = true;
            }
            if (runProperties.underline()) {
                if (!underline) {
                    underline = true;
                    underlineType = runProperties.underlineType().orElse(underlineType);
                } else if (underlineType == null) {
                    underlineType = runProperties.underlineType().orElse(null);
                }
            }
            if (runProperties.strike()) {
                strike = true;
            }
            if (runProperties.doubleStrike()) {
                doubleStrike = true;
                strike = true;
            }
            if (runProperties.smallCaps()) {
                smallCaps = true;
            }
            if (runProperties.allCaps()) {
                allCaps = true;
            }
            if (runProperties.vanish()) {
                vanish = true;
            }
            if (color == null) {
                color = runProperties.color().orElse(null);
            }
            if (highlight == null) {
                highlight = runProperties.highlight().orElse(null);
            }
            if (verticalAlign == null) {
                verticalAlign = runProperties.verticalAlignment().orElse(null);
            }
            if (size == null) {
                size = runProperties.size().orElse(null);
            }
            if (complexSize == null) {
                complexSize = runProperties.complexScriptSize().orElse(null);
            }
            if (!runProperties.fonts().isEmpty()) {
                runProperties.fonts().forEach(fonts::putIfAbsent);
            }
        }
        if (size == null && complexSize != null) {
            size = complexSize;
        }
        List<String> fontStack = DocxToHtml.computeFontStack(fonts);
        return new ResolvedRun(bold, italic, underline, underlineType, strike, doubleStrike,
                smallCaps, allCaps, vanish, Optional.ofNullable(color),
                Optional.ofNullable(highlight), Optional.ofNullable(verticalAlign),
                Optional.ofNullable(size), fontStack);
    }

    private List<StyleDefinitions.Style> styleChain(String styleId) {
        if (styleId == null || styleId.isBlank()) {
            return new ArrayList<>();
        }
        List<StyleDefinitions.Style> chain = new ArrayList<>();
        Set<String> seen = new LinkedHashSet<>();
        String current = styleId;
        while (current != null && seen.add(current)) {
            StyleDefinitions.Style style = stylesById.get(current);
            if (style == null) {
                break;
            }
            chain.add(style);
            current = style.basedOn().orElse(null);
        }
        return chain;
    }

    private static WordDocument.ParagraphProperties parseParagraphDefaults(Element docDefaults) {
        Element pPrDefault = XmlUtils.firstChild(docDefaults, Namespaces.WORD_MAIN, "pPrDefault").orElse(null);
        if (pPrDefault == null) {
            return null;
        }
        Element pPr = XmlUtils.firstChild(pPrDefault, Namespaces.WORD_MAIN, "pPr").orElse(null);
        if (pPr == null) {
            return null;
        }
        return toParagraphProperties(pPr);
    }

    private static WordDocument.RunProperties parseRunDefaults(Element docDefaults) {
        Element rPrDefault = XmlUtils.firstChild(docDefaults, Namespaces.WORD_MAIN, "rPrDefault").orElse(null);
        if (rPrDefault == null) {
            return null;
        }
        Element rPr = XmlUtils.firstChild(rPrDefault, Namespaces.WORD_MAIN, "rPr").orElse(null);
        if (rPr == null) {
            return null;
        }
        return toRunProperties(rPr);
    }

    private static WordDocument.ParagraphProperties toParagraphProperties(Element pPr) {
        WordDocument.Alignment alignment = readAlignment(pPr);
        WordDocument.Indentation indentation = readIndentation(pPr);
        WordDocument.Spacing spacing = readSpacing(pPr);
        Integer outline = XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "outlineLvl")
                .map(el -> XmlUtils.intAttribute(el, "w:val"))
                .orElse(null);
        boolean keepTogether = XmlUtils.booleanElement(pPr, Namespaces.WORD_MAIN, "keepLines");
        boolean keepWithNext = XmlUtils.booleanElement(pPr, Namespaces.WORD_MAIN, "keepNext");
        boolean pageBreakBefore = XmlUtils.booleanElement(pPr, Namespaces.WORD_MAIN, "pageBreakBefore");
        List<WordDocument.TabStop> tabs = readTabs(pPr);
        return new WordDocument.ParagraphProperties(null, null, alignment, indentation, spacing,
                outline, keepTogether, keepWithNext, pageBreakBefore, tabs, pPr);
    }

    private static WordDocument.RunProperties toRunProperties(Element rPr) {
        if (rPr == null) {
            return null;
        }
        String styleId = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "rStyle")
                .map(el -> readWordAttribute(el, "val"))
                .orElse(null);
        boolean bold = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "b");
        boolean italic = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "i");
        boolean underline = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "u").isPresent();
        String underlineType = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "u")
                .map(el -> readWordAttribute(el, "val"))
                .orElse(null);
        boolean strike = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "strike");
        boolean doubleStrike = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "dstrike");
        boolean smallCaps = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "smallCaps");
        boolean allCaps = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "caps");
        boolean vanish = XmlUtils.booleanElement(rPr, Namespaces.WORD_MAIN, "vanish");
        String color = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "color")
                .map(el -> readWordAttribute(el, "val"))
                .orElse(null);
        String highlight = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "highlight")
                .map(el -> readWordAttribute(el, "val"))
                .orElse(null);
        String vertAlign = XmlUtils.firstChild(rPr, Namespaces.WORD_MAIN, "vertAlign")
                .map(el -> readWordAttribute(el, "val"))
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
                String value = readWordAttribute(fontsEl, attr);
                if (value != null) {
                    fonts.put(attr, value);
                }
            }
        });
        return new WordDocument.RunProperties(styleId, bold, italic, underline, underlineType,
                strike, doubleStrike, smallCaps, allCaps, vanish, color, highlight, vertAlign,
                size, sizeCs, fonts, rPr);
    }

    private static Optional<WordDocument.RunProperties> paragraphRunProperties(WordDocument.ParagraphProperties properties) {
        if (properties == null) {
            return Optional.empty();
        }
        return properties.rawProperties()
                .flatMap(pPr -> XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "rPr"))
                .map(StyleResolver::toRunProperties);
    }

    private static WordDocument.Alignment readAlignment(Element pPr) {
        return XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "jc")
                .map(el -> {
                    String val = readWordAttribute(el, "val");
                    if (val == null) {
                        return null;
                    }
                    return switch (val) {
                        case "left" -> WordDocument.Alignment.LEFT;
                        case "center" -> WordDocument.Alignment.CENTER;
                        case "right" -> WordDocument.Alignment.RIGHT;
                        case "both", "justify" -> WordDocument.Alignment.JUSTIFIED;
                        case "distribute" -> WordDocument.Alignment.DISTRIBUTE;
                        case "thaiDistribute" -> WordDocument.Alignment.THAI_DISTRIBUTED;
                        case "justLow" -> WordDocument.Alignment.JUSTIFY_LOW;
                        default -> null;
                    };
                })
                .orElse(null);
    }

    private static WordDocument.Indentation readIndentation(Element pPr) {
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

    private static WordDocument.Spacing readSpacing(Element pPr) {
        Element spacing = XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "spacing").orElse(null);
        if (spacing == null) {
            return null;
        }
        Integer before = XmlUtils.intAttribute(spacing, "w:before");
        Integer after = XmlUtils.intAttribute(spacing, "w:after");
        Integer line = XmlUtils.intAttribute(spacing, "w:line");
        String rule = readWordAttribute(spacing, "lineRule");
        return new WordDocument.Spacing(before, after, line, rule);
    }

    private static List<WordDocument.TabStop> readTabs(Element pPr) {
        Element tabsEl = XmlUtils.firstChild(pPr, Namespaces.WORD_MAIN, "tabs").orElse(null);
        if (tabsEl == null) {
            return List.of();
        }
        List<WordDocument.TabStop> result = new ArrayList<>();
        for (Element tab : XmlUtils.children(tabsEl, Namespaces.WORD_MAIN, "tab")) {
            String alignment = readWordAttribute(tab, "val");
            Integer position = XmlUtils.intAttribute(tab, "w:pos");
            result.add(new WordDocument.TabStop(alignment, position));
        }
        return result;
    }

    private static String readWordAttribute(Element element, String localName) {
        if (element == null) {
            return null;
        }
        String value = element.getAttributeNS(Namespaces.WORD_MAIN, localName);
        if (value == null || value.isEmpty()) {
            value = element.getAttribute("w:" + localName);
        }
        if (value == null || value.isBlank()) {
            return null;
        }
        return value;
    }

    ResolvedTableStyle resolveTableStyle(WordDocument.TableProperties properties) {
        if (properties == null) {
            return ResolvedTableStyle.empty();
        }
        String styleId = properties.styleId().orElse(null);
        if (styleId == null || styleId.isBlank()) {
            return ResolvedTableStyle.empty();
        }
        Map<TableStyleRegion, RegionStyle> regions = new EnumMap<>(TableStyleRegion.class);
        String tableBackground = null;

        for (StyleDefinitions.Style style : styleChain(styleId)) {
            if (tableBackground == null) {
                tableBackground = style.tableProperties()
                        .flatMap(StyleDefinitions.TableStyleProperties::rawProperties)
                        .map(element -> DocxToHtml.shadingColorFromElement(element, themeColors))
                        .orElse(null);
            }
            style.rawStyle().ifPresent(raw -> {
                for (Element tblStylePr : XmlUtils.children(raw, Namespaces.WORD_MAIN, "tblStylePr")) {
                    String type = tblStylePr.getAttributeNS(Namespaces.WORD_MAIN, "type");
                    TableStyleRegion region = TableStyleRegion.from(type);
                    if (region == null || regions.containsKey(region)) {
                        continue;
                    }
                    RegionStyle regionStyle = parseTableRegion(tblStylePr);
                    if (regionStyle != null) {
                        regions.put(region, regionStyle);
                    }
                }
            });
        }
        return new ResolvedTableStyle(tableBackground, regions);
    }

    private RegionStyle parseTableRegion(Element tblStylePr) {
        if (tblStylePr == null) {
            return null;
        }
        Element tcPr = XmlUtils.firstChild(tblStylePr, Namespaces.WORD_MAIN, "tcPr").orElse(null);
        String background = DocxToHtml.shadingColorFromElement(tcPr, themeColors);
        if (background == null) {
            Element tblPr = XmlUtils.firstChild(tblStylePr, Namespaces.WORD_MAIN, "tblPr").orElse(null);
            background = DocxToHtml.shadingColorFromElement(tblPr, themeColors);
        }
        Element rPr = XmlUtils.firstChild(tblStylePr, Namespaces.WORD_MAIN, "rPr").orElse(null);
        WordDocument.RunProperties runProperties = toRunProperties(rPr);
        if (background == null && runProperties == null) {
            return null;
        }
        return new RegionStyle(background, runProperties);
    }

    public enum TableStyleRegion {
        WHOLE_TABLE("wholeTable"),
        FIRST_ROW("firstRow"),
        LAST_ROW("lastRow"),
        FIRST_COLUMN("firstCol"),
        LAST_COLUMN("lastCol"),
        BAND1_HORZ("band1Horz"),
        BAND2_HORZ("band2Horz"),
        BAND1_VERT("band1Vert"),
        BAND2_VERT("band2Vert");

        private final String type;

        TableStyleRegion(String type) {
            this.type = type;
        }

        static TableStyleRegion from(String type) {
            if (type == null || type.isBlank()) {
                return null;
            }
            for (TableStyleRegion region : values()) {
                if (region.type.equals(type)) {
                    return region;
                }
            }
            return null;
        }
    }

    public record RegionStyle(String backgroundColor,
                               WordDocument.RunProperties runProperties) {
    }

    public record ResolvedParagraph(WordDocument.Alignment alignment,
                             WordDocument.Indentation indentation,
                             WordDocument.Spacing spacing,
                             boolean keepTogether,
                             boolean keepWithNext,
                             boolean pageBreakBefore,
                             String shadingColor,
                             List<WordDocument.RunProperties> runFallbacks) {
    }

    public record ResolvedRun(boolean bold,
                       boolean italic,
                       boolean underline,
                       String underlineType,
                       boolean strike,
                       boolean doubleStrike,
                       boolean smallCaps,
                       boolean allCaps,
                       boolean vanish,
                       Optional<String> color,
                       Optional<String> highlight,
                       Optional<String> verticalAlignment,
                       Optional<Integer> fontSizeHalfPoints,
                       List<String> fontStack) {
    }
}
