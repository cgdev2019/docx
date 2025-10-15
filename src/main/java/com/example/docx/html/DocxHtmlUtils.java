package com.example.docx.html;

import com.example.docx.model.document.WordDocument;
import com.example.docx.parser.Namespaces;
import com.example.docx.parser.XmlUtils;
import org.w3c.dom.Element;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

final class DocxHtmlUtils {
    private DocxHtmlUtils() {
    }

    private static final Map<String, String> HIGHLIGHT_COLORS = createHighlightColorMap();

    static CssLength toCssLength(Integer twips) {
        if (twips == null) {
            return null;
        }
        return CssLength.points(twips / 20.0d);
    }

    static CssLength computeTextIndent(WordDocument.Indentation indentation) {
        if (indentation == null) {
            return null;
        }
        if (indentation.firstLine().isPresent()) {
            return toCssLength(indentation.firstLine().orElse(null));
        }
        if (indentation.hanging().isPresent()) {
            CssLength hanging = toCssLength(indentation.hanging().orElse(null));
            if (hanging != null) {
                return hanging.negate();
            }
        }
        return null;
    }

    static String computeLineHeight(WordDocument.Spacing spacing) {
        if (spacing == null) {
            return null;
        }
        Integer lineValue = spacing.line().orElse(null);
        if (lineValue == null || lineValue <= 0) {
            return null;
        }
        String rule = spacing.rule().orElse("auto");
        if ("auto".equalsIgnoreCase(rule)) {
            double multiple = lineValue / 240.0d;
            return formatDecimal(multiple);
        }
        double points = lineValue / 20.0d;
        return CssLength.points(points).css();
    }

    static String alignmentToCss(WordDocument.Alignment alignment) {
        if (alignment == null) {
            return null;
        }
        return switch (alignment) {
            case LEFT -> "left";
            case CENTER -> "center";
            case RIGHT -> "right";
            default -> "justify";
        };
    }

    static String formatDecimal(double value) {
        BigDecimal decimal = BigDecimal.valueOf(value).setScale(3, RoundingMode.HALF_UP).stripTrailingZeros();
        return decimal.toPlainString();
    }

    static double roundLength(double value) {
        return BigDecimal.valueOf(value).setScale(3, RoundingMode.HALF_UP).doubleValue();
    }

    static Optional<String> normalizeColor(String value) {
        if (value == null || value.isBlank()) {
            return Optional.empty();
        }
        String raw = value.trim();
        if ("auto".equalsIgnoreCase(raw)) {
            return Optional.empty();
        }
        if (raw.startsWith("#")) {
            raw = raw.substring(1);
        }
        if (raw.length() == 6 && raw.matches("[0-9a-fA-F]{6}")) {
            return Optional.of("#" + raw.toLowerCase(Locale.ROOT));
        }
        if (raw.length() == 3 && raw.matches("[0-9a-fA-F]{3}")) {
            StringBuilder expanded = new StringBuilder("#");
            for (char c : raw.toCharArray()) {
                char lower = Character.toLowerCase(c);
                expanded.append(lower).append(lower);
            }
            return Optional.of(expanded.toString());
        }
        try {
            int numeric = Integer.parseInt(raw, 16);
            return Optional.of(String.format(Locale.ROOT, "#%06x", numeric));
        } catch (NumberFormatException ignored) {
        }
        try {
            int numeric = Integer.parseInt(raw);
            return Optional.of(String.format(Locale.ROOT, "#%06x", numeric));
        } catch (NumberFormatException ignored) {
        }
        return Optional.empty();
    }

    static Optional<String> normalizeHighlight(String value) {
        if (value == null || value.isBlank() || "none".equalsIgnoreCase(value)) {
            return Optional.empty();
        }
        String key = value.trim().toLowerCase(Locale.ROOT);
        if (HIGHLIGHT_COLORS.containsKey(key)) {
            return Optional.of(HIGHLIGHT_COLORS.get(key));
        }
        return normalizeColor(value);
    }

    static String normalizeVerticalAlignment(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        return switch (value) {
            case "superscript" -> "super";
            case "subscript" -> "sub";
            default -> null;
        };
    }

    static String decorationStyle(String underlineType, boolean doubleStrike) {
        if (doubleStrike) {
            return "double";
        }
        if (underlineType == null || underlineType.isBlank()) {
            return null;
        }
        return switch (underlineType.toLowerCase(Locale.ROOT)) {
            case "double" -> "double";
            case "dotted", "dotdash", "dotdotdash" -> "dotted";
            case "dash", "dashdot", "dashdotdot" -> "dashed";
            case "wave" -> "wavy";
            case "thick" -> "solid";
            default -> null;
        };
    }

    static List<String> computeFontStack(Map<String, String> fonts) {
        LinkedHashSet<String> unique = new LinkedHashSet<>();
        List<String> stack = new ArrayList<>();
        List<String> priority = List.of("ascii", "hAnsi", "eastAsia", "cs");
        for (String key : priority) {
            String font = fonts.get(key);
            if (font != null && unique.add(font)) {
                stack.add(font);
            }
        }
        for (String font : fonts.values()) {
            if (font != null && unique.add(font)) {
                stack.add(font);
            }
        }
        return stack;
    }

    static String fontStackToCss(List<String> fonts) {
        if (fonts == null || fonts.isEmpty()) {
            return null;
        }
        List<String> list = new ArrayList<>();
        for (String font : fonts) {
            if (font == null || font.isBlank()) {
                continue;
            }
            String trimmed = font.trim();
            if (trimmed.contains(" ") && !trimmed.startsWith("\"")) {
                list.add('"' + trimmed + '"');
            } else {
                list.add(trimmed);
            }
        }
        if (list.isEmpty()) {
            return null;
        }
        return String.join(",", list);
    }

    static String escapeHtml(String value) {
        if (value == null) {
            return "";
        }
        StringBuilder builder = new StringBuilder(value.length());
        for (char c : value.toCharArray()) {
            switch (c) {
                case '<' -> builder.append("&lt;");
                case '>' -> builder.append("&gt;");
                case '&' -> builder.append("&amp;");
                case '"' -> builder.append("&quot;");
                case '\'' -> builder.append("&#39;");
                default -> builder.append(c);
            }
        }
        return builder.toString();
    }

    static String escapeHtmlAttribute(String value) {
        if (value == null) {
            return "";
        }
        StringBuilder builder = new StringBuilder(value.length());
        for (char c : value.toCharArray()) {
            switch (c) {
                case '<' -> builder.append("&lt;");
                case '>' -> builder.append("&gt;");
                case '&' -> builder.append("&amp;");
                case '"' -> builder.append("&quot;");
                default -> builder.append(c);
            }
        }
        return builder.toString();
    }

    static String paragraphShadingColor(WordDocument.ParagraphProperties properties,
                                        Map<String, String> themeColors) {
        if (properties == null) {
            return null;
        }
        return properties.rawProperties()
                .map(element -> shadingColorFromElement(element, themeColors))
                .orElse(null);
    }

    static String tableShadingColor(WordDocument.TableProperties properties,
                                    Map<String, String> themeColors) {
        if (properties == null) {
            return null;
        }
        return properties.rawProperties()
                .map(element -> shadingColorFromElement(element, themeColors))
                .orElse(null);
    }

    static String tableRowShadingColor(WordDocument.TableRowProperties properties,
                                       Map<String, String> themeColors) {
        if (properties == null) {
            return null;
        }
        return properties.rawProperties()
                .map(element -> shadingColorFromElement(element, themeColors))
                .orElse(null);
    }

    static String tableCellShadingColor(WordDocument.TableCellProperties properties,
                                        Map<String, String> themeColors) {
        if (properties == null) {
            return null;
        }
        return properties.rawProperties()
                .map(element -> shadingColorFromElement(element, themeColors))
                .orElse(null);
    }

    static TableBorders tableBorders(WordDocument.TableProperties properties,
                                     Map<String, String> themeColors) {
        if (properties == null) {
            return TableBorders.empty();
        }
        return properties.rawProperties()
                .map(element -> tableBordersFromRaw(element, themeColors))
                .orElse(TableBorders.empty());
    }

    static BorderDefinition tableCellBorders(WordDocument.TableCellProperties properties,
                                             Map<String, String> themeColors) {
        if (properties == null) {
            return BorderDefinition.empty();
        }
        return properties.rawProperties()
                .map(element -> bordersFromChild(element, "tcBorders", themeColors))
                .orElse(BorderDefinition.empty());
    }

    static BorderDefinition paragraphBorders(WordDocument.ParagraphProperties properties,
                                             Map<String, String> themeColors) {
        if (properties == null) {
            return BorderDefinition.empty();
        }
        return properties.rawProperties()
                .map(element -> bordersFromChild(element, "pBdr", themeColors))
                .orElse(BorderDefinition.empty());
    }

    static BorderDefinition runBorders(WordDocument.RunProperties properties,
                                       Map<String, String> themeColors) {
        if (properties == null) {
            return BorderDefinition.empty();
        }
        return properties.rawProperties()
                .map(element -> {
                    Element border = XmlUtils.firstChild(element, Namespaces.WORD_MAIN, "bdr").orElse(null);
                    if (border == null) {
                        return BorderDefinition.empty();
                    }
                    BorderDefinition.BorderEdge edge = parseBorderEdge(border, themeColors);
                    if (edge == null) {
                        return BorderDefinition.empty();
                    }
                    return BorderDefinition.of(edge, edge, edge, edge);
                })
                .orElse(BorderDefinition.empty());
    }

    static String shadingColorFromElement(Element parent, Map<String, String> themeColors) {
        if (parent == null) {
            return null;
        }
        Element shading = XmlUtils.firstChild(parent, Namespaces.WORD_MAIN, "shd").orElse(null);
        if (shading == null) {
            return null;
        }
        String fill = firstNonBlank(XmlUtils.attribute(shading, "w:fill"), XmlUtils.attribute(shading, "fill"));
        String direct = normalizeDirectShadingFill(fill);
        if (direct != null) {
            return direct;
        }
        String themeFill = firstNonBlank(XmlUtils.attribute(shading, "w:themeFill"), XmlUtils.attribute(shading, "themeFill"));
        if (themeFill != null) {
            String base = resolveThemeColor(themeFill, themeColors);
            if (base != null) {
                String tint = firstNonBlank(XmlUtils.attribute(shading, "w:themeFillTint"), XmlUtils.attribute(shading, "themeFillTint"));
                String shade = firstNonBlank(XmlUtils.attribute(shading, "w:themeFillShade"), XmlUtils.attribute(shading, "themeFillShade"));
                return applyThemeTintShade(base, tint, shade);
            }
        }
        String themeColor = firstNonBlank(XmlUtils.attribute(shading, "w:themeColor"), XmlUtils.attribute(shading, "themeColor"));
        if (themeColor != null) {
            String base = resolveThemeColor(themeColor, themeColors);
            if (base != null) {
                String tint = firstNonBlank(XmlUtils.attribute(shading, "w:themeTint"), XmlUtils.attribute(shading, "themeTint"));
                String shade = firstNonBlank(XmlUtils.attribute(shading, "w:themeShade"), XmlUtils.attribute(shading, "themeShade"));
                return applyThemeTintShade(base, tint, shade);
            }
        }
        return null;
    }

    static TableBorders tableBordersFromRaw(Element properties, Map<String, String> themeColors) {
        if (properties == null) {
            return TableBorders.empty();
        }
        Element borders = XmlUtils.firstChild(properties, Namespaces.WORD_MAIN, "tblBorders").orElse(null);
        if (borders == null) {
            return TableBorders.empty();
        }
        BorderDefinition perimeter = BorderDefinition.of(
                parseBorderEdge(XmlUtils.firstChild(borders, Namespaces.WORD_MAIN, "top").orElse(null), themeColors),
                parseBorderEdge(XmlUtils.firstChild(borders, Namespaces.WORD_MAIN, "right").orElse(null), themeColors),
                parseBorderEdge(XmlUtils.firstChild(borders, Namespaces.WORD_MAIN, "bottom").orElse(null), themeColors),
                parseBorderEdge(XmlUtils.firstChild(borders, Namespaces.WORD_MAIN, "left").orElse(null), themeColors)
        );
        BorderDefinition insideH = BorderDefinition.empty();
        BorderDefinition insideV = BorderDefinition.empty();
        BorderDefinition.BorderEdge horiz = parseBorderEdge(XmlUtils.firstChild(borders, Namespaces.WORD_MAIN, "insideH").orElse(null), themeColors);
        if (horiz != null) {
            insideH = BorderDefinition.of(horiz, null, horiz, null);
        }
        BorderDefinition.BorderEdge vert = parseBorderEdge(XmlUtils.firstChild(borders, Namespaces.WORD_MAIN, "insideV").orElse(null), themeColors);
        if (vert != null) {
            insideV = BorderDefinition.of(null, vert, null, vert);
        }
        return new TableBorders(perimeter, insideH, insideV);
    }

    private static BorderDefinition bordersFromChild(Element parent, String localName, Map<String, String> themeColors) {
        Element borders = XmlUtils.firstChild(parent, Namespaces.WORD_MAIN, localName).orElse(null);
        if (borders == null) {
            return BorderDefinition.empty();
        }
        BorderDefinition.BorderEdge top = parseBorderEdge(XmlUtils.firstChild(borders, Namespaces.WORD_MAIN, "top").orElse(null), themeColors);
        BorderDefinition.BorderEdge right = parseBorderEdge(XmlUtils.firstChild(borders, Namespaces.WORD_MAIN, "right").orElse(null), themeColors);
        BorderDefinition.BorderEdge bottom = parseBorderEdge(XmlUtils.firstChild(borders, Namespaces.WORD_MAIN, "bottom").orElse(null), themeColors);
        BorderDefinition.BorderEdge left = parseBorderEdge(XmlUtils.firstChild(borders, Namespaces.WORD_MAIN, "left").orElse(null), themeColors);
        return BorderDefinition.of(top, right, bottom, left);
    }

    private static BorderDefinition.BorderEdge parseBorderEdge(Element element, Map<String, String> themeColors) {
        if (element == null) {
            return null;
        }
        String value = firstNonBlank(XmlUtils.attribute(element, "w:val"), XmlUtils.attribute(element, "val"));
        if (value == null || value.isBlank()) {
            return null;
        }
        String style = borderStyleToCss(value.trim().toLowerCase(Locale.ROOT));
        if (style == null || "none".equals(style)) {
            return null;
        }
        CssLength width = borderWidth(XmlUtils.intAttribute(element, "w:sz"));
        String color = resolveBorderColor(element, themeColors);
        return new BorderDefinition.BorderEdge(width, style, color);
    }

    private static CssLength borderWidth(Integer size) {
        if (size == null || size <= 0) {
            return CssLength.points(0.5d);
        }
        return CssLength.points(size / 8.0d);
    }

    private static String resolveBorderColor(Element element, Map<String, String> themeColors) {
        String direct = firstNonBlank(XmlUtils.attribute(element, "w:color"), XmlUtils.attribute(element, "color"));
        String color = normalizeColor(direct).orElse(null);
        if (color != null) {
            return color;
        }
        String themeColor = firstNonBlank(XmlUtils.attribute(element, "w:themeColor"), XmlUtils.attribute(element, "themeColor"));
        if (themeColor != null) {
            String base = resolveThemeColor(themeColor, themeColors);
            if (base != null) {
                String tint = firstNonBlank(XmlUtils.attribute(element, "w:themeTint"), XmlUtils.attribute(element, "themeTint"));
                String shade = firstNonBlank(XmlUtils.attribute(element, "w:themeShade"), XmlUtils.attribute(element, "themeShade"));
                return applyThemeTintShade(base, tint, shade);
            }
        }
        return null;
    }

    private static String borderStyleToCss(String value) {
        return switch (value) {
            case "nil", "none" -> "none";
            case "double", "triple", "thickthinmediumgap", "thinthickmediumgap", "thickbetweenthinmediumgap",
                    "thinthickthinmediumgap", "thickthinlargegap", "thinthicklargegap",
                    "thickbetweenthinlargegap", "thinthickthinlargegap", "doublewave" -> "double";
            case "dotted", "dotdash", "dotdotdash", "dashdot", "dashdotdot" -> "dotted";
            case "dashed", "dashsmallgap", "dashlargegap" -> "dashed";
            default -> "solid";
        };
    }

    private static Map<String, String> createHighlightColorMap() {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("yellow", "#ffff00");
        map.put("green", "#00ff00");
        map.put("cyan", "#00ffff");
        map.put("magenta", "#ff00ff");
        map.put("blue", "#0000ff");
        map.put("red", "#ff0000");
        map.put("darkblue", "#00008b");
        map.put("darkcyan", "#008b8b");
        map.put("darkgreen", "#006400");
        map.put("darkmagenta", "#8b008b");
        map.put("darkred", "#8b0000");
        map.put("darkyellow", "#b8860b");
        map.put("gray50", "#808080");
        map.put("gray25", "#c0c0c0");
        map.put("black", "#000000");
        map.put("white", "#ffffff");
        map.put("teal", "#008080");
        map.put("violet", "#8000ff");
        map.put("orange", "#ffa500");
        return Map.copyOf(map);
    }

    private static String normalizeDirectShadingFill(String fill) {
        if (fill == null) {
            return null;
        }
        String trimmed = fill.trim();
        if (trimmed.isEmpty()) {
            return null;
        }
        if ("auto".equalsIgnoreCase(trimmed) || "none".equalsIgnoreCase(trimmed)) {
            return null;
        }
        return normalizeColor(trimmed).orElse(null);
    }

    private static String resolveThemeColor(String key, Map<String, String> themeColors) {
        if (key == null || key.isBlank() || themeColors == null || themeColors.isEmpty()) {
            return null;
        }
        return themeColors.get(key.trim().toLowerCase(Locale.ROOT));
    }

    private static String firstNonBlank(String first, String second) {
        if (first != null && !first.isBlank()) {
            return first;
        }
        if (second != null && !second.isBlank()) {
            return second;
        }
        return null;
    }

    private static String applyThemeTintShade(String baseColor, String tintHex, String shadeHex) {
        if (baseColor == null || baseColor.length() != 7 || !baseColor.startsWith("#")) {
            return baseColor;
        }
        int r = Integer.parseInt(baseColor.substring(1, 3), 16);
        int g = Integer.parseInt(baseColor.substring(3, 5), 16);
        int b = Integer.parseInt(baseColor.substring(5, 7), 16);
        if (tintHex != null && !tintHex.isBlank()) {
            try {
                double tint = Integer.parseInt(tintHex, 16) / 255.0d;
                r = applyTint(r, tint);
                g = applyTint(g, tint);
                b = applyTint(b, tint);
            } catch (NumberFormatException ignored) {
            }
        }
        if (shadeHex != null && !shadeHex.isBlank()) {
            try {
                double shade = Integer.parseInt(shadeHex, 16) / 255.0d;
                r = applyShade(r, shade);
                g = applyShade(g, shade);
                b = applyShade(b, shade);
            } catch (NumberFormatException ignored) {
            }
        }
        return String.format(Locale.ROOT, "#%02x%02x%02x", clampColor(r), clampColor(g), clampColor(b));
    }

    private static int applyTint(int channel, double tint) {
        return clampColor((int) Math.round(channel + (255 - channel) * tint));
    }

    private static int applyShade(int channel, double shade) {
        return clampColor((int) Math.round(channel * (1.0d - shade)));
    }

    private static int clampColor(int value) {
        if (value < 0) {
            return 0;
        }
        if (value > 255) {
            return 255;
        }
        return value;
    }

    static class TableBorders {
        private final BorderDefinition perimeter;
        private final BorderDefinition insideHorizontal;
        private final BorderDefinition insideVertical;

        private TableBorders(BorderDefinition perimeter,
                             BorderDefinition insideHorizontal,
                             BorderDefinition insideVertical) {
            this.perimeter = perimeter == null ? BorderDefinition.empty() : perimeter;
            this.insideHorizontal = insideHorizontal == null ? BorderDefinition.empty() : insideHorizontal;
            this.insideVertical = insideVertical == null ? BorderDefinition.empty() : insideVertical;
        }

        static TableBorders empty() {
            return new TableBorders(BorderDefinition.empty(), BorderDefinition.empty(), BorderDefinition.empty());
        }

        BorderDefinition perimeter() {
            return perimeter;
        }

        BorderDefinition insideHorizontal() {
            return insideHorizontal;
        }

        BorderDefinition insideVertical() {
            return insideVertical;
        }
    }
    private static String normalizeFontName(String font) {
        if (font == null) {
            return "";
        }
        String trimmed = font.trim();
        if (trimmed.isEmpty()) {
            return "";
        }
        String lower = trimmed.toLowerCase(Locale.ROOT);
        if (Set.of("serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui").contains(lower)) {
            return lower;
        }
        if (trimmed.contains(" ") || trimmed.contains("-") || trimmed.contains("'")) {
            return "\"" + trimmed.replace("\"", "\"\"") + "\"";
        }
        return trimmed;
    }
}
