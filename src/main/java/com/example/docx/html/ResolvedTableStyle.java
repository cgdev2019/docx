package com.example.docx.html;

import com.example.docx.model.document.WordDocument;
import com.example.docx.parser.Namespaces;
import com.example.docx.parser.XmlUtils;
import org.w3c.dom.Element;

import java.util.Map;

final class ResolvedTableStyle {
    private final String tableBackground;
    private final Map<StyleResolver.TableStyleRegion, StyleResolver.RegionStyle> regions;
    private final BorderDefinition tableBorders;
    private final BorderDefinition insideHorizontal;
    private final BorderDefinition insideVertical;

    ResolvedTableStyle(String tableBackground,
                       Map<StyleResolver.TableStyleRegion, StyleResolver.RegionStyle> regions,
                       BorderDefinition tableBorders,
                       BorderDefinition insideHorizontal,
                       BorderDefinition insideVertical) {
        this.tableBackground = tableBackground;
        this.regions = regions.isEmpty() ? Map.of() : Map.copyOf(regions);
        this.tableBorders = tableBorders == null ? BorderDefinition.empty() : tableBorders;
        this.insideHorizontal = insideHorizontal == null ? BorderDefinition.empty() : insideHorizontal;
        this.insideVertical = insideVertical == null ? BorderDefinition.empty() : insideVertical;
    }

    static ResolvedTableStyle empty() {
        return new ResolvedTableStyle(null, Map.of(), BorderDefinition.empty(), BorderDefinition.empty(), BorderDefinition.empty());
    }

    String tableBackground() {
        return tableBackground;
    }

    BorderDefinition tableBorders() {
        return tableBorders;
    }

    BorderDefinition insideHorizontal() {
        return insideHorizontal;
    }

    BorderDefinition insideVertical() {
        return insideVertical;
    }

    StyleResolver.RegionStyle rowRegion(WordDocument.TableRowProperties properties, int rowIndex, int rowCount) {
        StyleResolver.RegionStyle conditional = regionFromConditionalFormatting(properties);
        if (conditional != null) {
            return conditional;
        }
        if (rowIndex == 0) {
            StyleResolver.RegionStyle first = regions.get(StyleResolver.TableStyleRegion.FIRST_ROW);
            if (first != null) {
                return first;
            }
        }
        if (rowCount > 0 && rowIndex == rowCount - 1) {
            StyleResolver.RegionStyle last = regions.get(StyleResolver.TableStyleRegion.LAST_ROW);
            if (last != null) {
                return last;
            }
        }
        return regions.get(StyleResolver.TableStyleRegion.WHOLE_TABLE);
    }

    private StyleResolver.RegionStyle regionFromConditionalFormatting(WordDocument.TableRowProperties properties) {
        if (properties == null) {
            return null;
        }
        Element raw = properties.rawProperties().orElse(null);
        if (raw == null) {
            return null;
        }
        Element cnfStyle = XmlUtils.firstChild(raw, Namespaces.WORD_MAIN, "cnfStyle").orElse(null);
        if (cnfStyle == null) {
            return null;
        }
        String value = firstNonBlank(XmlUtils.attribute(cnfStyle, "w:val"), XmlUtils.attribute(cnfStyle, "val"));
        if (value == null || value.isBlank()) {
            return null;
        }
        String flags = normalizeFlags(value.trim());
        StyleResolver.RegionStyle region = selectRegion(flags, 0, StyleResolver.TableStyleRegion.FIRST_ROW);
        if (region != null) {
            return region;
        }
        region = selectRegion(flags, 1, StyleResolver.TableStyleRegion.LAST_ROW);
        if (region != null) {
            return region;
        }
        region = selectRegion(flags, 6, StyleResolver.TableStyleRegion.BAND1_HORZ);
        if (region != null) {
            return region;
        }
        region = selectRegion(flags, 7, StyleResolver.TableStyleRegion.BAND2_HORZ);
        if (region != null) {
            return region;
        }
        region = selectRegion(flags, 4, StyleResolver.TableStyleRegion.BAND1_VERT);
        if (region != null) {
            return region;
        }
        region = selectRegion(flags, 5, StyleResolver.TableStyleRegion.BAND2_VERT);
        if (region != null) {
            return region;
        }
        region = selectRegion(flags, 2, StyleResolver.TableStyleRegion.FIRST_COLUMN);
        if (region != null) {
            return region;
        }
        return selectRegion(flags, 3, StyleResolver.TableStyleRegion.LAST_COLUMN);
    }

    private StyleResolver.RegionStyle selectRegion(String flags, int index, StyleResolver.TableStyleRegion region) {
        if (index >= flags.length() || flags.charAt(index) != '1') {
            return null;
        }
        return regions.get(region);
    }

    private static String normalizeFlags(String value) {
        if (value.length() >= 12) {
            return value;
        }
        StringBuilder builder = new StringBuilder(12);
        for (int i = 0; i < 12 - value.length(); i++) {
            builder.append('0');
        }
        builder.append(value);
        return builder.toString();
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
}
