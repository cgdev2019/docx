package com.example.docx.html;

import java.util.Map;

final class ResolvedTableStyle {
    private final String tableBackground;
    private final Map<StyleResolver.TableStyleRegion, StyleResolver.RegionStyle> regions;

    ResolvedTableStyle(String tableBackground, Map<StyleResolver.TableStyleRegion, StyleResolver.RegionStyle> regions) {
        this.tableBackground = tableBackground;
        this.regions = regions.isEmpty() ? Map.of() : Map.copyOf(regions);
    }

    static ResolvedTableStyle empty() {
        return new ResolvedTableStyle(null, Map.of());
    }

    String tableBackground() {
        return tableBackground;
    }

    StyleResolver.RegionStyle rowRegion(int rowIndex, int rowCount) {
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
}
