package com.example.docx.html;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

public record RunCss(boolean bold,
              boolean italic,
              boolean underline,
              Set<String> decorationLines,
              String decorationStyle,
              boolean smallCaps,
              boolean allCaps,
              String color,
              String backgroundColor,
              CssLength fontSize,
              String fontFamily,
              String verticalAlign) {

    static RunCss empty() {
        return new RunCss(false, false, false, Set.of(), null, false, false, null, null, null, null, null);
    }

    static RunCss from(StyleResolver.ResolvedRun resolved) {
        String color = resolved.color().flatMap(DocxToHtml::normalizeColor).orElse(null);
        String background = resolved.highlight().flatMap(DocxToHtml::normalizeHighlight).orElse(null);
        CssLength fontSize = resolved.fontSizeHalfPoints()
                .map(size -> CssLength.points(size / 2.0d))
                .orElse(null);
        String fontFamily = DocxToHtml.fontStackToCss(resolved.fontStack());
        String verticalAlign = resolved.verticalAlignment().map(DocxToHtml::normalizeVerticalAlignment).orElse(null);
        String underlineType = resolved.underlineType();
        boolean underline = resolved.underline() && (underlineType == null || !"none".equalsIgnoreCase(underlineType));
        String decorationStyle = DocxToHtml.decorationStyle(underlineType, resolved.doubleStrike());
        Set<String> lines = new LinkedHashSet<>();
        if (underline) {
            lines.add("underline");
        }
        if (resolved.strike() || resolved.doubleStrike()) {
            lines.add("line-through");
        }
        return new RunCss(resolved.bold(), resolved.italic(), underline, Set.copyOf(lines),
                decorationStyle, resolved.smallCaps(), resolved.allCaps(), color,
                background, fontSize, fontFamily, verticalAlign);
    }

    String declarations() {
        List<String> items = new ArrayList<>();
        if (bold) {
            items.add("font-weight:bold");
        }
        if (italic) {
            items.add("font-style:italic");
        }
        if (fontSize != null) {
            items.add("font-size:" + fontSize.css());
        }
        if (fontFamily != null && !fontFamily.isBlank()) {
            items.add("font-family:" + fontFamily);
        }
        if (color != null) {
            items.add("color:" + color);
        }
        if (backgroundColor != null) {
            items.add("background-color:" + backgroundColor);
        }
        if (smallCaps) {
            items.add("font-variant:small-caps");
        }
        if (allCaps) {
            items.add("text-transform:uppercase");
        }
        if (!decorationLines.isEmpty()) {
            items.add("text-decoration-line:" + String.join(" ", decorationLines));
        }
        if (decorationStyle != null) {
            items.add("text-decoration-style:" + decorationStyle);
        }
        if (verticalAlign != null) {
            items.add("vertical-align:" + verticalAlign);
        }
        return String.join(";", items);
    }
}
