package com.example.docx.html;

import com.example.docx.model.document.WordDocument;

import java.util.ArrayList;
import java.util.List;

public record ParagraphCss(String textAlign,
                    CssLength marginTop,
                    CssLength marginBottom,
                    CssLength marginLeft,
                    CssLength marginRight,
                    CssLength textIndent,
                    String lineHeight,
                    CssLength fontSize,
                    String fontFamily,
                    String backgroundColor,
                    boolean keepTogether,
                    boolean keepWithNext,
                    boolean pageBreakBefore) {
    public static ParagraphCss empty() {
        return new ParagraphCss(null, null, null, null, null, null, null, null, null, null, false, false, false);
    }

    public static ParagraphCss from(StyleResolver.ResolvedParagraph resolved, StyleResolver.ResolvedRun defaultRun) {
        if (resolved == null) {
            return ParagraphCss.empty();
        }
        WordDocument.Spacing spacing = resolved.spacing();
        CssLength top = spacing != null ? DocxToHtml.toCssLength(spacing.before().orElse(null)) : null;
        CssLength bottom = spacing != null ? DocxToHtml.toCssLength(spacing.after().orElse(null)) : null;
        WordDocument.Indentation indentation = resolved.indentation();
        CssLength left = indentation != null ? DocxToHtml.toCssLength(indentation.left().orElse(null)) : null;
        CssLength right = indentation != null ? DocxToHtml.toCssLength(indentation.right().orElse(null)) : null;
        CssLength indent = DocxToHtml.computeTextIndent(indentation);
        String lineHeight = DocxToHtml.computeLineHeight(spacing);
        CssLength paragraphFontSize = defaultRun.fontSizeHalfPoints()
                .map(size -> CssLength.points(size / 2.0d))
                .orElse(null);
        String paragraphFontFamily = DocxToHtml.fontStackToCss(defaultRun.fontStack());
        if (paragraphFontFamily != null && paragraphFontFamily.isBlank()) {
            paragraphFontFamily = null;
        }
        String background = resolved.shadingColor();
        return new ParagraphCss(DocxToHtml.alignmentToCss(resolved.alignment()), top, bottom, left, right, indent,
                lineHeight, paragraphFontSize, paragraphFontFamily, background, resolved.keepTogether(),
                resolved.keepWithNext(), resolved.pageBreakBefore());
    }

    String declarations() {
        List<String> items = new ArrayList<>();
        if (textAlign != null) {
            items.add("text-align:" + textAlign);
        }
        if (marginTop != null) {
            items.add("margin-top:" + marginTop.css());
        }
        if (marginBottom != null) {
            items.add("margin-bottom:" + marginBottom.css());
        }
        if (marginLeft != null) {
            items.add("margin-left:" + marginLeft.css());
        }
        if (marginRight != null) {
            items.add("margin-right:" + marginRight.css());
        }
        if (textIndent != null) {
            items.add("text-indent:" + textIndent.css());
        }
        if (lineHeight != null) {
            items.add("line-height:" + lineHeight);
        }
        if (fontSize != null) {
            items.add("font-size:" + fontSize.css());
        }
        if (fontFamily != null && !fontFamily.isBlank()) {
            items.add("font-family:" + fontFamily);
        }
        if (backgroundColor != null) {
            items.add("background-color:" + backgroundColor);
        }
        if (keepTogether) {
            items.add("page-break-inside:avoid");
            items.add("break-inside:avoid");
        }
        if (keepWithNext) {
            items.add("page-break-after:avoid");
            items.add("break-after:avoid");
        }
        if (pageBreakBefore) {
            items.add("page-break-before:always");
            items.add("break-before:page");
        }
        return String.join(";", items);
    }
}
