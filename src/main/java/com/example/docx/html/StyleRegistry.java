package com.example.docx.html;

import com.example.docx.model.document.WordDocument;

import java.util.LinkedHashMap;
import java.util.Map;

public final class StyleRegistry {
    private final Map<ParagraphCss, String> paragraphClasses = new LinkedHashMap<>();
    private final Map<RunCss, String> runClasses = new LinkedHashMap<>();
    private final Map<TableCss, String> tableClasses = new LinkedHashMap<>();
    private final Map<TableRowCss, String> rowClasses = new LinkedHashMap<>();
    private final Map<TableCellCss, String> cellClasses = new LinkedHashMap<>();
    private final ParagraphCss baseParagraph;
    private final RunCss baseRun;
    private int paragraphIndex = 1;
    private int runIndex = 1;
    private int tableIndex = 1;
    private int rowIndex = 1;
    private int cellIndex = 1;
    private static final int DEFAULT_PAGE_WIDTH_TWIPS = 11906;
    private static final int DEFAULT_PAGE_HEIGHT_TWIPS = 16838;
    private static final int DEFAULT_MARGIN_TWIPS = 1440;
    private static final double TWIP_TO_CM = 0.0017638889d;

    StyleRegistry(ParagraphCss baseParagraph, RunCss baseRun) {
        this.baseParagraph = baseParagraph == null ? ParagraphCss.empty() : baseParagraph;
        this.baseRun = baseRun == null ? RunCss.empty() : baseRun;
    }

    String registerParagraph(ParagraphCss css) {
        ParagraphCss key = css == null ? ParagraphCss.empty() : css;
        return paragraphClasses.computeIfAbsent(key, unused -> "p" + paragraphIndex++);
    }

    String registerRun(RunCss css) {
        RunCss key = css == null ? RunCss.empty() : css;
        return runClasses.computeIfAbsent(key, unused -> "s" + runIndex++);
    }

    String registerTable(TableCss css) {
        if (css == null || !css.hasColor()) {
            return null;
        }
        return tableClasses.computeIfAbsent(css, unused -> "t" + tableIndex++);
    }

    String registerRow(TableRowCss css) {
        if (css == null || !css.hasColor()) {
            return null;
        }
        return rowClasses.computeIfAbsent(css, unused -> "r" + rowIndex++);
    }

    String registerCell(TableCellCss css) {
        if (css == null || !css.hasColor()) {
            return null;
        }
        return cellClasses.computeIfAbsent(css, unused -> "c" + cellIndex++);
    }

    String buildCss(WordDocument document) {
        PageLayout layout = resolvePageLayout(document);
        StringBuilder builder = new StringBuilder();
        builder.append("body.docx-body{");
        builder.append("margin:0 auto;width:100%;max-width:").append(layout.pageWidth()).append(';');
        builder.append("min-height:").append(layout.pageHeight()).append(';');
        builder.append("padding:").append(layout.padding()).append(';');
        builder.append("box-sizing:border-box;");
        if (baseRun.color() != null) {
            builder.append("color:").append(baseRun.color()).append(';');
        } else {
            builder.append("color:#222;");
        }
        if (baseRun.fontFamily() != null && !baseRun.fontFamily().isBlank()) {
            builder.append("font-family:").append(baseRun.fontFamily()).append(';');
        } else {
            builder.append("font-family:\"Segoe UI\",-apple-system,BlinkMacSystemFont,\"Helvetica Neue\",Arial,sans-serif;");
        }
        if (baseRun.fontSize() != null) {
            builder.append("font-size:").append(baseRun.fontSize().css()).append(';');
        }
        if (baseParagraph.lineHeight() != null) {
            builder.append("line-height:").append(baseParagraph.lineHeight()).append(';');
        } else {
            builder.append("line-height:1.6;");
        }
        builder.append("background-color:#fff;}");
        builder.append('\n');

        builder.append(".docx-body .docx-paragraph{");
        builder.append("margin-top:").append(cssLengthOrDefault(baseParagraph.marginTop(), "0")).append(';');
        builder.append("margin-right:").append(cssLengthOrDefault(baseParagraph.marginRight(), "0")).append(';');
        builder.append("margin-bottom:").append(cssLengthOrDefault(baseParagraph.marginBottom(), "0")).append(';');
        builder.append("margin-left:").append(cssLengthOrDefault(baseParagraph.marginLeft(), "0")).append(';');
        if (baseParagraph.textIndent() != null) {
            builder.append("text-indent:").append(baseParagraph.textIndent().css()).append(';');
        }
        if (baseParagraph.fontSize() != null) {
            builder.append("font-size:").append(baseParagraph.fontSize().css()).append(';');
        }
        if (baseParagraph.fontFamily() != null && !baseParagraph.fontFamily().isBlank()) {
            builder.append("font-family:").append(baseParagraph.fontFamily()).append(';');
        }
        if (baseParagraph.backgroundColor() != null) {
            builder.append("background-color:").append(baseParagraph.backgroundColor()).append(';');
        }
        builder.append("}");
        builder.append('\n');

        builder.append(".docx-body .docx-empty{font-style:italic;color:#666;}");
        builder.append('\n');
        builder.append(".docx-body .docx-span{}");
        builder.append('\n');
        builder.append(".docx-body .docx-link{color:#0b57d0;text-decoration:underline;}");
        builder.append('\n');
        builder.append(".docx-body .docx-page-break{display:block;border:0;border-top:1px dashed #bbb;margin:2rem 0;}");
        builder.append('\n');
        builder.append(".docx-body .docx-column-break{display:block;border:0;border-top:1px dotted #bbb;margin:1.5rem 0;}");
        builder.append('\n');
        builder.append(".docx-body .docx-section-break{display:block;border:0;border-top:1px solid #ccc;margin:2rem 0;}");
        builder.append('\n');
        builder.append(".docx-body .docx-note-ref{font-size:0.75em;vertical-align:super;}");
        builder.append('\n');
        builder.append(".docx-body .docx-note-separator{display:block;border:0;border-top:1px solid #ccc;margin:1rem 0;}");
        builder.append('\n');
        builder.append(".docx-body .docx-table{border-collapse:collapse;width:100%;margin:1rem 0;}");
        builder.append('\n');
        builder.append(".docx-body .docx-table td,.docx-body .docx-table th{border:1px solid #bbb;padding:0.35rem 0.5rem;vertical-align:top;}");
        builder.append('\n');
        builder.append(".docx-body .docx-cell-middle{vertical-align:middle;}");
        builder.append('\n');
        builder.append(".docx-body .docx-cell-bottom{vertical-align:bottom;}");
        builder.append('\n');
        builder.append(".docx-body .docx-tab{display:inline-block;min-width:2em;}");
        builder.append('\n');
        builder.append(".docx-body .docx-drawing{display:inline-block;color:#555;font-style:italic;border:1px solid #ddd;padding:0.1rem 0.3rem;border-radius:0.2rem;background-color:#f9f9f9;}");
        builder.append('\n');
        builder.append(".docx-body .docx-field{background-color:rgba(0,0,0,0.05);padding:0 0.2rem;border-radius:0.2rem;}");
        builder.append('\n');
        builder.append(".docx-body .docx-sdt{border:1px dashed #bbb;padding:0.35rem;margin:0.5rem 0;}");
        builder.append('\n');
        builder.append(".docx-body .docx-sdt-inline{border:1px dashed #bbb;padding:0 0.25rem;margin:0 0.15rem;display:inline-block;}");
        builder.append('\n');

        for (Map.Entry<ParagraphCss, String> entry : paragraphClasses.entrySet()) {
            String declarations = entry.getKey().declarations();
            if (!declarations.isEmpty()) {
                builder.append(".docx-body .").append(entry.getValue()).append("{").append(declarations).append("}");
                builder.append('\n');
            }
        }
        for (Map.Entry<RunCss, String> entry : runClasses.entrySet()) {
            String declarations = entry.getKey().declarations();
            if (!declarations.isEmpty()) {
                builder.append(".docx-body .").append(entry.getValue()).append("{").append(declarations).append("}");
                builder.append('\n');
            }
        }
        for (Map.Entry<TableCss, String> entry : tableClasses.entrySet()) {
            String declarations = entry.getKey().declarations();
            if (!declarations.isEmpty()) {
                String className = entry.getValue();
                builder.append(".docx-body table.").append(className).append("{").append(declarations).append("}");
                builder.append('\n');
                builder.append(".docx-body table.").append(className).append(" td,.docx-body table.")
                        .append(className).append(" th{").append(declarations).append("}");
                builder.append('\n');
            }
        }
        for (Map.Entry<TableRowCss, String> entry : rowClasses.entrySet()) {
            String declarations = entry.getKey().declarations();
            if (!declarations.isEmpty()) {
                String className = entry.getValue();
                builder.append(".docx-body tr.").append(className).append("{").append(declarations).append("}");
                builder.append('\n');
                builder.append(".docx-body tr.").append(className).append(" > td,.docx-body tr.")
                        .append(className).append(" > th{").append(declarations).append("}");
                builder.append('\n');
            }
        }
        for (Map.Entry<TableCellCss, String> entry : cellClasses.entrySet()) {
            String declarations = entry.getKey().declarations();
            if (!declarations.isEmpty()) {
                builder.append(".docx-body td.").append(entry.getValue()).append("{").append(declarations).append("}");
                builder.append('\n');
            }
        }
        builder.append("@media screen{");
        builder.append("html{background-color:#b1b1b1;}");
        builder.append("body.docx-body{margin:1.5rem auto;border:1px solid #000;box-shadow:0 0 18px rgba(0,0,0,0.12);}");
        builder.append(".docx-body .docx-header::before{content:\"HEADER\";font-weight:bold;display:block;margin-bottom:0.5rem;}");
        builder.append(".docx-body .docx-header{border-left:1px dashed #000;border-right:1px dashed #000;border-bottom:1px dashed #000;padding:0.75rem 1rem;margin-bottom:1.5rem;}");
        builder.append(".docx-body .docx-footer::before{content:\"FOOTER\";font-weight:bold;display:block;margin-bottom:0.5rem;}");
        builder.append(".docx-body .docx-footer{border-left:1px dashed #000;border-right:1px dashed #000;border-top:1px dashed #000;padding:0.75rem 1rem;margin-top:1.5rem;}");
        builder.append('}');
        builder.append('\n');
        builder.append("@media print{body.docx-body{box-shadow:none;border:none;margin:0 auto;}html{background-color:#fff;}}");
        builder.append('\n');
        builder.append("@page{size:").append(layout.pageWidth()).append(' ').append(layout.pageHeight()).append(";margin:0;}");
        return builder.toString();
    }

    private static String cssLengthOrDefault(CssLength length, String fallback) {
        return length != null ? length.css() : fallback;
    }

    private PageLayout resolvePageLayout(WordDocument document) {
        int widthTwips = DEFAULT_PAGE_WIDTH_TWIPS;
        int heightTwips = DEFAULT_PAGE_HEIGHT_TWIPS;
        int topTwips = DEFAULT_MARGIN_TWIPS;
        int rightTwips = DEFAULT_MARGIN_TWIPS;
        int bottomTwips = DEFAULT_MARGIN_TWIPS;
        int leftTwips = DEFAULT_MARGIN_TWIPS;

        if (document != null) {
            WordDocument.SectionProperties section = document.bodySectionProperties().orElse(null);
            if (section != null) {
                WordDocument.PageDimensions dimensions = section.pageDimensions().orElse(null);
                if (dimensions != null) {
                    widthTwips = dimensions.widthTwips();
                    heightTwips = dimensions.heightTwips();
                }
                WordDocument.PageMargins margins = section.pageMargins().orElse(null);
                if (margins != null) {
                    topTwips = margins.top();
                    rightTwips = margins.right();
                    bottomTwips = margins.bottom();
                    leftTwips = margins.left();
                }
            }
        }

        return new PageLayout(
                twipsToCssCm(widthTwips),
                twipsToCssCm(heightTwips),
                twipsToCssCm(topTwips),
                twipsToCssCm(rightTwips),
                twipsToCssCm(bottomTwips),
                twipsToCssCm(leftTwips)
        );
    }

    private static String twipsToCssCm(int twips) {
        double cm = twips * TWIP_TO_CM;
        return DocxToHtml.formatDecimal(cm) + "cm";
    }

    private record PageLayout(String pageWidth,
                              String pageHeight,
                              String paddingTop,
                              String paddingRight,
                              String paddingBottom,
                              String paddingLeft) {
        String padding() {
            return paddingTop + ' ' + paddingRight + ' ' + paddingBottom + ' ' + paddingLeft;
        }
    }
}
