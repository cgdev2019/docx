package com.example.docx.html;

import com.example.docx.model.DocxPackage;
import com.example.docx.model.document.WordDocument;
import com.example.docx.model.relationship.RelationshipSet;
import com.example.docx.model.styles.StyleDefinitions;
import com.example.docx.model.support.Theme;

import com.example.docx.parser.Namespaces;
import com.example.docx.parser.XmlUtils;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

public final class DocxToHtml {

    public static final String DOCUMENT_PART = "word/document.xml";
    public static final WordDocument.RunProperties EMPTY_RUN_PROPERTIES =
            new WordDocument.RunProperties(null, false, false, false, null,
                    false, false, false, false, false, null, null, null,
                    null, null, Map.of(), null);
    public static final WordDocument.ParagraphProperties EMPTY_PARAGRAPH_PROPERTIES =
            new WordDocument.ParagraphProperties(null, null, null, null, null,
                    null, false, false, false, List.of(), null);

    private static final Map<String, String> HIGHLIGHT_COLORS = createHighlightColorMap();

    private final String language;

    public DocxToHtml() {
        this("fr");
    }

    public DocxToHtml(String language) {
        this.language = language == null || language.isBlank() ? "fr" : language;
    }

    public String convert(DocxPackage docxPackage) {
        Objects.requireNonNull(docxPackage, "docxPackage");
        WordDocument document = docxPackage.document().orElse(null);
        StyleDefinitions styles = docxPackage.styles().orElse(StyleDefinitions.empty());
        Map<String, String> themeColors = extractThemeColors(docxPackage);
        StyleResolver styleResolver = new StyleResolver(styles, themeColors);
        RelationshipSet relationships = docxPackage.relationshipsByPart().get(DOCUMENT_PART);
        HyperlinkResolver hyperlinkResolver = new HyperlinkResolver(relationships);
        StyleResolver.ResolvedParagraph baseParagraph = styleResolver.resolveParagraph(null, List.of());
        StyleResolver.ResolvedRun baseRun = styleResolver.resolveRun(EMPTY_RUN_PROPERTIES, baseParagraph);
        StyleRegistry registry = new StyleRegistry(ParagraphCss.from(baseParagraph, baseRun), RunCss.from(baseRun));

        StringBuilder bodyBuilder = new StringBuilder();
        if (document == null || document.bodyElements().isEmpty()) {
            bodyBuilder.append("<p class=\"docx-paragraph docx-empty\">Document vide</p>");
        } else {
            for (WordDocument.Block block : document.bodyElements()) {
                String blockHtml = convertBlock(block, styleResolver, hyperlinkResolver, registry, themeColors);
                if (!blockHtml.isEmpty()) {
                    bodyBuilder.append(blockHtml).append("\n");
                }
            }
        }

        return buildHtml(document, registry, bodyBuilder.toString());
    }

    private static Map<String, String> extractThemeColors(DocxPackage docxPackage) {
        if (docxPackage == null) {
            return Map.of();
        }
        return docxPackage.theme()
                .flatMap(Theme::document)
                .map(DocxToHtml::readThemeColorScheme)
                .orElse(Map.of());
    }

    private static Map<String, String> readThemeColorScheme(Document document) {
        if (document == null) {
            return Map.of();
        }
        Element root = document.getDocumentElement();
        if (root == null) {
            return Map.of();
        }
        Element themeElements = XmlUtils.firstChild(root, Namespaces.DRAWINGML_MAIN, "themeElements").orElse(null);
        if (themeElements == null) {
            return Map.of();
        }
        Element clrScheme = XmlUtils.firstChild(themeElements, Namespaces.DRAWINGML_MAIN, "clrScheme").orElse(null);
        if (clrScheme == null) {
            return Map.of();
        }
        Map<String, String> colors = new LinkedHashMap<>();
        for (Element entry : XmlUtils.childElements(clrScheme)) {
            if (entry == null || entry.getLocalName() == null) {
                continue;
            }
            if (!Namespaces.DRAWINGML_MAIN.equals(entry.getNamespaceURI())) {
                continue;
            }
            String key = entry.getLocalName().toLowerCase(Locale.ROOT);
            String value = readSchemeColor(entry);
            if (value != null) {
                colors.put(key, value);
            }
        }
        return Map.copyOf(colors);
    }

    private static String readSchemeColor(Element entry) {
        for (Element child : XmlUtils.childElements(entry)) {
            if (child == null || !Namespaces.DRAWINGML_MAIN.equals(child.getNamespaceURI())) {
                continue;
            }
            String color = colorFromAttributes(child);
            if (color != null) {
                return color;
            }
        }
        return null;
    }

    private static String colorFromAttributes(Element element) {
        if (element == null) {
            return null;
        }
        String last = XmlUtils.attribute(element, "lastClr");
        if (last != null) {
            String normalized = normalizeDirectShadingFill(last);
            if (normalized != null) {
                return normalized;
            }
        }
        String val = XmlUtils.attribute(element, "val");
        if (val != null) {
            return normalizeDirectShadingFill(val);
        }
        return null;
    }

    private String convertBlock(WordDocument.Block block,
                                StyleResolver styleResolver,
                                HyperlinkResolver hyperlinkResolver,
                                StyleRegistry registry,
                                Map<String, String> themeColors) {
        return convertBlock(block, styleResolver, hyperlinkResolver, registry, themeColors, List.of());
    }

    private String convertBlock(WordDocument.Block block,
                                StyleResolver styleResolver,
                                HyperlinkResolver hyperlinkResolver,
                                StyleRegistry registry,
                                Map<String, String> themeColors,
                                List<WordDocument.RunProperties> extraRunFallbacks) {
        if (block instanceof WordDocument.Paragraph paragraph) {
            return convertParagraph(paragraph, styleResolver, hyperlinkResolver, registry, themeColors, extraRunFallbacks);
        }
        if (block instanceof WordDocument.Table table) {
            return convertTable(table, styleResolver, hyperlinkResolver, registry, themeColors, extraRunFallbacks);
        }
        if (block instanceof WordDocument.StructuredDocumentTag sdt) {
            return convertStructuredDocumentTag(sdt, styleResolver, hyperlinkResolver, registry, themeColors, extraRunFallbacks);
        }
        if (block instanceof WordDocument.SectionBreak) {
            return "<span class=\"docx-section-break\"></span>";
        }
        if (block instanceof WordDocument.Bookmark bookmark) {
            return bookmark.kind() == WordDocument.Bookmark.Kind.START
                    ? renderBookmarkAnchor(bookmark)
                    : "";
        }
        return "";
    }

    private String convertParagraph(WordDocument.Paragraph paragraph,
                                    StyleResolver styleResolver,
                                    HyperlinkResolver hyperlinkResolver,
                                    StyleRegistry registry,
                                    Map<String, String> themeColors) {
        return convertParagraph(paragraph, styleResolver, hyperlinkResolver, registry, themeColors, List.of());
    }

    private String convertParagraph(WordDocument.Paragraph paragraph,
                                    StyleResolver styleResolver,
                                    HyperlinkResolver hyperlinkResolver,
                                    StyleRegistry registry,
                                    Map<String, String> themeColors,
                                    List<WordDocument.RunProperties> extraRunFallbacks) {
        StyleResolver.ResolvedParagraph resolved = styleResolver.resolveParagraph(paragraph.properties(), extraRunFallbacks);
        StyleResolver.ResolvedRun defaultRun = styleResolver.resolveRun(EMPTY_RUN_PROPERTIES, resolved);
        ParagraphCss css = ParagraphCss.from(resolved, defaultRun);
        String paragraphClass = registry.registerParagraph(css);
        StringBuilder inner = new StringBuilder();
        for (WordDocument.ParagraphContent content : paragraph.content()) {
            inner.append(convertParagraphContent(content, styleResolver, hyperlinkResolver, registry, resolved, themeColors));
        }
        if (inner.length() == 0) {
            inner.append("&nbsp;");
        }
        return "<p class=\"docx-paragraph " + paragraphClass + "\">" + inner + "</p>";
    }
    private String convertParagraphContent(WordDocument.ParagraphContent content,
                                           StyleResolver styleResolver,
                                           HyperlinkResolver hyperlinkResolver,
                                           StyleRegistry registry,
                                           StyleResolver.ResolvedParagraph paragraph,
                                           Map<String, String> themeColors) {
        if (content instanceof WordDocument.Run run) {
            return convertRun(run, paragraph, styleResolver, registry);
        }
        if (content instanceof WordDocument.Hyperlink hyperlink) {
            return convertHyperlink(hyperlink, paragraph, styleResolver, hyperlinkResolver, registry);
        }
        if (content instanceof WordDocument.BookmarkStart bookmarkStart) {
            return renderBookmarkStart(bookmarkStart);
        }
        if (content instanceof WordDocument.BookmarkEnd) {
            return "";
        }
        if (content instanceof WordDocument.Field field) {
            return convertField(field, paragraph, styleResolver, hyperlinkResolver, registry);
        }
        if (content instanceof WordDocument.StructuredDocumentTagRun sdt) {
            return convertStructuredDocumentTagRun(sdt, paragraph, styleResolver, hyperlinkResolver, registry, themeColors);
        }
        return "";
    }
    private String convertHyperlink(WordDocument.Hyperlink hyperlink,
                                    StyleResolver.ResolvedParagraph paragraph,
                                    StyleResolver styleResolver,
                                    HyperlinkResolver hyperlinkResolver,
                                    StyleRegistry registry) {
        StringBuilder content = new StringBuilder();
        for (WordDocument.Run run : hyperlink.runs()) {
            content.append(convertRun(run, paragraph, styleResolver, registry));
        }
        if (content.length() == 0) {
            return "";
        }
        String href = hyperlinkResolver.resolve(
                        hyperlink.relationshipId().orElse(null),
                        hyperlink.anchor().orElse(null))
                .map(DocxToHtml::escapeHtmlAttribute)
                .orElse("#");
        return "<a class=\"docx-link\" href=\"" + href + "\">" + content + "</a>";
    }

    private String convertField(WordDocument.Field field,
                                StyleResolver.ResolvedParagraph paragraph,
                                StyleResolver styleResolver,
                                HyperlinkResolver hyperlinkResolver,
                                StyleRegistry registry) {
        StringBuilder result = new StringBuilder();
        for (WordDocument.Run run : field.resultRuns()) {
            result.append(convertRun(run, paragraph, styleResolver, registry));
        }
        if (result.length() == 0) {
            for (WordDocument.Run run : field.instructionRuns()) {
                result.append(convertRun(run, paragraph, styleResolver, registry));
            }
        }
        if (result.length() == 0) {
            return "";
        }
        return "<span class=\"docx-field\">" + result + "</span>";
    }

    private String convertStructuredDocumentTagRun(WordDocument.StructuredDocumentTagRun sdt,
                                                   StyleResolver.ResolvedParagraph paragraph,
                                                   StyleResolver styleResolver,
                                                   HyperlinkResolver hyperlinkResolver,
                                                   StyleRegistry registry,
                                                   Map<String, String> themeColors) {
        StringBuilder builder = new StringBuilder();
        builder.append("<span class=\"docx-sdt-inline\"");
        sdt.properties().tag().ifPresent(tag -> builder.append(" data-tag=\"").append(escapeHtmlAttribute(tag)).append("\""));
        sdt.properties().alias().ifPresent(alias -> builder.append(" data-alias=\"").append(escapeHtmlAttribute(alias)).append("\""));
        sdt.properties().id().ifPresent(id -> builder.append(" data-id=\"").append(escapeHtmlAttribute(id)).append("\""));
        builder.append(">");
        for (WordDocument.ParagraphContent child : sdt.content()) {
            builder.append(convertParagraphContent(child, styleResolver, hyperlinkResolver, registry, paragraph, themeColors));
        }
        builder.append("</span>");
        return builder.toString();
    }

    private String convertStructuredDocumentTag(WordDocument.StructuredDocumentTag sdt,
                                                StyleResolver styleResolver,
                                                HyperlinkResolver hyperlinkResolver,
                                                StyleRegistry registry,
                                                Map<String, String> themeColors) {
        return convertStructuredDocumentTag(sdt, styleResolver, hyperlinkResolver, registry, themeColors, List.of());
    }

    private String convertStructuredDocumentTag(WordDocument.StructuredDocumentTag sdt,
                                                StyleResolver styleResolver,
                                                HyperlinkResolver hyperlinkResolver,
                                                StyleRegistry registry,
                                                Map<String, String> themeColors,
                                                List<WordDocument.RunProperties> extraRunFallbacks) {
        StringBuilder builder = new StringBuilder();
        builder.append("<section class=\"docx-sdt\"");
        sdt.properties().tag().ifPresent(tag -> builder.append(" data-tag=\"").append(escapeHtmlAttribute(tag)).append("\""));
        sdt.properties().alias().ifPresent(alias -> builder.append(" data-alias=\"").append(escapeHtmlAttribute(alias)).append("\""));
        sdt.properties().id().ifPresent(id -> builder.append(" data-id=\"").append(escapeHtmlAttribute(id)).append("\""));
        builder.append(">");
        for (WordDocument.Block child : sdt.content()) {
            String blockContent = convertBlock(child, styleResolver, hyperlinkResolver, registry, themeColors, extraRunFallbacks);
            if (!blockContent.isEmpty()) {
                builder.append(blockContent);
            }
        }
        builder.append("</section>");
        return builder.toString();
    }

    private String convertTable(WordDocument.Table table,
                                StyleResolver styleResolver,
                                HyperlinkResolver hyperlinkResolver,
                                StyleRegistry registry,
                                Map<String, String> themeColors) {
        return convertTable(table, styleResolver, hyperlinkResolver, registry, themeColors, List.of());
    }

    private String convertTable(WordDocument.Table table,
                                StyleResolver styleResolver,
                                HyperlinkResolver hyperlinkResolver,
                                StyleRegistry registry,
                                Map<String, String> themeColors,
                                List<WordDocument.RunProperties> inheritedRunFallbacks) {
        ResolvedTableStyle tableStyle = styleResolver.resolveTableStyle(table.properties());
        int rowCount = table.rows().size();
        String tableBackground = tableShadingColor(table.properties(), themeColors);
        if (tableBackground == null) {
            tableBackground = tableStyle.tableBackground();
        }
        List<String> tableClasses = new ArrayList<>();
        tableClasses.add("docx-table");
        if (tableBackground != null) {
            String tableClass = registry.registerTable(new TableCss(tableBackground));
            if (tableClass != null) {
                tableClasses.add(tableClass);
            }
        }
        StringBuilder builder = new StringBuilder();
        builder.append("<table class=\"").append(String.join(" ", tableClasses)).append("\">");
        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            WordDocument.TableRow row = table.rows().get(rowIndex);
            StyleResolver.RegionStyle region = tableStyle.rowRegion(rowIndex, rowCount);
            String rowBackground = tableRowShadingColor(row.properties(), themeColors);
            if (rowBackground == null && region != null) {
                rowBackground = region.backgroundColor();
            }
            List<WordDocument.RunProperties> rowFallbacks = new ArrayList<>(inheritedRunFallbacks);
            if (region != null && region.runProperties() != null) {
                rowFallbacks.add(region.runProperties());
            }
            List<String> rowClasses = new ArrayList<>();
            if (rowBackground != null) {
                String rowClass = registry.registerRow(new TableRowCss(rowBackground));
                if (rowClass != null) {
                    rowClasses.add("docx-row");
                    rowClasses.add(rowClass);
                }
            }
            builder.append("<tr");
            if (!rowClasses.isEmpty()) {
                builder.append(" class=\"").append(String.join(" ", rowClasses)).append("\"");
            }
            builder.append(">");
            List<WordDocument.TableCell> cells = row.cells();
            for (WordDocument.TableCell cell : cells) {
                builder.append(convertTableCell(cell, rowBackground, tableBackground, styleResolver, hyperlinkResolver, registry, themeColors, rowFallbacks));
            }
            builder.append("</tr>");
        }
        builder.append("</table>");
        return builder.toString();
    }

    private String convertTableCell(WordDocument.TableCell cell,
                                    String rowBackground,
                                    String tableBackground,
                                    StyleResolver styleResolver,
                                    HyperlinkResolver hyperlinkResolver,
                                    StyleRegistry registry,
                                    Map<String, String> themeColors,
                                    List<WordDocument.RunProperties> rowRunFallbacks) {
        List<WordDocument.RunProperties> effectiveRunFallbacks = (rowRunFallbacks == null || rowRunFallbacks.isEmpty())
                ? List.of()
                : List.copyOf(rowRunFallbacks);
        StringBuilder content = new StringBuilder();
        for (WordDocument.Block block : cell.content()) {
            String fragment = convertBlock(block, styleResolver, hyperlinkResolver, registry, themeColors, effectiveRunFallbacks);
            if (!fragment.isEmpty()) {
                content.append(fragment);
            }
        }
        if (content.length() == 0) {
            content.append("&nbsp;");
        }
        List<String> classes = new ArrayList<>();
        classes.add("docx-cell");
        cell.properties().verticalAlignment().ifPresent(val -> {
            switch (val) {
                case "center" -> classes.add("docx-cell-middle");
                case "bottom" -> classes.add("docx-cell-bottom");
                default -> {
                }
            }
        });
        String background = tableCellShadingColor(cell.properties(), themeColors);
        if (background != null
                && !Objects.equals(background, rowBackground)
                && !Objects.equals(background, tableBackground)) {
            String cellClass = registry.registerCell(new TableCellCss(background));
            if (cellClass != null) {
                classes.add(cellClass);
            }
        }
        StringBuilder cellBuilder = new StringBuilder();
        cellBuilder.append("<td");
        if (!classes.isEmpty()) {
            cellBuilder.append(" class=\"").append(String.join(" ", classes)).append("\"");
        }
        cell.properties().gridSpan().ifPresent(span -> {
            if (span != null && span > 1) {
                cellBuilder.append(" colspan=\"").append(span).append("\"");
            }
        });
        cellBuilder.append(">");
        cellBuilder.append(content);
        cellBuilder.append("</td>");
        return cellBuilder.toString();
    }

    private String convertRun(WordDocument.Run run,
                              StyleResolver.ResolvedParagraph paragraph,
                              StyleResolver styleResolver,
                              StyleRegistry registry) {
        StyleResolver.ResolvedRun resolvedRun = styleResolver.resolveRun(run.properties(), paragraph);
        if (resolvedRun.vanish()) {
            return "";
        }
        StringBuilder content = new StringBuilder();
        for (WordDocument.Inline inline : run.elements()) {
            String fragment = convertInline(inline);
            if (!fragment.isEmpty()) {
                content.append(fragment);
            }
        }
        if (content.length() == 0) {
            return "";
        }
        RunCss css = RunCss.from(resolvedRun);
        String declarations = css.declarations();
        if (declarations.isEmpty()) {
            return content.toString();
        }
        String className = registry.registerRun(css);
        return "<span class=\"docx-span " + className + "\">" + content + "</span>";
    }

    private String convertInline(WordDocument.Inline inline) {
        if (inline instanceof WordDocument.Text text) {
            return renderText(text);
        }
        if (inline instanceof WordDocument.Break br) {
            return renderBreak(br);
        }
        if (inline instanceof WordDocument.Tab) {
            return "<span class=\"docx-tab\">&emsp;</span>";
        }
        if (inline instanceof WordDocument.Drawing drawing) {
            return renderDrawing(drawing);
        }
        if (inline instanceof WordDocument.FootnoteReference footnoteReference) {
            return "<sup class=\"docx-note-ref\" data-note-type=\"footnote\">" + footnoteReference.id() + "</sup>";
        }
        if (inline instanceof WordDocument.EndnoteReference endnoteReference) {
            return "<sup class=\"docx-note-ref\" data-note-type=\"endnote\">" + endnoteReference.id() + "</sup>";
        }
        if (inline instanceof WordDocument.CommentReference commentReference) {
            return "<sup class=\"docx-note-ref\" data-note-type=\"comment\">" + commentReference.id() + "</sup>";
        }
        if (inline instanceof WordDocument.FieldInstruction) {
            return "";
        }
        if (inline instanceof WordDocument.FieldCharacter) {
            return "";
        }
        if (inline instanceof WordDocument.Symbol symbol) {
            return renderSymbol(symbol);
        }
        if (inline instanceof WordDocument.SoftHyphen) {
            return "&shy;";
        }
        if (inline instanceof WordDocument.NoBreakHyphen) {
            return "&#8209;";
        }
        if (inline instanceof WordDocument.Separator separator) {
            return "<span class=\"docx-note-separator\" data-kind=\"" + separator.kind().name().toLowerCase(Locale.ROOT) + "\"></span>";
        }
        if (inline instanceof WordDocument.ReferenceMark) {
            return "";
        }
        return "";
    }

    private String renderBreak(WordDocument.Break br) {
        return switch (br.type()) {
            case PAGE -> "<span class=\"docx-page-break\"></span>";
            case COLUMN -> "<span class=\"docx-column-break\"></span>";
            default -> "<br/>";
        };
    }

    private String renderDrawing(WordDocument.Drawing drawing) {
        StringBuilder builder = new StringBuilder("<span class=\"docx-drawing\"");
        drawing.relationshipId().ifPresent(rel -> builder.append(" data-rel=\"").append(escapeHtmlAttribute(rel)).append("\""));
        if (drawing.width() > 0 && drawing.height() > 0) {
            builder.append(" data-size=\"").append(drawing.width()).append("x").append(drawing.height()).append("\"");
        }
        builder.append(">");
        builder.append("[Image");
        drawing.description().ifPresent(desc -> {
            if (!desc.isBlank()) {
                builder.append(": ").append(escapeHtml(desc));
            }
        });
        builder.append("]");
        builder.append("</span>");
        return builder.toString();
    }

    private String renderSymbol(WordDocument.Symbol symbol) {
        String code = symbol.charCode();
        int codePoint = -1;
        if (code != null) {
            try {
                codePoint = Integer.parseInt(code, 16);
            } catch (NumberFormatException e) {
                try {
                    codePoint = Integer.parseInt(code);
                } catch (NumberFormatException ignored) {
                    codePoint = -1;
                }
            }
        }
        if (codePoint == -1) {
            return escapeHtml(code == null ? "" : code);
        }
        return escapeHtml(new String(Character.toChars(codePoint)));
    }

    private String renderText(WordDocument.Text text) {
        String value = text.text();
        if (value.isEmpty()) {
            return "";
        }
        StringBuilder builder = new StringBuilder();
        boolean preserve = text.preserveSpace();
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            switch (c) {
                case ' ' -> {
                    if (preserve && (i == 0 || i == value.length() - 1 || value.charAt(i - 1) == ' ' || (i + 1 < value.length() && value.charAt(i + 1) == ' '))) {
                        builder.append("&nbsp;");
                    } else {
                        builder.append(' ');
                    }
                }
                case '\t' -> builder.append("&nbsp;&nbsp;&nbsp;&nbsp;");
                case '\r' -> {
                    if (i + 1 < value.length() && value.charAt(i + 1) == '\n') {
                        // skip, newline will be handled by the next iteration
                    } else {
                        builder.append("<br/>");
                    }
                }
                case '\n' -> builder.append("<br/>");
                default -> builder.append(escapeChar(c));
            }
        }
        return builder.toString();
    }
    private String renderBookmarkAnchor(WordDocument.Bookmark bookmark) {
        String name = bookmark.name().orElse(bookmark.id().orElse(null));
        if (name == null || name.isBlank()) {
            return "";
        }
        return "<a id=\"" + escapeHtmlAttribute(name) + "\"></a>";
    }

    private String renderBookmarkStart(WordDocument.BookmarkStart bookmarkStart) {
        String name = bookmarkStart.name().orElse(bookmarkStart.id().orElse(null));
        if (name == null || name.isBlank()) {
            return "";
        }
        return "<a id=\"" + escapeHtmlAttribute(name) + "\"></a>";
    }

    private String buildHtml(WordDocument document, StyleRegistry registry, String bodyContent) {
        StringBuilder builder = new StringBuilder();
        builder.append("<!DOCTYPE html>\n");
        builder.append("<html lang=\"").append(escapeHtmlAttribute(language)).append("\">\n");
        builder.append("<head>\n<meta charset=\"utf-8\">\n<style>\n");
        builder.append(registry.buildCss(document));
        builder.append("</style>\n</head>\n<body class=\"docx-body\">\n");
        builder.append(bodyContent);
        builder.append("\n</body>\n</html>");
        return builder.toString();
    }

    public static CssLength toCssLength(Integer twips) {
        if (twips == null) {
            return null;
        }
        return CssLength.points(twips / 20.0d);
    }

    public static CssLength computeTextIndent(WordDocument.Indentation indentation) {
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

    public static String computeLineHeight(WordDocument.Spacing spacing) {
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

    public static String alignmentToCss(WordDocument.Alignment alignment) {
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

    public static String formatDecimal(double value) {
        BigDecimal decimal = BigDecimal.valueOf(value).setScale(3, RoundingMode.HALF_UP).stripTrailingZeros();
        return decimal.toPlainString();
    }

    public static double roundLength(double value) {
        return BigDecimal.valueOf(value).setScale(3, RoundingMode.HALF_UP).doubleValue();
    }

    public static Optional<String> normalizeColor(String value) {
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

    public static Optional<String> normalizeHighlight(String value) {
        if (value == null || value.isBlank() || "none".equalsIgnoreCase(value)) {
            return Optional.empty();
        }
        String key = value.trim().toLowerCase(Locale.ROOT);
        if (HIGHLIGHT_COLORS.containsKey(key)) {
            return Optional.of(HIGHLIGHT_COLORS.get(key));
        }
        return normalizeColor(value);
    }

    public static String normalizeVerticalAlignment(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        return switch (value) {
            case "superscript" -> "super";
            case "subscript" -> "sub";
            default -> null;
        };
    }

    public static String paragraphShadingColor(WordDocument.ParagraphProperties properties,
                                                Map<String, String> themeColors) {
        if (properties == null) {
            return null;
        }
        return properties.rawProperties()
                .map(element -> shadingColorFromElement(element, themeColors))
                .orElse(null);
    }

    private static String tableShadingColor(WordDocument.TableProperties properties,
                                            Map<String, String> themeColors) {
        if (properties == null) {
            return null;
        }
        return properties.rawProperties()
                .map(element -> shadingColorFromElement(element, themeColors))
                .orElse(null);
    }

    private static String tableRowShadingColor(WordDocument.TableRowProperties properties,
                                               Map<String, String> themeColors) {
        if (properties == null) {
            return null;
        }
        return properties.rawProperties()
                .map(element -> shadingColorFromElement(element, themeColors))
                .orElse(null);
    }

    private static String tableCellShadingColor(WordDocument.TableCellProperties properties,
                                                Map<String, String> themeColors) {
        if (properties == null) {
            return null;
        }
        return properties.rawProperties()
                .map(element -> shadingColorFromElement(element, themeColors))
                .orElse(null);
    }

    public static String shadingColorFromElement(Element parent, Map<String, String> themeColors) {
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
    public static String decorationStyle(String underlineType, boolean doubleStrike) {
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

    public static List<String> computeFontStack(Map<String, String> fonts) {
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
        if (unique.add("system-ui")) {
            stack.add("system-ui");
        }
        if (unique.add("sans-serif")) {
            stack.add("sans-serif");
        }
        if (unique.add("serif")) {
            stack.add("serif");
        }
        return stack;
    }

    public static String fontStackToCss(List<String> fontStack) {
        return fontStack.stream()
                .map(DocxToHtml::normalizeFontName)
                .filter(name -> !name.isBlank())
                .collect(Collectors.joining(", "));
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
            return "\"" + trimmed.replace("\"", "\"") + "\"";
        }
        return trimmed;
    }

    private static String escapeHtml(String value) {
        if (value == null || value.isEmpty()) {
            return "";
        }
        StringBuilder builder = new StringBuilder(value.length());
        for (int i = 0; i < value.length(); i++) {
            builder.append(escapeChar(value.charAt(i)));
        }
        return builder.toString();
    }

    private static String escapeHtmlAttribute(String value) {
        return escapeHtml(value).replace("\n", " ").replace("\r", " ");
    }

    private static String escapeChar(char c) {
        return switch (c) {
            case '&' -> "&amp;";
            case '<' -> "&lt;";
            case '>' -> "&gt;";
            case '"' -> "&quot;";
            case '\'' -> "&#39;";
            default -> String.valueOf(c);
        };
    }

    private static Map<String, String> createHighlightColorMap() {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("yellow", "#ffff00");
        map.put("brightgreen", "#ccff00");
        map.put("green", "#00ff00");
        map.put("cyan", "#00ffff");
        map.put("turquoise", "#40e0d0");
        map.put("pink", "#ffb6c1");
        map.put("magenta", "#ff00ff");
        map.put("blue", "#0000ff");
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
        map.put("red", "#ff0000");
        map.put("teal", "#008080");
        map.put("violet", "#8000ff");
        map.put("orange", "#ffa500");
        return Map.copyOf(map);
    }

}















































