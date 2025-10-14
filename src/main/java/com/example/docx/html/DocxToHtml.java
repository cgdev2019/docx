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
import java.util.EnumMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

public final class DocxToHtml {

    private static final String DOCUMENT_PART = "word/document.xml";
    private static final WordDocument.RunProperties EMPTY_RUN_PROPERTIES =
            new WordDocument.RunProperties(null, false, false, false, null,
                    false, false, false, false, false, null, null, null,
                    null, null, Map.of(), null);
    private static final WordDocument.ParagraphProperties EMPTY_PARAGRAPH_PROPERTIES =
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
        StyleResolver.ResolvedTableStyle tableStyle = styleResolver.resolveTableStyle(table.properties());
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

    private static CssLength toCssLength(Integer twips) {
        if (twips == null) {
            return null;
        }
        return CssLength.points(twips / 20.0d);
    }

    private static CssLength computeTextIndent(WordDocument.Indentation indentation) {
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

    private static String computeLineHeight(WordDocument.Spacing spacing) {
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

    private static String alignmentToCss(WordDocument.Alignment alignment) {
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

    private static String formatDecimal(double value) {
        BigDecimal decimal = BigDecimal.valueOf(value).setScale(3, RoundingMode.HALF_UP).stripTrailingZeros();
        return decimal.toPlainString();
    }

    private static double roundLength(double value) {
        return BigDecimal.valueOf(value).setScale(3, RoundingMode.HALF_UP).doubleValue();
    }

    private static Optional<String> normalizeColor(String value) {
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

    private static Optional<String> normalizeHighlight(String value) {
        if (value == null || value.isBlank() || "none".equalsIgnoreCase(value)) {
            return Optional.empty();
        }
        String key = value.trim().toLowerCase(Locale.ROOT);
        if (HIGHLIGHT_COLORS.containsKey(key)) {
            return Optional.of(HIGHLIGHT_COLORS.get(key));
        }
        return normalizeColor(value);
    }

    private static String normalizeVerticalAlignment(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        return switch (value) {
            case "superscript" -> "super";
            case "subscript" -> "sub";
            default -> null;
        };
    }

    private static String paragraphShadingColor(WordDocument.ParagraphProperties properties,
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

    private static String shadingColorFromElement(Element parent, Map<String, String> themeColors) {
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
    private static String decorationStyle(String underlineType, boolean doubleStrike) {
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

    private static List<String> computeFontStack(Map<String, String> fonts) {
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

    private static String fontStackToCss(List<String> fontStack) {
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
    private static final class HyperlinkResolver {
        private final RelationshipSet relationships;

        HyperlinkResolver(RelationshipSet relationships) {
            this.relationships = relationships;
        }

        Optional<String> resolve(String relationshipId, String anchor) {
            String target = null;
            if (relationshipId != null && relationships != null) {
                RelationshipSet.Relationship rel = relationships.byId(relationshipId).orElse(null);
                if (rel != null) {
                    target = rel.target();
                    if (rel.targetMode().filter(mode -> "External".equalsIgnoreCase(mode)).isEmpty() && target != null && !target.startsWith("#") && !target.startsWith("http")) {
                        target = sanitizeInternalTarget(target);
                    }
                }
            }
            if (anchor != null && !anchor.isBlank()) {
                if (target == null || target.isBlank()) {
                    target = "#" + anchor;
                } else if (!target.contains("#")) {
                    target = target + "#" + anchor;
                }
            }
            return Optional.ofNullable(target);
        }

        private String sanitizeInternalTarget(String target) {
            if (target == null) {
                return null;
            }
            if (target.startsWith("/")) {
                return target;
            }
            if (target.startsWith("#")) {
                return target;
            }
            return target.replace('\\', '/');
        }
    }

    private static final class CssLength {
        private final double value;
        private final String unit;

        private CssLength(double value, String unit) {
            this.value = roundLength(value);
            this.unit = unit;
        }

        static CssLength points(double value) {
            return new CssLength(value, "pt");
        }

        CssLength negate() {
            return new CssLength(-value, unit);
        }

        String css() {
            if (Math.abs(value) < 0.0005d) {
                return "0";
            }
            return formatDecimal(value) + unit;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) {
                return true;
            }
            if (!(o instanceof CssLength that)) {
                return false;
            }
            return Double.compare(that.value, value) == 0 && Objects.equals(unit, that.unit);
        }

        @Override
        public int hashCode() {
            return Objects.hash(value, unit);
        }
    }

    private record ParagraphCss(String textAlign,
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
        static ParagraphCss empty() {
            return new ParagraphCss(null, null, null, null, null, null, null, null, null, null, false, false, false);
        }

        static ParagraphCss from(StyleResolver.ResolvedParagraph resolved, StyleResolver.ResolvedRun defaultRun) {
            if (resolved == null) {
                return ParagraphCss.empty();
            }
            WordDocument.Spacing spacing = resolved.spacing();
            CssLength top = spacing != null ? toCssLength(spacing.before().orElse(null)) : null;
            CssLength bottom = spacing != null ? toCssLength(spacing.after().orElse(null)) : null;
            WordDocument.Indentation indentation = resolved.indentation();
            CssLength left = indentation != null ? toCssLength(indentation.left().orElse(null)) : null;
            CssLength right = indentation != null ? toCssLength(indentation.right().orElse(null)) : null;
            CssLength indent = computeTextIndent(indentation);
            String lineHeight = computeLineHeight(spacing);
            CssLength paragraphFontSize = defaultRun.fontSizeHalfPoints()
                    .map(size -> CssLength.points(size / 2.0d))
                    .orElse(null);
            String paragraphFontFamily = fontStackToCss(defaultRun.fontStack());
            if (paragraphFontFamily != null && paragraphFontFamily.isBlank()) {
                paragraphFontFamily = null;
            }
            String background = resolved.shadingColor();
            return new ParagraphCss(alignmentToCss(resolved.alignment()), top, bottom, left, right, indent,
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
    private record RunCss(boolean bold,
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
            String fontFamily = fontStackToCss(resolved.fontStack());
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
    private record TableCss(String backgroundColor) {

        boolean hasColor() {
            return backgroundColor != null && !backgroundColor.isBlank();
        }

        String declarations() {
            if (!hasColor()) {
                return "";
            }
            return "background-color:" + backgroundColor;
        }
    }

    private record TableRowCss(String backgroundColor) {

        boolean hasColor() {
            return backgroundColor != null && !backgroundColor.isBlank();
        }

        String declarations() {
            if (!hasColor()) {
                return "";
            }
            return "background-color:" + backgroundColor;
        }
    }

    private record TableCellCss(String backgroundColor) {

        boolean hasColor() {
            return backgroundColor != null && !backgroundColor.isBlank();
        }

        String declarations() {
            if (!hasColor()) {
                return "";
            }
            return "background-color:" + backgroundColor;
        }
    }
    private static final class StyleRegistry {
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
            return formatDecimal(cm) + "cm";
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
    private static final class StyleResolver {
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
            WordDocument.ParagraphProperties effective = properties == null ? EMPTY_PARAGRAPH_PROPERTIES : properties;
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
            String shading = paragraphShadingColor(effective, themeColors);
            if (shading == null) {
                for (WordDocument.ParagraphProperties fallback : paragraphFallbacks) {
                    shading = paragraphShadingColor(fallback, themeColors);
                    if (shading != null) {
                        break;
                    }
                }
            }

            return new ResolvedParagraph(alignment, indentation, spacing, keepTogether, keepWithNext, pageBreakBefore, shading, List.copyOf(runFallbacks));
        }

        ResolvedRun resolveRun(WordDocument.RunProperties properties, ResolvedParagraph paragraph) {
            WordDocument.RunProperties effective = properties == null ? EMPTY_RUN_PROPERTIES : properties;
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
            List<String> fontStack = computeFontStack(fonts);
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
                            .map(element -> shadingColorFromElement(element, themeColors))
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
            String background = shadingColorFromElement(tcPr, themeColors);
            if (background == null) {
                Element tblPr = XmlUtils.firstChild(tblStylePr, Namespaces.WORD_MAIN, "tblPr").orElse(null);
                background = shadingColorFromElement(tblPr, themeColors);
            }
            Element rPr = XmlUtils.firstChild(tblStylePr, Namespaces.WORD_MAIN, "rPr").orElse(null);
            WordDocument.RunProperties runProperties = toRunProperties(rPr);
            if (background == null && runProperties == null) {
                return null;
            }
            return new RegionStyle(background, runProperties);
        }

        private static final class ResolvedTableStyle {
            private final String tableBackground;
            private final Map<TableStyleRegion, RegionStyle> regions;

            ResolvedTableStyle(String tableBackground, Map<TableStyleRegion, RegionStyle> regions) {
                this.tableBackground = tableBackground;
                this.regions = regions.isEmpty() ? Map.of() : Map.copyOf(regions);
            }

            static ResolvedTableStyle empty() {
                return new ResolvedTableStyle(null, Map.of());
            }

            String tableBackground() {
                return tableBackground;
            }

            RegionStyle rowRegion(int rowIndex, int rowCount) {
                if (rowIndex == 0) {
                    RegionStyle first = regions.get(TableStyleRegion.FIRST_ROW);
                    if (first != null) {
                        return first;
                    }
                }
                if (rowCount > 0 && rowIndex == rowCount - 1) {
                    RegionStyle last = regions.get(TableStyleRegion.LAST_ROW);
                    if (last != null) {
                        return last;
                    }
                }
                return regions.get(TableStyleRegion.WHOLE_TABLE);
            }
        }

        private enum TableStyleRegion {
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

        private record RegionStyle(String backgroundColor,
                                   WordDocument.RunProperties runProperties) {
        }
        record ResolvedParagraph(WordDocument.Alignment alignment,
                                 WordDocument.Indentation indentation,
                                 WordDocument.Spacing spacing,
                                 boolean keepTogether,
                                 boolean keepWithNext,
                                 boolean pageBreakBefore,
                                 String shadingColor,
                                 List<WordDocument.RunProperties> runFallbacks) {
        }

        record ResolvedRun(boolean bold,
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
}















































