package com.example.docx.html;

import com.example.docx.model.document.WordDocument;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

final class TableRenderer {
    private final RenderContext context;
    private final BlockRenderer blockRenderer;

    TableRenderer(RenderContext context, BlockRenderer blockRenderer) {
        this.context = context;
        this.blockRenderer = blockRenderer;
    }

    String renderTable(WordDocument.Table table, List<WordDocument.RunProperties> inheritedRunFallbacks) {
        ResolvedTableStyle tableStyle = context.styleResolver().resolveTableStyle(table.properties());
        int rowCount = table.rows().size();
        String tableBackground = DocxHtmlUtils.tableShadingColor(table.properties(), context.themeColors());
        if (tableBackground == null) {
            tableBackground = tableStyle.tableBackground();
        }
        DocxHtmlUtils.TableBorders directBorders = DocxHtmlUtils.tableBorders(table.properties(), context.themeColors());
        BorderDefinition tablePerimeter = tableStyle.tableBorders().overrideWith(directBorders.perimeter());
        BorderDefinition insideHorizontal = tableStyle.insideHorizontal().overrideWith(directBorders.insideHorizontal());
        BorderDefinition insideVertical = tableStyle.insideVertical().overrideWith(directBorders.insideVertical());
        List<String> tableClasses = new ArrayList<>();
        tableClasses.add("docx-table");
        TableCss tableCss = new TableCss(tableBackground, tablePerimeter);
        String tableClass = context.styleRegistry().registerTable(tableCss);
        if (tableClass != null) {
            tableClasses.add(tableClass);
        }
        StringBuilder builder = new StringBuilder();
        builder.append("<table class=\"").append(String.join(" ", tableClasses)).append("\">");
        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            WordDocument.TableRow row = table.rows().get(rowIndex);
            StyleResolver.RegionStyle region = tableStyle.rowRegion(row.properties(), rowIndex, rowCount);
            String rowBackground = DocxHtmlUtils.tableRowShadingColor(row.properties(), context.themeColors());
            if (rowBackground == null && region != null) {
                rowBackground = region.backgroundColor();
            }
            List<WordDocument.RunProperties> rowFallbacks = new ArrayList<>(inheritedRunFallbacks);
            if (region != null && region.runProperties() != null) {
                rowFallbacks.add(region.runProperties());
            }
            List<String> rowClasses = new ArrayList<>();
            TableRowCss rowCss = new TableRowCss(rowBackground);
            String rowClass = context.styleRegistry().registerRow(rowCss);
            if (rowClass != null) {
                rowClasses.add("docx-row");
                rowClasses.add(rowClass);
            }
            builder.append("<tr");
            if (!rowClasses.isEmpty()) {
                builder.append(" class=\"").append(String.join(" ", rowClasses)).append("\"");
            }
            builder.append(">");
            int columnCount = row.cells().size();
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                WordDocument.TableCell cell = row.cells().get(columnIndex);
                BorderDefinition cellBorder = DocxHtmlUtils.tableCellBorders(cell.properties(), context.themeColors());
                if (region != null && region.borders() != null && !region.borders().isEmpty()) {
                    cellBorder = cellBorder.fillMissing(region.borders());
                }
                BorderDefinition.BorderEdge topEdge = cellBorder.top();
                if (topEdge == null) {
                    topEdge = (rowIndex == 0) ? tablePerimeter.top() : insideHorizontal.top();
                }
                BorderDefinition.BorderEdge bottomEdge = cellBorder.bottom();
                if (bottomEdge == null) {
                    bottomEdge = (rowIndex == rowCount - 1) ? tablePerimeter.bottom() : insideHorizontal.bottom();
                }
                BorderDefinition.BorderEdge leftEdge = cellBorder.left();
                if (leftEdge == null) {
                    leftEdge = (columnIndex == 0) ? tablePerimeter.left() : insideVertical.left();
                }
                BorderDefinition.BorderEdge rightEdge = cellBorder.right();
                if (rightEdge == null) {
                    rightEdge = (columnIndex == columnCount - 1) ? tablePerimeter.right() : insideVertical.right();
                }
                BorderDefinition finalBorder = BorderDefinition.of(topEdge, rightEdge, bottomEdge, leftEdge);
                builder.append(renderTableCell(cell, rowBackground, tableBackground, rowFallbacks, finalBorder));
            }
            builder.append("</tr>");
        }
        builder.append("</table>");
        return builder.toString();
    }

    private String renderTableCell(WordDocument.TableCell cell,
                                   String rowBackground,
                                   String tableBackground,
                                   List<WordDocument.RunProperties> rowRunFallbacks,
                                   BorderDefinition borders) {
        List<WordDocument.RunProperties> effectiveFallbacks = rowRunFallbacks == null || rowRunFallbacks.isEmpty()
                ? List.of()
                : List.copyOf(rowRunFallbacks);
        StringBuilder content = new StringBuilder();
        for (WordDocument.Block block : cell.content()) {
            String fragment = blockRenderer.renderBlock(block, effectiveFallbacks);
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
        String background = DocxHtmlUtils.tableCellShadingColor(cell.properties(), context.themeColors());
        if (Objects.equals(background, rowBackground) || Objects.equals(background, tableBackground)) {
            background = null;
        }
        TableCellCss cellCss = new TableCellCss(background, borders);
        String cellClass = context.styleRegistry().registerCell(cellCss);
        if (cellClass != null) {
            classes.add(cellClass);
        }
        StringBuilder builder = new StringBuilder();
        builder.append("<td");
        if (!classes.isEmpty()) {
            builder.append(" class=\"").append(String.join(" ", classes)).append("\"");
        }
        cell.properties().gridSpan().ifPresent(span -> {
            if (span != null && span > 1) {
                builder.append(" colspan=\"").append(span).append("\"");
            }
        });
        builder.append(">");
        builder.append(content);
        builder.append("</td>");
        return builder.toString();
    }
}
