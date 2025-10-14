package com.example.docx.html;

import com.example.docx.model.document.WordDocument;

import java.util.List;

final class BlockRenderer {
    private final RenderContext context;
    private final ParagraphRenderer paragraphRenderer;
    private final TableRenderer tableRenderer;
    private final StructuredDocumentTagRenderer structuredDocumentTagRenderer;

    BlockRenderer(RenderContext context) {
        this.context = context;
        this.paragraphRenderer = new ParagraphRenderer(context);
        this.tableRenderer = new TableRenderer(context, this);
        this.structuredDocumentTagRenderer = new StructuredDocumentTagRenderer(context, this);
    }

    String renderBlocks(List<WordDocument.Block> blocks) {
        if (blocks == null || blocks.isEmpty()) {
            return "";
        }
        StringBuilder builder = new StringBuilder();
        for (WordDocument.Block block : blocks) {
            String html = renderBlock(block, List.of());
            if (!html.isEmpty()) {
                builder.append(html).append('\n');
            }
        }
        return builder.toString();
    }

    String renderBlock(WordDocument.Block block, List<WordDocument.RunProperties> extraRunFallbacks) {
        if (block instanceof WordDocument.Paragraph paragraph) {
            return paragraphRenderer.renderParagraph(paragraph, extraRunFallbacks);
        }
        if (block instanceof WordDocument.Table table) {
            return tableRenderer.renderTable(table, extraRunFallbacks);
        }
        if (block instanceof WordDocument.StructuredDocumentTag sdt) {
            return structuredDocumentTagRenderer.renderStructuredDocumentTag(sdt, extraRunFallbacks);
        }
        if (block instanceof WordDocument.SectionBreak) {
            return "<span class=\"docx-section-break\"></span>";
        }
        if (block instanceof WordDocument.Bookmark bookmark) {
            return renderBookmark(bookmark);
        }
        return "";
    }

    private String renderBookmark(WordDocument.Bookmark bookmark) {
        if (bookmark.kind() != WordDocument.Bookmark.Kind.START) {
            return "";
        }
        String name = bookmark.name().orElse(bookmark.id().orElse(null));
        if (name == null || name.isBlank()) {
            return "";
        }
        return "<a id=\"" + DocxHtmlUtils.escapeHtmlAttribute(name) + "\"></a>";
    }

    RenderContext context() {
        return context;
    }
}
