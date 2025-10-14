package com.example.docx.html;

import com.example.docx.model.document.WordDocument;

import java.util.List;

final class StructuredDocumentTagRenderer {
    private final RenderContext context;
    private final BlockRenderer blockRenderer;

    StructuredDocumentTagRenderer(RenderContext context, BlockRenderer blockRenderer) {
        this.context = context;
        this.blockRenderer = blockRenderer;
    }

    String renderStructuredDocumentTag(WordDocument.StructuredDocumentTag sdt,
                                       List<WordDocument.RunProperties> extraRunFallbacks) {
        StringBuilder builder = new StringBuilder();
        builder.append("<section class=\"docx-sdt\"");
        sdt.properties().tag().ifPresent(tag -> builder.append(" data-tag=\"").append(DocxHtmlUtils.escapeHtmlAttribute(tag)).append("\""));
        sdt.properties().alias().ifPresent(alias -> builder.append(" data-alias=\"").append(DocxHtmlUtils.escapeHtmlAttribute(alias)).append("\""));
        sdt.properties().id().ifPresent(id -> builder.append(" data-id=\"").append(DocxHtmlUtils.escapeHtmlAttribute(id)).append("\""));
        builder.append(">");
        for (WordDocument.Block child : sdt.content()) {
            String blockContent = blockRenderer.renderBlock(child, extraRunFallbacks);
            if (!blockContent.isEmpty()) {
                builder.append(blockContent);
            }
        }
        builder.append("</section>");
        return builder.toString();
    }
}
