package com.example.docx.html;

import com.example.docx.model.document.WordDocument;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

final class ParagraphRenderer {
    private final RenderContext context;

    ParagraphRenderer(RenderContext context) {
        this.context = context;
    }

    String renderParagraph(WordDocument.Paragraph paragraph, List<WordDocument.RunProperties> extraRunFallbacks) {
        StyleResolver.ResolvedParagraph resolved = context.styleResolver().resolveParagraph(paragraph.properties(), extraRunFallbacks);
        StyleResolver.ResolvedRun defaultRun = context.styleResolver().resolveRun(DocxToHtml.EMPTY_RUN_PROPERTIES, resolved);
        ParagraphCss css = ParagraphCss.from(resolved, defaultRun);
        String paragraphClass = context.styleRegistry().registerParagraph(css);
        StringBuilder inner = new StringBuilder();
        for (WordDocument.ParagraphContent content : paragraph.content()) {
            inner.append(renderParagraphContent(content, resolved));
        }
        if (inner.length() == 0) {
            inner.append("&nbsp;");
        }
        return "<p class=\"docx-paragraph " + paragraphClass + "\">" + inner + "</p>";
    }

    private String renderParagraphContent(WordDocument.ParagraphContent content,
                                          StyleResolver.ResolvedParagraph paragraph) {
        if (content instanceof WordDocument.Run run) {
            return renderRun(run, paragraph);
        }
        if (content instanceof WordDocument.Hyperlink hyperlink) {
            return renderHyperlink(hyperlink, paragraph);
        }
        if (content instanceof WordDocument.BookmarkStart bookmarkStart) {
            return renderBookmarkStart(bookmarkStart);
        }
        if (content instanceof WordDocument.Field field) {
            return renderField(field, paragraph);
        }
        if (content instanceof WordDocument.StructuredDocumentTagRun sdt) {
            return renderStructuredDocumentTagRun(sdt, paragraph);
        }
        return "";
    }

    private String renderHyperlink(WordDocument.Hyperlink hyperlink, StyleResolver.ResolvedParagraph paragraph) {
        StringBuilder content = new StringBuilder();
        for (WordDocument.Run run : hyperlink.runs()) {
            content.append(renderRun(run, paragraph));
        }
        if (content.length() == 0) {
            return "";
        }
        String href = context.hyperlinkResolver()
                .resolve(hyperlink.relationshipId().orElse(null), hyperlink.anchor().orElse(null))
                .map(DocxHtmlUtils::escapeHtmlAttribute)
                .orElse("#");
        return "<a class=\"docx-link\" href=\"" + href + "\">" + content + "</a>";
    }

    private String renderField(WordDocument.Field field, StyleResolver.ResolvedParagraph paragraph) {
        StringBuilder result = new StringBuilder();
        for (WordDocument.Run run : field.resultRuns()) {
            result.append(renderRun(run, paragraph));
        }
        if (result.length() == 0) {
            for (WordDocument.Run run : field.instructionRuns()) {
                result.append(renderRun(run, paragraph));
            }
        }
        if (result.length() == 0) {
            return "";
        }
        return "<span class=\"docx-field\">" + result + "</span>";
    }

    private String renderStructuredDocumentTagRun(WordDocument.StructuredDocumentTagRun sdt,
                                                  StyleResolver.ResolvedParagraph paragraph) {
        StringBuilder builder = new StringBuilder();
        builder.append("<span class=\"docx-sdt-inline\"");
        sdt.properties().tag().ifPresent(tag -> builder.append(" data-tag=\"").append(DocxHtmlUtils.escapeHtmlAttribute(tag)).append("\""));
        sdt.properties().alias().ifPresent(alias -> builder.append(" data-alias=\"").append(DocxHtmlUtils.escapeHtmlAttribute(alias)).append("\""));
        sdt.properties().id().ifPresent(id -> builder.append(" data-id=\"").append(DocxHtmlUtils.escapeHtmlAttribute(id)).append("\""));
        builder.append(">");
        for (WordDocument.ParagraphContent child : sdt.content()) {
            builder.append(renderParagraphContent(child, paragraph));
        }
        builder.append("</span>");
        return builder.toString();
    }

    String renderRun(WordDocument.Run run, StyleResolver.ResolvedParagraph paragraph) {
        StyleResolver.ResolvedRun resolvedRun = context.styleResolver().resolveRun(run.properties(), paragraph);
        if (resolvedRun.vanish()) {
            return "";
        }
        StringBuilder content = new StringBuilder();
        for (WordDocument.Inline inline : run.elements()) {
            String fragment = renderInline(inline);
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
        String className = context.styleRegistry().registerRun(css);
        return "<span class=\"docx-span " + className + "\">" + content + "</span>";
    }

    private String renderInline(WordDocument.Inline inline) {
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
        drawing.relationshipId().ifPresent(rel -> builder.append(" data-rel=\"").append(DocxHtmlUtils.escapeHtmlAttribute(rel)).append("\""));
        if (drawing.width() > 0 && drawing.height() > 0) {
            builder.append(" data-size=\"").append(drawing.width()).append("x").append(drawing.height()).append("\"");
        }
        builder.append(">");
        builder.append("[Image");
        drawing.description().ifPresent(desc -> {
            if (!desc.isBlank()) {
                builder.append(": ").append(DocxHtmlUtils.escapeHtml(desc));
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
            return DocxHtmlUtils.escapeHtml(code == null ? "" : code);
        }
        return DocxHtmlUtils.escapeHtml(new String(Character.toChars(codePoint)));
    }

    private String renderText(WordDocument.Text text) {
        String value = text.text();
        if (value.isEmpty()) {
            return "";
        }
        /*if (text.preserveSpace()) {
            return value.replace(" ", "&nbsp;");
        }*/
        return DocxHtmlUtils.escapeHtml(value);
    }

    private String renderBookmarkStart(WordDocument.BookmarkStart bookmarkStart) {
        String name = bookmarkStart.name().orElse(bookmarkStart.id().orElse(null));
        if (name == null || name.isBlank()) {
            return "";
        }
        return "<a id=\"" + DocxHtmlUtils.escapeHtmlAttribute(name) + "\"></a>";
    }
}
