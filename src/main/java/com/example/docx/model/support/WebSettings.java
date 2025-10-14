package com.example.docx.model.support;

import org.w3c.dom.Document;

import java.util.Optional;

/**
 * Wrapper for {@code word/webSettings.xml}.
 */
public final class WebSettings {
    private final Document document;

    public WebSettings(Document document) {
        this.document = document;
    }

    public Optional<Document> document() {
        return Optional.ofNullable(document);
    }
}
