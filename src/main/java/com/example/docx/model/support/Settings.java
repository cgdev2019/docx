package com.example.docx.model.support;

import org.w3c.dom.Document;

import java.util.Optional;

/**
 * Wrapper for {@code word/settings.xml}.
 */
public final class Settings {
    private final Document document;

    public Settings(Document document) {
        this.document = document;
    }

    public Optional<Document> document() {
        return Optional.ofNullable(document);
    }
}
