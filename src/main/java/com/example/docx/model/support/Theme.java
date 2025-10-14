package com.example.docx.model.support;

import org.w3c.dom.Document;

import java.util.Optional;

/**
 * Wrapper for {@code word/theme/themeX.xml}.
 */
public final class Theme {
    private final String partName;
    private final Document document;

    public Theme(String partName, Document document) {
        this.partName = partName;
        this.document = document;
    }

    public String partName() {
        return partName;
    }

    public Optional<Document> document() {
        return Optional.ofNullable(document);
    }
}
