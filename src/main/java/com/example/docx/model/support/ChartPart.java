package com.example.docx.model.support;

import org.w3c.dom.Document;

import java.util.Optional;

/**
 * Represents a chart definition stored under {@code word/charts/}.
 */
public final class ChartPart {

    private final String partName;
    private final Document document;

    public ChartPart(String partName, Document document) {
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
