package com.example.docx.model.support;

import org.w3c.dom.Document;

import java.util.Optional;

/**
 * Represents an arbitrary custom XML part stored under {@code customXml/}.
 */
public final class CustomXmlPart {
    private final String partName;
    private final Document document;

    public CustomXmlPart(String partName, Document document) {
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
