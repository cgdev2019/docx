package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.relationship.RelationshipSet;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Parses relationship parts.
 */
public final class RelationshipsParser {

    public RelationshipSet parse(DocxArchive archive, String partName) throws IOException {
        if (!archive.exists(partName)) {
            return null;
        }
        try (InputStream input = archive.open(partName)) {
            Document document = XmlUtils.parse(input);
            Element root = document.getDocumentElement();
            Map<String, RelationshipSet.Relationship> entries = new LinkedHashMap<>();
            for (Element rel : XmlUtils.children(root, Namespaces.RELATIONSHIPS, "Relationship")) {
                String id = rel.getAttribute("Id");
                String type = rel.getAttribute("Type");
                String target = rel.getAttribute("Target");
                String targetMode = rel.getAttribute("TargetMode");
                targetMode = targetMode.isEmpty() ? null : targetMode;
                entries.put(id, new RelationshipSet.Relationship(id, type, target, targetMode));
            }
            return new RelationshipSet(entries);
        }
    }
}
