package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.support.ContentTypes;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Parses the package {@code [Content_Types].xml} part.
 */
public final class ContentTypesParser {

    public ContentTypes parse(DocxArchive archive) throws IOException {
        if (!archive.exists("[Content_Types].xml")) {
            return null;
        }
        try (InputStream input = archive.open("[Content_Types].xml")) {
            Document document = XmlUtils.parse(input);
            Element root = document.getDocumentElement();
            Map<String, String> defaults = new LinkedHashMap<>();
            Map<String, String> overrides = new LinkedHashMap<>();
            for (Element child : XmlUtils.childElements(root)) {
                switch (child.getLocalName()) {
                    case "Default" -> {
                        String extension = child.getAttribute("Extension");
                        String type = child.getAttribute("ContentType");
                        defaults.put(extension, type);
                    }
                    case "Override" -> {
                        String partName = child.getAttribute("PartName");
                        String type = child.getAttribute("ContentType");
                        if (partName.startsWith("/")) {
                            partName = partName.substring(1);
                        }
                        overrides.put(partName, type);
                    }
                    default -> {
                    }
                }
            }
            return new ContentTypes(defaults, overrides);
        }
    }
}
