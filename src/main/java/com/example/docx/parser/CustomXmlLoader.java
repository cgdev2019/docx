package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.DocxPackage;
import com.example.docx.model.support.CustomXmlPart;
import org.w3c.dom.Document;

import java.io.IOException;
import java.io.InputStream;
import java.util.Set;

/**
 * Loads XML parts stored under {@code customXml/}.
 */
public final class CustomXmlLoader {

    public void load(DocxArchive archive, DocxPackage.Builder builder) throws IOException {
        Set<String> entries = archive.list("customXml");
        for (String entry : entries) {
            if (!entry.endsWith(".xml")) {
                continue;
            }
            try (InputStream input = archive.open(entry)) {
                Document document = XmlUtils.parse(input);
                builder.customXmlPart(new CustomXmlPart(entry, document));
            }
        }
    }
}
