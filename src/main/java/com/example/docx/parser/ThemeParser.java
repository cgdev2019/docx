package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.support.Theme;
import org.w3c.dom.Document;

import java.io.IOException;
import java.io.InputStream;

/**
 * Reads the primary theme part if present.
 */
public final class ThemeParser {

    public Theme parse(DocxArchive archive) throws IOException {
        String path = "word/theme/theme1.xml";
        if (!archive.exists(path)) {
            return null;
        }
        try (InputStream input = archive.open(path)) {
            Document document = XmlUtils.parse(input);
            return new Theme(path, document);
        }
    }
}
