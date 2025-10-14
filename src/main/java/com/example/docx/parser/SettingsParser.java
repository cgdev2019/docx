package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.support.Settings;
import com.example.docx.model.support.WebSettings;
import org.w3c.dom.Document;

import java.io.IOException;
import java.io.InputStream;

/**
 * Reads Word settings parts.
 */
public final class SettingsParser {

    public Settings parseSettings(DocxArchive archive) throws IOException {
        Document document = parseDocument(archive, "word/settings.xml");
        return document != null ? new Settings(document) : null;
    }

    public WebSettings parseWebSettings(DocxArchive archive) throws IOException {
        Document document = parseDocument(archive, "word/webSettings.xml");
        return document != null ? new WebSettings(document) : null;
    }

    private Document parseDocument(DocxArchive archive, String path) throws IOException {
        if (!archive.exists(path)) {
            return null;
        }
        try (InputStream input = archive.open(path)) {
            return XmlUtils.parse(input);
        }
    }
}
