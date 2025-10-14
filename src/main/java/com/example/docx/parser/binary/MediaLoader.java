package com.example.docx.parser.binary;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.DocxPackage;
import com.example.docx.model.support.ContentTypes;
import com.example.docx.model.support.MediaFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.Set;

/**
 * Loads binary media files located under {@code word/media}.
 */
public final class MediaLoader {

    public void loadMedia(DocxArchive archive, ContentTypes contentTypes, DocxPackage.Builder builder) throws IOException {
        Set<String> entries = archive.list("word/media");
        for (String entry : entries) {
            if (entry.endsWith("/")) {
                continue;
            }
            try (InputStream input = archive.open(entry)) {
                byte[] data = input.readAllBytes();
                String contentType = contentTypes != null ? contentTypes.lookup("/" + entry) : null;
                builder.mediaFile(new MediaFile(entry, contentType, data));
            }
        }
    }
}
