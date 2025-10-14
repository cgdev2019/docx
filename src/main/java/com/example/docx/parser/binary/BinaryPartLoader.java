package com.example.docx.parser.binary;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.DocxPackage;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Set;

/**
 * Collects additional binary parts (embedded objects, fonts, ink data, ...).
 */
public final class BinaryPartLoader {

    private static final List<String> DIRECTORIES = List.of(
            "word/fonts",
            "word/embeddings",
            "word/ink"
    );

    public void load(DocxArchive archive, DocxPackage.Builder builder) throws IOException {
        for (String directory : DIRECTORIES) {
            Set<String> entries = archive.list(directory);
            for (String entry : entries) {
                if (entry.endsWith("/")) {
                    continue;
                }
                try (InputStream input = archive.open(entry)) {
                    builder.binaryPart(entry, input.readAllBytes());
                }
            }
        }
    }
}
