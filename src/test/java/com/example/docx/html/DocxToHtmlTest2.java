package com.example.docx.html;

import com.example.docx.DocxReader;
import com.example.docx.model.DocxPackage;
import com.example.docx.model.document.WordDocument;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

class DocxToHtmlTest2 {

    private final DocxReader reader = new DocxReader();

    @Test
    void convertsParagraphWithStylesAndRuns() throws IOException {
        DocxPackage docx = reader.read(Path.of("samples", "demo.docx"));

        DocxToHtml converter = new DocxToHtml("fr");
        String html = converter.convert(docx);

        Path out = Path.of("C:/tmp/out.html");
        Files.writeString(out, html, StandardCharsets.UTF_8,
                StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING);
    }
}
