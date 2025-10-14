package com.example.docx;

import com.example.docx.model.DocxPackage;
import com.example.docx.model.document.WordDocument;
import com.example.docx.model.styles.StyleDefinitions;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

class DocxReaderFileSamplesTest {

    private final DocxReader reader = new DocxReader();

    @ParameterizedTest
    @ValueSource(strings = {"file-sample_100kB.docx", "file-sample_500kB.docx", "file-sample_1MB.docx"})
    void chartAndMediaPartsAreLoaded(String sample) {
        DocxPackage docx = reader.read(Path.of("samples", sample));

        assertEquals(1, docx.charts().size(), "chart count");
        assertTrue(docx.charts().containsKey("word/charts/chart1.xml"));
        assertTrue(docx.relationshipsByPart().containsKey("word/document.xml"), "document relationships");

        assertEquals(1, docx.mediaFiles().size(), "media count");
        assertTrue(docx.mediaFiles().containsKey("word/media/image1.jpeg"));
        assertNotNull(docx.mediaFiles().get("word/media/image1.jpeg"));
        assertTrue(docx.mediaFiles().get("word/media/image1.jpeg").data().length > 0);
    }

    @ParameterizedTest
    @ValueSource(strings = {"file-sample_100kB.docx", "file-sample_500kB.docx", "file-sample_1MB.docx"})
    void basicDocumentStructureIsPresent(String sample) {
        DocxPackage docx = reader.read(Path.of("samples", sample));
        WordDocument document = docx.document().orElseThrow();

        assertFalse(document.bodyElements().isEmpty(), "body elements");
        assertTrue(document.bodySectionProperties().isPresent(), "section properties");

        StyleDefinitions styles = docx.styles().orElseThrow();
        assertFalse(styles.styles().isEmpty(), "styles present");
        assertNotNull(styles.defaultParagraphStyleHierarchy());
    }
}

