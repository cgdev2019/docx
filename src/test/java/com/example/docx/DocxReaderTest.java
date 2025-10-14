package com.example.docx;

import com.example.docx.model.DocxPackage;
import com.example.docx.model.document.WordDocument;
import com.example.docx.model.notes.NoteCollection;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.*;

class DocxReaderTest {

    private final DocxReader reader = new DocxReader();

    @Test
    void readsCoreProperties() {
        DocxPackage docx = reader.read(Path.of("samples", "demo.docx"));
        assertTrue(docx.coreProperties().isPresent(), "core properties");
        assertEquals("DOCX Demo", docx.coreProperties().get().title());
        assertEquals("Kovid Goyal", docx.coreProperties().get().creator());
    }

    @Test
    void readsDocumentStructure() {
        DocxPackage docx = reader.read(Path.of("samples", "demo.docx"));
        WordDocument document = docx.document().orElseThrow();
        assertFalse(document.bodyElements().isEmpty(), "document should contain blocks");
        WordDocument.Block firstBlock = document.bodyElements().get(0);
        assertTrue(firstBlock instanceof WordDocument.Paragraph, "first block is paragraph");
        WordDocument.Paragraph paragraph = (WordDocument.Paragraph) firstBlock;
        StringBuilder text = new StringBuilder();
        paragraph.content().stream()
                .filter(content -> content instanceof WordDocument.Run)
                .map(content -> (WordDocument.Run) content)
                .flatMap(run -> run.elements().stream())
                .filter(inline -> inline instanceof WordDocument.Text)
                .map(inline -> ((WordDocument.Text) inline).text())
                .forEach(text::append);
        assertEquals("Demonstration of DOCX support in calibre", text.toString());
    }

    @Test
    void loadsFootnotesMediaAndCharts() {
        List<String> samples = List.of(
                "demo.docx",
                "file-sample_100kB.docx",
                "file-sample_500kB.docx",
                "file-sample_1MB.docx"
        );
        for (String sample : samples) {
            DocxPackage docx = reader.read(Path.of("samples", sample));
            assertTrue(docx.document().isPresent(), "document missing for " + sample);
            assertFalse(docx.mediaFiles().isEmpty(), "media files should be extracted for " + sample);
            if (sample.startsWith("file-sample")) {
                assertFalse(docx.charts().isEmpty(), "charts should be loaded for " + sample);
            }
        }
    }

    @Test
    void unknownPartRaisesException() throws IOException {
        Path docx = createMinimalDocx("unknown-part", minimalParagraph("Unknown part"),
                Map.of("word/unknownPart.xml", "<w:ignored/>".getBytes(StandardCharsets.UTF_8)));
        assertThrows(DocxException.class, () -> reader.read(docx));
    }

    @Test
    void unknownTagRaisesException() throws IOException {
        String body = "<w:p><w:unknown/></w:p>";
        Path docx = createMinimalDocx("unknown-tag", wrapBody(body), Map.of());
        DocxException ex = assertThrows(DocxException.class, () -> reader.read(docx));
        assertTrue(ex.getMessage().contains("unsupported".toUpperCase()) || ex.getMessage().contains("Unsupported"));
    }

    private static String minimalParagraph(String text) {
        return wrapBody("<w:p><w:r><w:t>" + text + "</w:t></w:r></w:p>");
    }

    private static String wrapBody(String innerXml) {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<w:body>" + innerXml + "<w:sectPr/></w:body></w:document>";
    }

    private Path createMinimalDocx(String prefix, String documentXml, Map<String, byte[]> extraParts) throws IOException {
        Map<String, byte[]> parts = new HashMap<>();
        parts.put("[Content_Types].xml", minimalContentTypes());
        parts.put("_rels/.rels", minimalRootRels());
        parts.put("docProps/core.xml", minimalCoreProperties());
        parts.put("docProps/app.xml", minimalAppProperties());
        parts.put("word/_rels/document.xml.rels", emptyRelationships());
        parts.put("word/document.xml", documentXml.getBytes(StandardCharsets.UTF_8));
        parts.put("word/styles.xml", minimalStyles());
        parts.put("word/settings.xml", minimalSettings());
        extraParts.forEach(parts::put);

        Path temp = Files.createTempFile(prefix, ".docx");
        temp.toFile().deleteOnExit();
        try (ZipOutputStream out = new ZipOutputStream(Files.newOutputStream(temp))) {
            for (Map.Entry<String, byte[]> entry : parts.entrySet()) {
                ZipEntry zipEntry = new ZipEntry(entry.getKey());
                out.putNextEntry(zipEntry);
                out.write(entry.getValue());
                out.closeEntry();
            }
        }
        return temp;
    }

    private byte[] minimalContentTypes() {
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>" +
                "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
                "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>" +
                "<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>" +
                "</Types>";
        return xml.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] minimalRootRels() {
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
                "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
                "</Relationships>";
        return xml.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] minimalCoreProperties() {
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\">" +
                "<dc:title>Minimal</dc:title><dc:creator>Test</dc:creator><cp:revision>1</cp:revision>" +
                "</cp:coreProperties>";
        return xml.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] minimalAppProperties() {
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
                "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                "<Application>DocxReader</Application>" +
                "<DocSecurity>0</DocSecurity>" +
                "<ScaleCrop>false</ScaleCrop>" +
                "<HeadingPairs>" +
                "<vt:vector size=\"2\" baseType=\"variant\">" +
                "<vt:variant><vt:lpstr>Title</vt:lpstr></vt:variant>" +
                "<vt:variant><vt:i4>1</vt:i4></vt:variant>" +
                "</vt:vector>" +
                "</HeadingPairs>" +
                "<TitlesOfParts>" +
                "<vt:vector size=\"1\" baseType=\"lpstr\"><vt:lpstr>Document</vt:lpstr></vt:vector>" +
                "</TitlesOfParts>" +
                "<Company></Company>" +
                "<LinksUpToDate>false</LinksUpToDate>" +
                "<SharedDoc>false</SharedDoc>" +
                "<HyperlinksChanged>false</HyperlinksChanged>" +
                "<AppVersion>16.0000</AppVersion>" +
                "</Properties>";
        return xml.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] emptyRelationships() {
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>";
        return xml.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] minimalStyles() {
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">" +
                "<w:name w:val=\"Normal\"/>" +
                "</w:style>" +
                "<w:style w:type=\"character\" w:default=\"1\" w:styleId=\"DefaultParagraphFont\">" +
                "<w:name w:val=\"Default Paragraph Font\"/>" +
                "</w:style>" +
                "</w:styles>";
        return xml.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] minimalSettings() {
        String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                "<w:compat/>" +
                "</w:settings>";
        return xml.getBytes(StandardCharsets.UTF_8);
    }
}
