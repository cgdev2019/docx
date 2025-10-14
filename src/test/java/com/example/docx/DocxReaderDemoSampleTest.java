package com.example.docx;

import com.example.docx.model.DocxPackage;
import com.example.docx.model.document.WordDocument;
import com.example.docx.model.notes.NoteCollection;
import com.example.docx.model.relationship.RelationshipSet;
import com.example.docx.model.support.CustomXmlPart;
import com.example.docx.model.support.MediaFile;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import static org.junit.jupiter.api.Assertions.*;

class DocxReaderDemoSampleTest {

    private static final Path DEMO_DOCX = Path.of("samples", "demo.docx");
    private static final Path DEMO_DIRECTORY = Path.of("samples", "demo");

    private final DocxReader reader = new DocxReader();

    @Test
    void loadsAllStandardParts() {
        DocxPackage docx = reader.read(DEMO_DOCX);

        assertAll(
                () -> assertEquals("DOCX Demo", docx.coreProperties().orElseThrow().title()),
                () -> assertTrue(docx.appProperties().isPresent(), "app properties"),
                () -> assertTrue(docx.customProperties().isEmpty(), "custom properties"),
                () -> assertTrue(docx.document().isPresent(), "document"),
                () -> assertTrue(docx.styles().isPresent(), "styles"),
                () -> assertTrue(docx.numbering().isPresent(), "numbering"),
                () -> assertTrue(docx.footnotes().isPresent(), "footnotes"),
                () -> assertTrue(docx.endnotes().isPresent(), "endnotes"),
                () -> assertTrue(docx.fontTable().isPresent(), "font table"),
                () -> assertTrue(docx.settings().isPresent(), "settings"),
                () -> assertTrue(docx.webSettings().isPresent(), "web settings"),
                () -> assertTrue(docx.theme().isPresent(), "theme"),
                () -> assertTrue(docx.contentTypes().isPresent(), "content types"),
                () -> assertTrue(docx.packageRelationships().isPresent(), "package relationships"));

        WordDocument document = docx.document().orElseThrow();
        assertFalse(document.bodyElements().isEmpty(), "body elements");
        WordDocument.SectionProperties section = document.bodySectionProperties().orElseThrow();
        assertEquals(12240, section.pageDimensions().orElseThrow().widthTwips());
        assertEquals(15840, section.pageDimensions().orElseThrow().heightTwips());
        assertEquals(1440, section.pageMargins().orElseThrow().top());
    }

    @Test
    void loadsMediaCustomXmlAndBinaryParts() {
        DocxPackage docx = reader.read(DEMO_DOCX);

        assertEquals(
                Set.of(
                        "word/media/image1.gif",
                        "word/media/image2.png",
                        "word/media/image3.png",
                        "word/media/image4.png"
                ),
                docx.mediaFiles().keySet(),
                "media parts");

        MediaFile gif = docx.mediaFiles().get("word/media/image1.gif");
        assertNotNull(gif, "gif media file");
        assertTrue(gif.data().length > 0, "gif payload");

        assertEquals(
                Set.of(
                        "customXml/item1.xml",
                        "customXml/item2.xml",
                        "customXml/itemProps1.xml",
                        "customXml/itemProps2.xml"
                ),
                docx.customXmlParts().keySet(),
                "custom XML parts");

        for (CustomXmlPart part : docx.customXmlParts().values()) {
            part.document().ifPresent(doc -> assertNotNull(doc.getDocumentElement(), part.partName()));
        }

        assertEquals(
                Set.of(
                        "word/fonts/font1.odttf",
                        "word/fonts/font2.odttf",
                        "word/fonts/font3.odttf",
                        "word/fonts/font4.odttf",
                        "word/fonts/font5.odttf",
                        "word/fonts/font6.odttf"
                ),
                docx.binaryParts().keySet(),
                "binary parts");

        docx.binaryParts().values().forEach(bytes -> assertTrue(bytes.length > 0));
    }

    @Test
    void footnoteAndEndnoteContentsAreParsed() {
        DocxPackage docx = reader.read(DEMO_DOCX);

        NoteCollection footnotes = docx.footnotes().orElseThrow();
        NoteCollection.Note footnote = footnotes.noteById(2).orElseThrow();
        String footnoteText = extractText(footnote.content());
        assertTrue(footnoteText.contains("paged media"));

        NoteCollection endnotes = docx.endnotes().orElseThrow();
        NoteCollection.Note endnote = endnotes.noteById(2).orElseThrow();
        String endnoteText = extractText(endnote.content());
        assertTrue(endnoteText.contains("Endnotes are typically"));
    }

    @Test
    void documentRelationshipsIncludeKeyParts() {
        DocxPackage docx = reader.read(DEMO_DOCX);
        RelationshipSet relationships = docx.relationshipsByPart().get("word/document.xml");
        assertNotNull(relationships, "document relationships");

        Map<String, RelationshipSet.Relationship> byId = relationships.relationships();
        assertTrue(byId.containsKey("rId4"));
        assertTrue(byId.get("rId4").type().endsWith("/styles"));
        assertEquals("styles.xml", byId.get("rId4").target());
        assertTrue(byId.get("rId7").type().endsWith("/footnotes"));
        assertEquals("footnotes.xml", byId.get("rId7").target());
        assertTrue(byId.get("rId9").type().endsWith("/hyperlink"));
        assertEquals("http://calibre-ebook.com/download", byId.get("rId9").target());

        RelationshipSet packageRels = docx.packageRelationships().orElseThrow();
        Set<String> packageTargets = packageRels.relationships().values().stream()
                .map(RelationshipSet.Relationship::target)
                .collect(Collectors.toSet());
        assertEquals(Set.of("word/document.xml", "docProps/core.xml", "docProps/app.xml"), packageTargets);
    }

    @Test
    void readingUnzippedPackageMatchesZipArchive() {
        DocxPackage zipped = reader.read(DEMO_DOCX);
        DocxPackage directory = reader.readDirectory(DEMO_DIRECTORY);

        assertEquals(zipped.mediaFiles().keySet(), directory.mediaFiles().keySet(), "media equality");
        assertEquals(zipped.customXmlParts().keySet(), directory.customXmlParts().keySet(), "custom XML equality");
        assertEquals(zipped.binaryParts().keySet(), directory.binaryParts().keySet(), "binary equality");
        assertEquals(
                zipped.document().orElseThrow().bodyElements().size(),
                directory.document().orElseThrow().bodyElements().size(),
                "block count");
    }

    private static String extractText(List<WordDocument.Block> blocks) {
        StringBuilder buffer = new StringBuilder();
        for (WordDocument.Block block : blocks) {
            if (block instanceof WordDocument.Paragraph paragraph) {
                for (WordDocument.ParagraphContent content : paragraph.content()) {
                    if (content instanceof WordDocument.Run run) {
                        for (WordDocument.Inline inline : run.elements()) {
                            if (inline instanceof WordDocument.Text text) {
                                buffer.append(text.text());
                            }
                        }
                    }
                }
            }
        }
        return buffer.toString().replace('\n', ' ').trim().replaceAll("\\s+", " ");
    }
}




