package com.example.docx;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.DocxPackage;
import com.example.docx.model.document.WordDocument;
import com.example.docx.model.metadata.AppProperties;
import com.example.docx.model.metadata.CoreProperties;
import com.example.docx.model.metadata.CustomProperties;
import com.example.docx.model.notes.NoteCollection;
import com.example.docx.model.numbering.NumberingDefinitions;
import com.example.docx.model.relationship.RelationshipSet;
import com.example.docx.model.styles.StyleDefinitions;
import com.example.docx.model.support.ContentTypes;
import com.example.docx.model.support.FontTable;
import com.example.docx.model.support.Settings;
import com.example.docx.model.support.Theme;
import com.example.docx.model.support.WebSettings;
import com.example.docx.model.support.ChartPart;
import com.example.docx.parser.ContentTypesParser;
import com.example.docx.parser.CustomXmlLoader;
import com.example.docx.parser.FontTableParser;
import com.example.docx.parser.MainDocumentParser;
import com.example.docx.parser.MetadataParser;
import com.example.docx.parser.NotesParser;
import com.example.docx.parser.NumberingParser;
import com.example.docx.parser.ParsingContext;
import com.example.docx.parser.RelationshipsParser;
import com.example.docx.parser.SettingsParser;
import com.example.docx.parser.StylesParser;
import com.example.docx.parser.ThemeParser;
import com.example.docx.parser.XmlUtils;
import com.example.docx.parser.binary.BinaryPartLoader;
import com.example.docx.parser.binary.MediaLoader;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Set;

/**
 * Entry point for reading DOCX packages.
 */
public final class DocxReader {

    private static final Set<String> ALLOWED_PARTS = Set.of(
            "[Content_Types].xml",
            "_rels/.rels",
            "docProps/app.xml",
            "docProps/core.xml",
            "docProps/custom.xml",
            "word/document.xml",
            "word/styles.xml",
            "word/numbering.xml",
            "word/fontTable.xml",
            "word/settings.xml",
            "word/webSettings.xml",
            "word/theme/theme1.xml",
            "word/footnotes.xml",
            "word/endnotes.xml"
    );

    private static final List<String> ALLOWED_PREFIXES = List.of(
            "word/_rels/",
            "word/media/",
            "word/charts/",
            "word/charts/_rels/",
            "word/fonts/",
            "word/embeddings/",
            "word/ink/",
            "word/theme/",
            "customXml/",
            "customXml/_rels/"
    );

    private final ParsingContext parsingContext = new ParsingContext();

    private final ContentTypesParser contentTypesParser = new ContentTypesParser();
    private final RelationshipsParser relationshipsParser = new RelationshipsParser();
    private final MetadataParser metadataParser = new MetadataParser();
    private final MainDocumentParser mainDocumentParser = new MainDocumentParser(parsingContext);
    private final StylesParser stylesParser = new StylesParser(parsingContext);
    private final CustomXmlLoader customXmlLoader = new CustomXmlLoader();
    private final NumberingParser numberingParser = new NumberingParser();
    private final NotesParser notesParser = new NotesParser(parsingContext);
    private final FontTableParser fontTableParser = new FontTableParser();
    private final SettingsParser settingsParser = new SettingsParser();
    private final ThemeParser themeParser = new ThemeParser();
    private final MediaLoader mediaLoader = new MediaLoader();
    private final BinaryPartLoader binaryPartLoader = new BinaryPartLoader();

    public DocxPackage read(Path path) {
        try (DocxArchive archive = DocxArchive.open(path)) {
            return readInternal(archive);
        } catch (IOException e) {
            throw new DocxException("Unable to read DOCX package: " + path, e);
        }
    }

    public DocxPackage readDirectory(Path directory) {
        if (!Files.isDirectory(directory)) {
            throw new IllegalArgumentException("Path is not a directory: " + directory);
        }
        try (DocxArchive archive = DocxArchive.open(directory)) {
            return readInternal(archive);
        } catch (IOException e) {
            throw new DocxException("Unable to read DOCX directory: " + directory, e);
        }
    }

    private DocxPackage readInternal(DocxArchive archive) throws IOException {
        validatePackageParts(archive);

        DocxPackage.Builder builder = DocxPackage.builder();

        ContentTypes contentTypes = contentTypesParser.parse(archive);
        if (contentTypes != null) {
            builder.contentTypes(contentTypes);
        }

        RelationshipSet packageRelationships = relationshipsParser.parse(archive, "_rels/.rels");
        if (packageRelationships != null) {
            builder.packageRelationships(packageRelationships);
        }

        metadataParser.parse(archive, builder);

        RelationshipSet documentRelationships = relationshipsParser.parse(archive, "word/_rels/document.xml.rels");
        if (documentRelationships != null) {
            builder.relationshipForPart("word/document.xml", documentRelationships);
        }

        WordDocument document = mainDocumentParser.parse(archive, documentRelationships);
        if (document != null) {
            builder.document(document);
        }

        StyleDefinitions styles = stylesParser.parse(archive, builder);
        if (styles != null) {
            builder.styles(styles);
        }

        NumberingDefinitions numbering = numberingParser.parse(archive);
        if (numbering != null) {
            builder.numbering(numbering);
        }

        NoteCollection footnotes = notesParser.parse(archive, "word/footnotes.xml", NoteCollection.Type.FOOTNOTE, builder);
        if (footnotes != null) {
            builder.footnotes(footnotes);
        }

        NoteCollection endnotes = notesParser.parse(archive, "word/endnotes.xml", NoteCollection.Type.ENDNOTE, builder);
        if (endnotes != null) {
            builder.endnotes(endnotes);
        }

        FontTable fontTable = fontTableParser.parse(archive);
        if (fontTable != null) {
            builder.fontTable(fontTable);
        }

        Settings settings = settingsParser.parseSettings(archive);
        if (settings != null) {
            builder.settings(settings);
        }

        WebSettings webSettings = settingsParser.parseWebSettings(archive);
        if (webSettings != null) {
            builder.webSettings(webSettings);
        }

        Theme theme = themeParser.parse(archive);
        if (theme != null) {
            builder.theme(theme);
        }

        loadCharts(archive, builder);
        customXmlLoader.load(archive, builder);
        mediaLoader.loadMedia(archive, contentTypes, builder);
        binaryPartLoader.load(archive, builder);

        return builder.build();
    }

    private void validatePackageParts(DocxArchive archive) throws IOException {
        for (String part : archive.list("")) {
            if (!isAllowedPart(part)) {
                throw new DocxException("Unsupported part detected: " + part);
            }
        }
    }

    private boolean isAllowedPart(String part) {
        if (ALLOWED_PARTS.contains(part)) {
            return true;
        }
        for (String prefix : ALLOWED_PREFIXES) {
            if (part.startsWith(prefix)) {
                return true;
            }
        }
        return false;
    }

    private void loadCharts(DocxArchive archive, DocxPackage.Builder builder) throws IOException {
        for (String entry : archive.list("word/charts")) {
            if (entry.contains("/_rels/")) {
                continue;
            }
            if (!entry.endsWith(".xml")) {
                throw new DocxException("Unsupported chart resource: " + entry);
            }
            try (var stream = archive.open(entry)) {
                builder.chart(new ChartPart(entry, XmlUtils.parse(stream)));
            }
            RelationshipSet relationships = relationshipsParser.parse(archive, toRelsPath(entry));
            if (relationships != null) {
                builder.relationshipForPart(entry, relationships);
            }
        }
    }

    private String toRelsPath(String partName) {
        int index = partName.lastIndexOf('/');
        String directory = index == -1 ? "" : partName.substring(0, index + 1);
        String file = index == -1 ? partName : partName.substring(index + 1);
        return directory + "_rels/" + file + ".rels";
    }

}
