package com.example.docx.model;

import com.example.docx.model.document.WordDocument;
import com.example.docx.model.metadata.AppProperties;
import com.example.docx.model.metadata.CoreProperties;
import com.example.docx.model.metadata.CustomProperties;
import com.example.docx.model.notes.NoteCollection;
import com.example.docx.model.numbering.NumberingDefinitions;
import com.example.docx.model.relationship.RelationshipSet;
import com.example.docx.model.styles.StyleDefinitions;
import com.example.docx.model.support.ContentTypes;
import com.example.docx.model.support.CustomXmlPart;
import com.example.docx.model.support.FontTable;
import com.example.docx.model.support.MediaFile;
import com.example.docx.model.support.ChartPart;
import com.example.docx.model.support.Settings;
import com.example.docx.model.support.Theme;
import com.example.docx.model.support.WebSettings;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/**
 * In-memory representation of an unpacked DOCX package. Holds all parsed parts that have
 * dedicated model classes and exposes raw access to the remaining unused parts.
 */
public final class DocxPackage {

    private final CoreProperties coreProperties;
    private final AppProperties appProperties;
    private final CustomProperties customProperties;
    private final WordDocument document;
    private final StyleDefinitions styles;
    private final NumberingDefinitions numbering;
    private final NoteCollection footnotes;
    private final NoteCollection endnotes;
    private final FontTable fontTable;
    private final Settings settings;
    private final WebSettings webSettings;
    private final Theme theme;
    private final ContentTypes contentTypes;
    private final RelationshipSet packageRelationships;
    private final Map<String, RelationshipSet> partRelationships;
    private final Map<String, MediaFile> mediaFiles;
    private final Map<String, CustomXmlPart> customXmlParts;
    private final Map<String, ChartPart> charts;
    private final Map<String, byte[]> binaryParts;

    private DocxPackage(Builder builder) {
        this.coreProperties = builder.coreProperties;
        this.appProperties = builder.appProperties;
        this.customProperties = builder.customProperties;
        this.document = builder.document;
        this.styles = builder.styles;
        this.numbering = builder.numbering;
        this.footnotes = builder.footnotes;
        this.endnotes = builder.endnotes;
        this.fontTable = builder.fontTable;
        this.settings = builder.settings;
        this.webSettings = builder.webSettings;
        this.theme = builder.theme;
        this.contentTypes = builder.contentTypes;
        this.packageRelationships = builder.packageRelationships;
        this.partRelationships = Collections.unmodifiableMap(new LinkedHashMap<>(builder.partRelationships));
        this.mediaFiles = Collections.unmodifiableMap(new LinkedHashMap<>(builder.mediaFiles));
        this.customXmlParts = Collections.unmodifiableMap(new LinkedHashMap<>(builder.customXmlParts));
        this.charts = Collections.unmodifiableMap(new LinkedHashMap<>(builder.charts));
        this.binaryParts = Collections.unmodifiableMap(new LinkedHashMap<>(builder.binaryParts));
    }

    public Optional<CoreProperties> coreProperties() {
        return Optional.ofNullable(coreProperties);
    }

    public Optional<AppProperties> appProperties() {
        return Optional.ofNullable(appProperties);
    }

    public Optional<CustomProperties> customProperties() {
        return Optional.ofNullable(customProperties);
    }

    public Optional<WordDocument> document() {
        return Optional.ofNullable(document);
    }

    public Optional<StyleDefinitions> styles() {
        return Optional.ofNullable(styles);
    }

    public Optional<NumberingDefinitions> numbering() {
        return Optional.ofNullable(numbering);
    }

    public Optional<NoteCollection> footnotes() {
        return Optional.ofNullable(footnotes);
    }

    public Optional<NoteCollection> endnotes() {
        return Optional.ofNullable(endnotes);
    }

    public Optional<FontTable> fontTable() {
        return Optional.ofNullable(fontTable);
    }

    public Optional<Settings> settings() {
        return Optional.ofNullable(settings);
    }

    public Optional<WebSettings> webSettings() {
        return Optional.ofNullable(webSettings);
    }

    public Optional<Theme> theme() {
        return Optional.ofNullable(theme);
    }

    public Optional<ContentTypes> contentTypes() {
        return Optional.ofNullable(contentTypes);
    }

    public Optional<RelationshipSet> packageRelationships() {
        return Optional.ofNullable(packageRelationships);
    }

    public Map<String, RelationshipSet> relationshipsByPart() {
        return partRelationships;
    }

    public Map<String, MediaFile> mediaFiles() {
        return mediaFiles;
    }

    public Map<String, CustomXmlPart> customXmlParts() {
        return customXmlParts;
    }

    public Map<String, ChartPart> charts() {
        return charts;
    }

    public Map<String, byte[]> binaryParts() {
        return binaryParts;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static final class Builder {
        private CoreProperties coreProperties;
        private AppProperties appProperties;
        private CustomProperties customProperties;
        private WordDocument document;
        private StyleDefinitions styles;
        private NumberingDefinitions numbering;
        private NoteCollection footnotes;
        private NoteCollection endnotes;
        private FontTable fontTable;
        private Settings settings;
        private WebSettings webSettings;
        private Theme theme;
        private ContentTypes contentTypes;
        private RelationshipSet packageRelationships;
        private final Map<String, RelationshipSet> partRelationships = new LinkedHashMap<>();
        private final Map<String, MediaFile> mediaFiles = new LinkedHashMap<>();
        private final Map<String, CustomXmlPart> customXmlParts = new LinkedHashMap<>();
        private final Map<String, ChartPart> charts = new LinkedHashMap<>();
        private final Map<String, byte[]> binaryParts = new LinkedHashMap<>();

        public Builder coreProperties(CoreProperties value) {
            this.coreProperties = value;
            return this;
        }

        public Builder appProperties(AppProperties value) {
            this.appProperties = value;
            return this;
        }

        public Builder customProperties(CustomProperties value) {
            this.customProperties = value;
            return this;
        }

        public Builder document(WordDocument value) {
            this.document = value;
            return this;
        }

        public Builder styles(StyleDefinitions value) {
            this.styles = value;
            return this;
        }

        public Builder numbering(NumberingDefinitions value) {
            this.numbering = value;
            return this;
        }

        public Builder footnotes(NoteCollection value) {
            this.footnotes = value;
            return this;
        }

        public Builder endnotes(NoteCollection value) {
            this.endnotes = value;
            return this;
        }

        public Builder fontTable(FontTable value) {
            this.fontTable = value;
            return this;
        }

        public Builder settings(Settings value) {
            this.settings = value;
            return this;
        }

        public Builder webSettings(WebSettings value) {
            this.webSettings = value;
            return this;
        }

        public Builder theme(Theme value) {
            this.theme = value;
            return this;
        }

        public Builder contentTypes(ContentTypes value) {
            this.contentTypes = value;
            return this;
        }

        public Builder packageRelationships(RelationshipSet value) {
            this.packageRelationships = value;
            return this;
        }

        public Builder relationshipForPart(String partName, RelationshipSet value) {
            Objects.requireNonNull(partName, "partName");
            Objects.requireNonNull(value, "value");
            this.partRelationships.put(partName, value);
            return this;
        }

        public Builder mediaFile(MediaFile mediaFile) {
            Objects.requireNonNull(mediaFile, "mediaFile");
            this.mediaFiles.put(mediaFile.partName(), mediaFile);
            return this;
        }

        public Builder customXmlPart(CustomXmlPart part) {
            Objects.requireNonNull(part, "part");
            this.customXmlParts.put(part.partName(), part);
            return this;
        }

        public Builder chart(ChartPart chartPart) {
            Objects.requireNonNull(chartPart, "chartPart");
            this.charts.put(chartPart.partName(), chartPart);
            return this;
        }

        public Builder binaryPart(String partName, byte[] content) {
            Objects.requireNonNull(partName, "partName");
            Objects.requireNonNull(content, "content");
            this.binaryParts.put(partName, content.clone());
            return this;
        }

        public DocxPackage build() {
            return new DocxPackage(this);
        }
    }
}
