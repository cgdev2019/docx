package com.example.docx.model.document;

import org.w3c.dom.Element;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

public final class WordDocument {

    private final List<Block> bodyElements;
    private final SectionProperties bodySectionProperties;

    private WordDocument(List<Block> bodyElements, SectionProperties bodySectionProperties) {
        this.bodyElements = List.copyOf(bodyElements);
        this.bodySectionProperties = bodySectionProperties;
    }

    public List<Block> bodyElements() {
        return bodyElements;
    }

    public Optional<SectionProperties> bodySectionProperties() {
        return Optional.ofNullable(bodySectionProperties);
    }

    public static Builder builder() {
        return new Builder();
    }

    public static final class Builder {
        private final List<Block> blocks = new ArrayList<>();
        private SectionProperties sectionProperties;

        private Builder() {
        }

        public Builder addBlock(Block block) {
            Objects.requireNonNull(block, "block");
            blocks.add(block);
            return this;
        }

        public Builder sectionProperties(SectionProperties sectionProperties) {
            this.sectionProperties = sectionProperties;
            return this;
        }

        public WordDocument build() {
            return new WordDocument(blocks, sectionProperties);
        }
    }

    public sealed interface Block permits Paragraph, Table, StructuredDocumentTag, SectionBreak, Bookmark {
    }

    public static final class Paragraph implements Block {
        private final ParagraphProperties properties;
        private final List<ParagraphContent> content;

        public Paragraph(ParagraphProperties properties, List<ParagraphContent> content) {
            this.properties = Objects.requireNonNull(properties, "properties");
            this.content = List.copyOf(content);
        }

        public ParagraphProperties properties() {
            return properties;
        }

        public List<ParagraphContent> content() {
            return content;
        }
    }

    public static final class Table implements Block {
        private final TableProperties properties;
        private final List<TableRow> rows;

        public Table(TableProperties properties, List<TableRow> rows) {
            this.properties = Objects.requireNonNull(properties, "properties");
            this.rows = List.copyOf(rows);
        }

        public TableProperties properties() {
            return properties;
        }

        public List<TableRow> rows() {
            return rows;
        }
    }

    public static final class StructuredDocumentTag implements Block {
        private final SdtProperties properties;
        private final List<Block> content;

        public StructuredDocumentTag(SdtProperties properties, List<Block> content) {
            this.properties = Objects.requireNonNull(properties, "properties");
            this.content = List.copyOf(content);
        }

        public SdtProperties properties() {
            return properties;
        }

        public List<Block> content() {
            return content;
        }
    }

    public static final class SectionBreak implements Block {
        private final SectionProperties sectionProperties;

        public SectionBreak(SectionProperties sectionProperties) {
            this.sectionProperties = Objects.requireNonNull(sectionProperties, "sectionProperties");
        }

        public SectionProperties sectionProperties() {
            return sectionProperties;
        }
    }

    public static final class Bookmark implements Block {
        public enum Kind {
            START,
            END
        }

        private final Kind kind;
        private final String id;
        private final String name;

        public Bookmark(Kind kind, String id, String name) {
            this.kind = Objects.requireNonNull(kind, "kind");
            this.id = id;
            this.name = name;
        }

        public Kind kind() {
            return kind;
        }

        public Optional<String> id() {
            return Optional.ofNullable(id);
        }

        public Optional<String> name() {
            return Optional.ofNullable(name);
        }
    }

    public sealed interface ParagraphContent permits Run, Hyperlink, BookmarkStart, BookmarkEnd, Field, StructuredDocumentTagRun {
    }

    public static final class Run implements ParagraphContent {
        private final RunProperties properties;
        private final List<Inline> elements;

        public Run(RunProperties properties, List<Inline> elements) {
            this.properties = Objects.requireNonNull(properties, "properties");
            this.elements = List.copyOf(elements);
        }

        public RunProperties properties() {
            return properties;
        }

        public List<Inline> elements() {
            return elements;
        }
    }

    public static final class Hyperlink implements ParagraphContent {
        private final String relationshipId;
        private final String anchor;
        private final List<Run> runs;

        public Hyperlink(String relationshipId, String anchor, List<Run> runs) {
            this.relationshipId = relationshipId;
            this.anchor = anchor;
            this.runs = List.copyOf(runs);
        }

        public Optional<String> relationshipId() {
            return Optional.ofNullable(relationshipId);
        }

        public Optional<String> anchor() {
            return Optional.ofNullable(anchor);
        }

        public List<Run> runs() {
            return runs;
        }
    }

    public static final class BookmarkStart implements ParagraphContent {
        private final String id;
        private final String name;

        public BookmarkStart(String id, String name) {
            this.id = id;
            this.name = name;
        }

        public Optional<String> id() {
            return Optional.ofNullable(id);
        }

        public Optional<String> name() {
            return Optional.ofNullable(name);
        }
    }

    public static final class BookmarkEnd implements ParagraphContent {
        private final String id;

        public BookmarkEnd(String id) {
            this.id = id;
        }

        public Optional<String> id() {
            return Optional.ofNullable(id);
        }
    }

    public static final class Field implements ParagraphContent {
        private final List<Run> instructionRuns;
        private final List<Run> resultRuns;

        public Field(List<Run> instructionRuns, List<Run> resultRuns) {
            this.instructionRuns = List.copyOf(instructionRuns);
            this.resultRuns = List.copyOf(resultRuns);
        }

        public List<Run> instructionRuns() {
            return instructionRuns;
        }

        public List<Run> resultRuns() {
            return resultRuns;
        }
    }

    public static final class StructuredDocumentTagRun implements ParagraphContent {
        private final SdtProperties properties;
        private final List<ParagraphContent> content;

        public StructuredDocumentTagRun(SdtProperties properties, List<ParagraphContent> content) {
            this.properties = Objects.requireNonNull(properties, "properties");
            this.content = List.copyOf(content);
        }

        public SdtProperties properties() {
            return properties;
        }

        public List<ParagraphContent> content() {
            return content;
        }
    }

    public sealed interface Inline permits Text, Break, Tab, Drawing, FootnoteReference, EndnoteReference,
            CommentReference, FieldCharacter, FieldInstruction, Symbol, SoftHyphen, NoBreakHyphen, Separator, ReferenceMark {
    }

    public static final class Text implements Inline {
        private final String text;
        private final boolean preserveSpace;

        public Text(String text, boolean preserveSpace) {
            this.text = Objects.requireNonNull(text, "text");
            this.preserveSpace = preserveSpace;
        }

        public String text() {
            return text;
        }

        public boolean preserveSpace() {
            return preserveSpace;
        }
    }

    public static final class Break implements Inline {
        public enum Type {
            LINE,
            PAGE,
            COLUMN
        }

        private final Type type;
        private final String clear;

        public Break(Type type, String clear) {
            this.type = Objects.requireNonNull(type, "type");
            this.clear = clear;
        }

        public Type type() {
            return type;
        }

        public Optional<String> clear() {
            return Optional.ofNullable(clear);
        }
    }

    public static final class Tab implements Inline {
    }

    public static final class Drawing implements Inline {
        private final String relationshipId;
        private final String description;
        private final long width;
        private final long height;
        private final boolean inline;

        public Drawing(String relationshipId, String description, long width, long height, boolean inline) {
            this.relationshipId = relationshipId;
            this.description = description;
            this.width = width;
            this.height = height;
            this.inline = inline;
        }

        public Optional<String> relationshipId() {
            return Optional.ofNullable(relationshipId);
        }

        public Optional<String> description() {
            return Optional.ofNullable(description);
        }

        public long width() {
            return width;
        }

        public long height() {
            return height;
        }

        public boolean inline() {
            return inline;
        }
    }

    public static final class FootnoteReference implements Inline {
        private final int id;

        public FootnoteReference(int id) {
            this.id = id;
        }

        public int id() {
            return id;
        }
    }

    public static final class EndnoteReference implements Inline {
        private final int id;

        public EndnoteReference(int id) {
            this.id = id;
        }

        public int id() {
            return id;
        }
    }

    public static final class CommentReference implements Inline {
        private final int id;

        public CommentReference(int id) {
            this.id = id;
        }

        public int id() {
            return id;
        }
    }

    public static final class FieldCharacter implements Inline {
        public enum CharacterType {
            BEGIN,
            SEPARATE,
            END
        }

        private final CharacterType type;

        public FieldCharacter(CharacterType type) {
            this.type = Objects.requireNonNull(type, "type");
        }

        public CharacterType type() {
            return type;
        }
    }

    public static final class FieldInstruction implements Inline {
        private final String instruction;

        public FieldInstruction(String instruction) {
            this.instruction = Objects.requireNonNull(instruction, "instruction");
        }

        public String instruction() {
            return instruction;
        }
    }

    public static final class Symbol implements Inline {
        private final String font;
        private final String charCode;

        public Symbol(String font, String charCode) {
            this.font = font;
            this.charCode = Objects.requireNonNull(charCode, "charCode");
        }

        public Optional<String> font() {
            return Optional.ofNullable(font);
        }

        public String charCode() {
            return charCode;
        }
    }

    public static final class SoftHyphen implements Inline {
    }

    public static final class NoBreakHyphen implements Inline {
    }

    public static final class Separator implements Inline {
        public enum Kind { FOOTNOTE, CONTINUATION }

        private final Kind kind;

        public Separator(Kind kind) {
            this.kind = Objects.requireNonNull(kind, "kind");
        }

        public Kind kind() {
            return kind;
        }
    }

    public static final class ReferenceMark implements Inline {
        public enum Kind { FOOTNOTE, ENDNOTE }

        private final Kind kind;

        public ReferenceMark(Kind kind) {
            this.kind = Objects.requireNonNull(kind, "kind");
        }

        public Kind kind() {
            return kind;
        }
    }


    public static final class RunProperties {
        private final String styleId;
        private final boolean bold;
        private final boolean italic;
        private final boolean underline;
        private final String underlineType;
        private final boolean strike;
        private final boolean doubleStrike;
        private final boolean smallCaps;
        private final boolean allCaps;
        private final boolean vanish;
        private final String color;
        private final String highlight;
        private final String verticalAlignment;
        private final Integer size;
        private final Integer complexScriptSize;
        private final Map<String, String> fonts;
        private final Element rawProperties;

        public RunProperties(String styleId,
                             boolean bold,
                             boolean italic,
                             boolean underline,
                             String underlineType,
                             boolean strike,
                             boolean doubleStrike,
                             boolean smallCaps,
                             boolean allCaps,
                             boolean vanish,
                             String color,
                             String highlight,
                             String verticalAlignment,
                             Integer size,
                             Integer complexScriptSize,
                             Map<String, String> fonts,
                             Element rawProperties) {
            this.styleId = styleId;
            this.bold = bold;
            this.italic = italic;
            this.underline = underline;
            this.underlineType = underlineType;
            this.strike = strike;
            this.doubleStrike = doubleStrike;
            this.smallCaps = smallCaps;
            this.allCaps = allCaps;
            this.vanish = vanish;
            this.color = color;
            this.highlight = highlight;
            this.verticalAlignment = verticalAlignment;
            this.size = size;
            this.complexScriptSize = complexScriptSize;
            this.fonts = fonts == null ? Map.of() : Collections.unmodifiableMap(new LinkedHashMap<>(fonts));
            this.rawProperties = rawProperties;
        }

        public Optional<String> styleId() {
            return Optional.ofNullable(styleId);
        }

        public boolean bold() {
            return bold;
        }

        public boolean italic() {
            return italic;
        }

        public boolean underline() {
            return underline;
        }

        public Optional<String> underlineType() {
            return Optional.ofNullable(underlineType);
        }

        public boolean strike() {
            return strike;
        }

        public boolean doubleStrike() {
            return doubleStrike;
        }

        public boolean smallCaps() {
            return smallCaps;
        }

        public boolean allCaps() {
            return allCaps;
        }

        public boolean vanish() {
            return vanish;
        }

        public Optional<String> color() {
            return Optional.ofNullable(color);
        }

        public Optional<String> highlight() {
            return Optional.ofNullable(highlight);
        }

        public Optional<String> verticalAlignment() {
            return Optional.ofNullable(verticalAlignment);
        }

        public Optional<Integer> size() {
            return Optional.ofNullable(size);
        }

        public Optional<Integer> complexScriptSize() {
            return Optional.ofNullable(complexScriptSize);
        }

        public Map<String, String> fonts() {
            return fonts;
        }

        public Optional<Element> rawProperties() {
            return Optional.ofNullable(rawProperties);
        }
    }

    public static final class ParagraphProperties {
        private final String styleId;
        private final NumberingReference numbering;
        private final Alignment alignment;
        private final Indentation indentation;
        private final Spacing spacing;
        private final Integer outlineLevel;
        private final boolean keepTogether;
        private final boolean keepWithNext;
        private final boolean pageBreakBefore;
        private final List<TabStop> tabs;
        private final Element rawProperties;

        public ParagraphProperties(String styleId,
                                   NumberingReference numbering,
                                   Alignment alignment,
                                   Indentation indentation,
                                   Spacing spacing,
                                   Integer outlineLevel,
                                   boolean keepTogether,
                                   boolean keepWithNext,
                                   boolean pageBreakBefore,
                                   List<TabStop> tabs,
                                   Element rawProperties) {
            this.styleId = styleId;
            this.numbering = numbering;
            this.alignment = alignment;
            this.indentation = indentation;
            this.spacing = spacing;
            this.outlineLevel = outlineLevel;
            this.keepTogether = keepTogether;
            this.keepWithNext = keepWithNext;
            this.pageBreakBefore = pageBreakBefore;
            this.tabs = tabs == null ? List.of() : List.copyOf(tabs);
            this.rawProperties = rawProperties;
        }

        public Optional<String> styleId() {
            return Optional.ofNullable(styleId);
        }

        public Optional<NumberingReference> numbering() {
            return Optional.ofNullable(numbering);
        }

        public Optional<Alignment> alignment() {
            return Optional.ofNullable(alignment);
        }

        public Optional<Indentation> indentation() {
            return Optional.ofNullable(indentation);
        }

        public Optional<Spacing> spacing() {
            return Optional.ofNullable(spacing);
        }

        public Optional<Integer> outlineLevel() {
            return Optional.ofNullable(outlineLevel);
        }

        public boolean keepTogether() {
            return keepTogether;
        }

        public boolean keepWithNext() {
            return keepWithNext;
        }

        public boolean pageBreakBefore() {
            return pageBreakBefore;
        }

        public List<TabStop> tabs() {
            return tabs;
        }

        public Optional<Element> rawProperties() {
            return Optional.ofNullable(rawProperties);
        }
    }

    public static final class NumberingReference {
        private final int numberingId;
        private final int level;

        public NumberingReference(int numberingId, int level) {
            this.numberingId = numberingId;
            this.level = level;
        }

        public int numberingId() {
            return numberingId;
        }

        public int level() {
            return level;
        }
    }

    public enum Alignment {
        LEFT,
        CENTER,
        RIGHT,
        JUSTIFIED,
        DISTRIBUTE,
        THAI_DISTRIBUTED,
        JUSTIFY_LOW
    }

    public static final class Indentation {
        private final Integer left;
        private final Integer right;
        private final Integer firstLine;
        private final Integer hanging;

        public Indentation(Integer left, Integer right, Integer firstLine, Integer hanging) {
            this.left = left;
            this.right = right;
            this.firstLine = firstLine;
            this.hanging = hanging;
        }

        public Optional<Integer> left() {
            return Optional.ofNullable(left);
        }

        public Optional<Integer> right() {
            return Optional.ofNullable(right);
        }

        public Optional<Integer> firstLine() {
            return Optional.ofNullable(firstLine);
        }

        public Optional<Integer> hanging() {
            return Optional.ofNullable(hanging);
        }
    }

    public static final class Spacing {
        private final Integer before;
        private final Integer after;
        private final Integer line;
        private final String rule;

        public Spacing(Integer before, Integer after, Integer line, String rule) {
            this.before = before;
            this.after = after;
            this.line = line;
            this.rule = rule;
        }

        public Optional<Integer> before() {
            return Optional.ofNullable(before);
        }

        public Optional<Integer> after() {
            return Optional.ofNullable(after);
        }

        public Optional<Integer> line() {
            return Optional.ofNullable(line);
        }

        public Optional<String> rule() {
            return Optional.ofNullable(rule);
        }
    }

    public static final class TabStop {
        private final String alignment;
        private final Integer position;

        public TabStop(String alignment, Integer position) {
            this.alignment = alignment;
            this.position = position;
        }

        public Optional<String> alignment() {
            return Optional.ofNullable(alignment);
        }

        public Optional<Integer> position() {
            return Optional.ofNullable(position);
        }
    }

    public static final class TableRow {
        private final TableRowProperties properties;
        private final List<TableCell> cells;

        public TableRow(TableRowProperties properties, List<TableCell> cells) {
            this.properties = Objects.requireNonNull(properties, "properties");
            this.cells = List.copyOf(cells);
        }

        public TableRowProperties properties() {
            return properties;
        }

        public List<TableCell> cells() {
            return cells;
        }
    }

    public static final class TableCell {
        private final TableCellProperties properties;
        private final List<Block> content;

        public TableCell(TableCellProperties properties, List<Block> content) {
            this.properties = Objects.requireNonNull(properties, "properties");
            this.content = List.copyOf(content);
        }

        public TableCellProperties properties() {
            return properties;
        }

        public List<Block> content() {
            return content;
        }
    }

    public static final class TableProperties {
        private final String styleId;
        private final Integer width;
        private final String widthType;
        private final Integer look;
        private final Element rawProperties;

        public TableProperties(String styleId, Integer width, String widthType, Integer look, Element rawProperties) {
            this.styleId = styleId;
            this.width = width;
            this.widthType = widthType;
            this.look = look;
            this.rawProperties = rawProperties;
        }

        public Optional<String> styleId() {
            return Optional.ofNullable(styleId);
        }

        public Optional<Integer> width() {
            return Optional.ofNullable(width);
        }

        public Optional<String> widthType() {
            return Optional.ofNullable(widthType);
        }

        public Optional<Integer> look() {
            return Optional.ofNullable(look);
        }

        public Optional<Element> rawProperties() {
            return Optional.ofNullable(rawProperties);
        }
    }

    public static final class TableRowProperties {
        private final boolean cantSplit;
        private final Integer gridAfter;
        private final Integer gridBefore;
        private final Integer heightTwips;
        private final Element rawProperties;

        public TableRowProperties(boolean cantSplit,
                                  Integer gridAfter,
                                  Integer gridBefore,
                                  Integer heightTwips,
                                  Element rawProperties) {
            this.cantSplit = cantSplit;
            this.gridAfter = gridAfter;
            this.gridBefore = gridBefore;
            this.heightTwips = heightTwips;
            this.rawProperties = rawProperties;
        }

        public boolean cantSplit() {
            return cantSplit;
        }

        public Optional<Integer> gridAfter() {
            return Optional.ofNullable(gridAfter);
        }

        public Optional<Integer> gridBefore() {
            return Optional.ofNullable(gridBefore);
        }

        public Optional<Integer> heightTwips() {
            return Optional.ofNullable(heightTwips);
        }

        public Optional<Element> rawProperties() {
            return Optional.ofNullable(rawProperties);
        }
    }

    public static final class TableCellProperties {
        private final Integer gridSpan;
        private final Integer width;
        private final String widthType;
        private final String verticalAlignment;
        private final boolean verticalMerge;
        private final Element rawProperties;

        public TableCellProperties(Integer gridSpan,
                                   Integer width,
                                   String widthType,
                                   String verticalAlignment,
                                   boolean verticalMerge,
                                   Element rawProperties) {
            this.gridSpan = gridSpan;
            this.width = width;
            this.widthType = widthType;
            this.verticalAlignment = verticalAlignment;
            this.verticalMerge = verticalMerge;
            this.rawProperties = rawProperties;
        }

        public Optional<Integer> gridSpan() {
            return Optional.ofNullable(gridSpan);
        }

        public Optional<Integer> width() {
            return Optional.ofNullable(width);
        }

        public Optional<String> widthType() {
            return Optional.ofNullable(widthType);
        }

        public Optional<String> verticalAlignment() {
            return Optional.ofNullable(verticalAlignment);
        }

        public boolean verticalMerge() {
            return verticalMerge;
        }

        public Optional<Element> rawProperties() {
            return Optional.ofNullable(rawProperties);
        }
    }

    public static final class SectionProperties {
        private final PageDimensions pageDimensions;
        private final PageMargins pageMargins;
        private final HeaderFooterReference headerFooterReference;
        private final String sectionType;
        private final Element rawProperties;

        public SectionProperties(PageDimensions pageDimensions,
                                 PageMargins pageMargins,
                                 HeaderFooterReference headerFooterReference,
                                 String sectionType,
                                 Element rawProperties) {
            this.pageDimensions = pageDimensions;
            this.pageMargins = pageMargins;
            this.headerFooterReference = headerFooterReference;
            this.sectionType = sectionType;
            this.rawProperties = rawProperties;
        }

        public Optional<PageDimensions> pageDimensions() {
            return Optional.ofNullable(pageDimensions);
        }

        public Optional<PageMargins> pageMargins() {
            return Optional.ofNullable(pageMargins);
        }

        public Optional<HeaderFooterReference> headerFooterReference() {
            return Optional.ofNullable(headerFooterReference);
        }

        public Optional<String> sectionType() {
            return Optional.ofNullable(sectionType);
        }

        public Optional<Element> rawProperties() {
            return Optional.ofNullable(rawProperties);
        }
    }

    public static final class PageDimensions {
        private final int widthTwips;
        private final int heightTwips;
        private final String orientation;

        public PageDimensions(int widthTwips, int heightTwips, String orientation) {
            this.widthTwips = widthTwips;
            this.heightTwips = heightTwips;
            this.orientation = orientation;
        }

        public int widthTwips() {
            return widthTwips;
        }

        public int heightTwips() {
            return heightTwips;
        }

        public Optional<String> orientation() {
            return Optional.ofNullable(orientation);
        }
    }

    public static final class PageMargins {
        private final int top;
        private final int right;
        private final int bottom;
        private final int left;
        private final int header;
        private final int footer;
        private final int gutter;

        public PageMargins(int top,
                           int right,
                           int bottom,
                           int left,
                           int header,
                           int footer,
                           int gutter) {
            this.top = top;
            this.right = right;
            this.bottom = bottom;
            this.left = left;
            this.header = header;
            this.footer = footer;
            this.gutter = gutter;
        }

        public int top() {
            return top;
        }

        public int right() {
            return right;
        }

        public int bottom() {
            return bottom;
        }

        public int left() {
            return left;
        }

        public int header() {
            return header;
        }

        public int footer() {
            return footer;
        }

        public int gutter() {
            return gutter;
        }
    }

    public static final class HeaderFooterReference {
        private final Map<String, String> headerReferences;
        private final Map<String, String> footerReferences;

        public HeaderFooterReference(Map<String, String> headerReferences, Map<String, String> footerReferences) {
            this.headerReferences = headerReferences == null
                    ? Map.of()
                    : Collections.unmodifiableMap(new LinkedHashMap<>(headerReferences));
            this.footerReferences = footerReferences == null
                    ? Map.of()
                    : Collections.unmodifiableMap(new LinkedHashMap<>(footerReferences));
        }

        public Map<String, String> headerReferences() {
            return headerReferences;
        }

        public Map<String, String> footerReferences() {
            return footerReferences;
        }
    }

    public static final class SdtProperties {
        private final String tag;
        private final String alias;
        private final String id;
        private final Element rawProperties;

        public SdtProperties(String tag, String alias, String id, Element rawProperties) {
            this.tag = tag;
            this.alias = alias;
            this.id = id;
            this.rawProperties = rawProperties;
        }

        public Optional<String> tag() {
            return Optional.ofNullable(tag);
        }

        public Optional<String> alias() {
            return Optional.ofNullable(alias);
        }

        public Optional<String> id() {
            return Optional.ofNullable(id);
        }

        public Optional<Element> rawProperties() {
            return Optional.ofNullable(rawProperties);
        }
    }
}

