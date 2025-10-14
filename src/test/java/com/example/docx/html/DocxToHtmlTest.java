package com.example.docx.html;
import com.example.docx.DocxReader;
import com.example.docx.model.DocxPackage;
import com.example.docx.model.document.WordDocument;
import com.example.docx.model.support.Theme;
import com.example.docx.parser.Namespaces;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import javax.xml.parsers.DocumentBuilderFactory;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import static org.junit.jupiter.api.Assertions.*;


class DocxToHtmlTest {

    @Test

    void convertsParagraphWithStylesAndRuns() {

        WordDocument.Spacing spacing = new WordDocument.Spacing(120, 240, 240, "auto");

        WordDocument.Indentation indentation = new WordDocument.Indentation(720, null, null, null);

        WordDocument.ParagraphProperties paragraphProperties = new WordDocument.ParagraphProperties(

                null,

                null,

                WordDocument.Alignment.CENTER,

                indentation,

                spacing,

                null,

                true,

                false,

                false,

                List.of(),

                null);

        WordDocument.RunProperties runProperties = new WordDocument.RunProperties(

                null,

                true,

                true,

                true,

                null,

                false,

                false,

                false,

                false,

                false,

                "FF0000",

                "yellow",

                null,

                32,

                null,

                Map.of("ascii", "Calibri", "hAnsi", "Arial"),

                null);

        WordDocument.Run run = new WordDocument.Run(runProperties, List.of(

                new WordDocument.Text("Bonjour", false)

        ));

        WordDocument.Paragraph paragraph = new WordDocument.Paragraph(paragraphProperties, List.of(run));

        WordDocument document = WordDocument.builder()

                .addBlock(paragraph)

                .build();

        DocxPackage pkg = DocxPackage.builder()

                .document(document)

                .build();

        DocxToHtml converter = new DocxToHtml("fr");

        String html = converter.convert(pkg);

        assertNotNull(html);

        assertTrue(html.contains("<style>"), "style block missing");

        assertFalse(html.contains("style=\""), "inline styles should be avoided");

        assertTrue(html.contains("<body class=\"docx-body\">"), "body class missing");

        assertTrue(html.contains("class=\"docx-paragraph p1\""), "paragraph class missing");

        assertTrue(html.contains(".docx-body .p1{text-align:center"), "paragraph alignment not mapped");

        assertTrue(html.contains("margin-left:36pt"), "indentation not converted");

        assertTrue(html.contains("line-height:1"), "line height not normalised");

        assertTrue(html.contains("class=\"docx-span s1\""), "run class missing");

        assertTrue(html.contains(".docx-body .s1{"), "run css definition missing");

        assertTrue(html.contains("font-weight:bold"), "bold formatting missing");

        assertTrue(html.contains("font-style:italic"), "italic formatting missing");

        assertTrue(html.contains("text-decoration-line:underline"), "underline missing");

        assertTrue(html.contains("color:#ff0000"), "color not normalised");

        assertTrue(html.contains("background-color:#ffff00"), "highlight not normalised");

        assertTrue(html.contains("font-family:Calibri"), "font family not cascaded");

        assertTrue(html.contains("Bonjour"), "text content missing");

    }

    @Test

    void rendersBackgroundShadingForParagraphsAndTables() throws Exception {

        Document xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();

        Element paragraphProps = xml.createElementNS(Namespaces.WORD_MAIN, "w:pPr");

        Element paragraphShd = xml.createElementNS(Namespaces.WORD_MAIN, "w:shd");

        paragraphShd.setAttributeNS(Namespaces.WORD_MAIN, "w:fill", "92D050");

        paragraphProps.appendChild(paragraphShd);

        WordDocument.ParagraphProperties shadedParagraphProperties = new WordDocument.ParagraphProperties(

                null,

                null,

                WordDocument.Alignment.LEFT,

                null,

                null,

                null,

                false,

                false,

                false,

                List.of(),

                paragraphProps);

        WordDocument.RunProperties baseRunProperties = new WordDocument.RunProperties(

                null,

                false,

                false,

                false,

                null,

                false,

                false,

                false,

                false,

                false,

                null,

                null,

                null,

                null,

                null,

                Map.of(),

                null);

        WordDocument.Run shadedParagraphRun = new WordDocument.Run(baseRunProperties, List.of(new WordDocument.Text("Paragraphe", false)));

        WordDocument.Paragraph shadedParagraph = new WordDocument.Paragraph(shadedParagraphProperties, List.of(shadedParagraphRun));

        WordDocument.ParagraphProperties emptyParagraphProperties = new WordDocument.ParagraphProperties(

                null,

                null,

                WordDocument.Alignment.LEFT,

                null,

                null,

                null,

                false,

                false,

                false,

                List.of(),

                null);

        WordDocument.Run cellRun1 = new WordDocument.Run(baseRunProperties, List.of(new WordDocument.Text("Cellule 1", false)));

        WordDocument.Run cellRun2 = new WordDocument.Run(baseRunProperties, List.of(new WordDocument.Text("Cellule 2", false)));

        WordDocument.Run cellRun3 = new WordDocument.Run(baseRunProperties, List.of(new WordDocument.Text("Cellule 3", false)));

        WordDocument.Paragraph cellParagraph1 = new WordDocument.Paragraph(emptyParagraphProperties, List.of(cellRun1));

        WordDocument.Paragraph cellParagraph2 = new WordDocument.Paragraph(emptyParagraphProperties, List.of(cellRun2));

        WordDocument.Paragraph cellParagraph3 = new WordDocument.Paragraph(emptyParagraphProperties, List.of(cellRun3));

        Element tableProps = xml.createElementNS(Namespaces.WORD_MAIN, "w:tblPr");

        Element tableShd = xml.createElementNS(Namespaces.WORD_MAIN, "w:shd");

        tableShd.setAttributeNS(Namespaces.WORD_MAIN, "w:fill", "C6EFCE");

        tableProps.appendChild(tableShd);

        WordDocument.TableProperties tableProperties = new WordDocument.TableProperties(null, null, null, null, tableProps);

        Element rowPropsShaded = xml.createElementNS(Namespaces.WORD_MAIN, "w:trPr");

        Element rowShd = xml.createElementNS(Namespaces.WORD_MAIN, "w:shd");

        rowShd.setAttributeNS(Namespaces.WORD_MAIN, "w:fill", "FFD966");

        rowPropsShaded.appendChild(rowShd);

        WordDocument.TableRowProperties rowWithShading = new WordDocument.TableRowProperties(false, null, null, null, rowPropsShaded);

        WordDocument.TableRowProperties plainRow = new WordDocument.TableRowProperties(false, null, null, null, null);

        Element cellPropsDirect = xml.createElementNS(Namespaces.WORD_MAIN, "w:tcPr");

        Element cellShdDirect = xml.createElementNS(Namespaces.WORD_MAIN, "w:shd");

        cellShdDirect.setAttributeNS(Namespaces.WORD_MAIN, "w:fill", "00B0F0");

        cellPropsDirect.appendChild(cellShdDirect);

        WordDocument.TableCellProperties directCellProps = new WordDocument.TableCellProperties(null, null, null, null, false, cellPropsDirect);

        WordDocument.TableCellProperties emptyCellProps = new WordDocument.TableCellProperties(null, null, null, null, false, null);

        WordDocument.TableCell directCell = new WordDocument.TableCell(directCellProps, List.of(cellParagraph1));

        WordDocument.TableCell rowFallbackCell = new WordDocument.TableCell(emptyCellProps, List.of(cellParagraph2));

        WordDocument.TableRow firstRow = new WordDocument.TableRow(rowWithShading, List.of(directCell, rowFallbackCell));

        WordDocument.TableCell tableFallbackCell = new WordDocument.TableCell(emptyCellProps, List.of(cellParagraph3));

        WordDocument.TableRow secondRow = new WordDocument.TableRow(plainRow, List.of(tableFallbackCell));

        WordDocument.Table table = new WordDocument.Table(tableProperties, List.of(firstRow, secondRow));

        WordDocument document = WordDocument.builder()

                .addBlock(shadedParagraph)

                .addBlock(table)

                .build();

        DocxPackage pkg = DocxPackage.builder()

                .document(document)

                .build();

        DocxToHtml converter = new DocxToHtml("fr");

        String html = converter.convert(pkg);

        assertTrue(html.contains("background-color:#92d050"), "paragraph shading missing");

        assertTrue(html.contains("background-color:#00b0f0"), "cell shading missing");

        assertTrue(html.contains("background-color:#ffd966"), "row shading fallback missing");

        assertTrue(html.contains("background-color:#c6efce"), "table shading fallback missing");

    }



    @Test
    void convertsTableThemeShading() throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        Document xml = factory.newDocumentBuilder().newDocument();
        Document themeDoc = factory.newDocumentBuilder().newDocument();
        Element themeRoot = themeDoc.createElementNS(Namespaces.DRAWINGML_MAIN, "a:theme");
        themeRoot.setAttribute("xmlns:a", Namespaces.DRAWINGML_MAIN);
        themeDoc.appendChild(themeRoot);
        Element themeElements = themeDoc.createElementNS(Namespaces.DRAWINGML_MAIN, "a:themeElements");
        themeRoot.appendChild(themeElements);
        Element clrScheme = themeDoc.createElementNS(Namespaces.DRAWINGML_MAIN, "a:clrScheme");
        clrScheme.setAttribute("name", "Custom");
        themeElements.appendChild(clrScheme);

        Element accent1 = themeDoc.createElementNS(Namespaces.DRAWINGML_MAIN, "a:accent1");
        Element accent1Color = themeDoc.createElementNS(Namespaces.DRAWINGML_MAIN, "a:srgbClr");
        accent1Color.setAttribute("val", "FF0000");
        accent1.appendChild(accent1Color);
        clrScheme.appendChild(accent1);

        Element accent2 = themeDoc.createElementNS(Namespaces.DRAWINGML_MAIN, "a:accent2");
        Element accent2Color = themeDoc.createElementNS(Namespaces.DRAWINGML_MAIN, "a:srgbClr");
        accent2Color.setAttribute("val", "0000FF");
        accent2.appendChild(accent2Color);
        clrScheme.appendChild(accent2);

        Element accent3 = themeDoc.createElementNS(Namespaces.DRAWINGML_MAIN, "a:accent3");
        Element accent3Color = themeDoc.createElementNS(Namespaces.DRAWINGML_MAIN, "a:srgbClr");
        accent3Color.setAttribute("val", "00FF00");
        accent3.appendChild(accent3Color);
        clrScheme.appendChild(accent3);

        WordDocument.RunProperties baseRunProperties = new WordDocument.RunProperties(
                null,
                false,
                false,
                false,
                null,
                false,
                false,
                false,
                false,
                false,
                null,
                null,
                null,
                null,
                null,
                Map.of(),
                null);

        WordDocument.ParagraphProperties emptyParagraphProperties = new WordDocument.ParagraphProperties(
                null,
                null,
                WordDocument.Alignment.LEFT,
                null,
                null,
                null,
                false,
                false,
                false,
                List.of(),
                null);

        WordDocument.Run themedCellRun = new WordDocument.Run(baseRunProperties, List.of(new WordDocument.Text("Theme 1", false)));
        WordDocument.Run fallbackRun = new WordDocument.Run(baseRunProperties, List.of(new WordDocument.Text("Fallback", false)));
        WordDocument.Run tableRun = new WordDocument.Run(baseRunProperties, List.of(new WordDocument.Text("Table", false)));

        WordDocument.Paragraph themedCellParagraph = new WordDocument.Paragraph(emptyParagraphProperties, List.of(themedCellRun));
        WordDocument.Paragraph fallbackParagraph = new WordDocument.Paragraph(emptyParagraphProperties, List.of(fallbackRun));
        WordDocument.Paragraph tableParagraph = new WordDocument.Paragraph(emptyParagraphProperties, List.of(tableRun));

        Element tableProps = xml.createElementNS(Namespaces.WORD_MAIN, "w:tblPr");
        Element tableShd = xml.createElementNS(Namespaces.WORD_MAIN, "w:shd");
        tableShd.setAttributeNS(Namespaces.WORD_MAIN, "w:themeFill", "accent1");
        tableProps.appendChild(tableShd);
        WordDocument.TableProperties tableProperties = new WordDocument.TableProperties(null, null, null, null, tableProps);

        Element rowProps = xml.createElementNS(Namespaces.WORD_MAIN, "w:trPr");
        Element rowShd = xml.createElementNS(Namespaces.WORD_MAIN, "w:shd");
        rowShd.setAttributeNS(Namespaces.WORD_MAIN, "w:themeFill", "accent2");
        rowShd.setAttributeNS(Namespaces.WORD_MAIN, "w:themeFillTint", "80");
        rowProps.appendChild(rowShd);
        WordDocument.TableRowProperties themedRowProperties = new WordDocument.TableRowProperties(false, null, null, null, rowProps);
        WordDocument.TableRowProperties plainRow = new WordDocument.TableRowProperties(false, null, null, null, null);

        Element cellProps = xml.createElementNS(Namespaces.WORD_MAIN, "w:tcPr");
        Element cellShd = xml.createElementNS(Namespaces.WORD_MAIN, "w:shd");
        cellShd.setAttributeNS(Namespaces.WORD_MAIN, "w:themeFill", "accent3");
        cellShd.setAttributeNS(Namespaces.WORD_MAIN, "w:themeFillShade", "66");
        cellProps.appendChild(cellShd);
        WordDocument.TableCellProperties themedCellProps = new WordDocument.TableCellProperties(null, null, null, null, false, cellProps);
        WordDocument.TableCellProperties emptyCellProps = new WordDocument.TableCellProperties(null, null, null, null, false, null);

        WordDocument.TableCell themedCell = new WordDocument.TableCell(themedCellProps, List.of(themedCellParagraph));
        WordDocument.TableCell rowFallbackCell = new WordDocument.TableCell(emptyCellProps, List.of(fallbackParagraph));
        WordDocument.TableCell tableFallbackCell = new WordDocument.TableCell(emptyCellProps, List.of(tableParagraph));

        WordDocument.TableRow themedRow = new WordDocument.TableRow(themedRowProperties, List.of(themedCell, rowFallbackCell));
        WordDocument.TableRow tableRow = new WordDocument.TableRow(plainRow, List.of(tableFallbackCell));

        WordDocument.Table table = new WordDocument.Table(tableProperties, List.of(themedRow, tableRow));

        WordDocument document = WordDocument.builder()
                .addBlock(table)
                .build();

        Theme theme = new Theme("word/theme/theme1.xml", themeDoc);

        DocxPackage pkg = DocxPackage.builder()
                .document(document)
                .theme(theme)
                .build();

        DocxToHtml converter = new DocxToHtml("fr");
        String html = converter.convert(pkg);

        assertTrue(html.contains("<table class=\"docx-table t1\">"), "table class missing");
        assertTrue(html.contains(".docx-body table.t1{background-color:#ff0000}"), "table theme shading missing");
        assertTrue(html.contains(".docx-body table.t1 td,.docx-body table.t1 th{background-color:#ff0000}"), "table cascade shading missing");
        assertTrue(html.contains("<tr class=\"docx-row r1\">"), "row shading class missing");
        assertTrue(html.contains(".docx-body tr.r1{background-color:#8080ff}"), "row theme shading missing");
        assertTrue(html.contains(".docx-body tr.r1 > td,.docx-body tr.r1 > th{background-color:#8080ff}"), "row cascade shading missing");
        assertTrue(html.contains("<td class=\"docx-cell c1\">"), "direct cell shading class missing");
        assertTrue(html.contains(".docx-body td.c1{background-color:#009900}"), "cell theme shading missing");
        assertTrue(html.contains("<td class=\"docx-cell\">"), "fallback cell should rely on row/table shading");
        assertFalse(html.contains("docx-cell c2"), "unexpected extra cell class for fallback shading");
    }


    @Test
    void rendersDemoTableHeaderWithTableStyle() throws Exception {
        DocxReader reader = new DocxReader();
        DocxPackage pkg = reader.read(Path.of("samples", "demo.docx"));

        DocxToHtml converter = new DocxToHtml("fr");
        String html = converter.convert(pkg);

        Matcher rowMatcher = Pattern.compile("<tr class=\"docx-row (r\\d+)\">").matcher(html);
        assertTrue(rowMatcher.find(), "header row class missing");
        String rowClass = rowMatcher.group(1);
        assertTrue(html.contains(".docx-body tr." + rowClass + "{background-color:#9bbb59}"), "header row background missing");
        assertTrue(html.contains(".docx-body tr." + rowClass + " > td,.docx-body tr." + rowClass + " > th{background-color:#9bbb59}"),
                "header row cascade background missing");

        String itemClass = extractRunClass(html, "ITEM");
        assertRunColor(html, itemClass, "#ffffff");
        String neededClass = extractRunClass(html, "NEEDED");
        assertRunColor(html, neededClass, "#ffffff");
    }

    private static String extractRunClass(String html, String text) {
        Matcher matcher = Pattern.compile("<span class=\"docx-span (s\\d+)\">" + text + "</span>").matcher(html);
        assertTrue(matcher.find(), "missing run class for " + text);
        return matcher.group(1);
    }

    private static void assertRunColor(String html, String cssClass, String expectedColor) {
        String marker = ".docx-body ." + cssClass + "{";
        int start = html.indexOf(marker);
        assertTrue(start >= 0, "missing CSS rule for " + cssClass);
        int end = html.indexOf('}', start);
        assertTrue(end > start, "malformed CSS rule for " + cssClass);
        String chunk = html.substring(start, end);
        assertTrue(chunk.contains("color:" + expectedColor), "expected " + expectedColor + " for " + cssClass);
    }
}



