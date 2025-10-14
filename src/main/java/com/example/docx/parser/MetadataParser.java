package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.DocxPackage;
import com.example.docx.model.metadata.AppProperties;
import com.example.docx.model.metadata.CoreProperties;
import com.example.docx.model.metadata.CustomProperties;
import com.example.docx.model.metadata.CustomProperties.CustomProperty;
import com.example.docx.model.metadata.CustomProperties.ValueType;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.IOException;
import java.io.InputStream;
import java.time.Instant;
import java.time.format.DateTimeFormatter;
import java.util.Base64;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

/**
 * Parses core/app/custom metadata documents.
 */
public final class MetadataParser {

    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ISO_DATE_TIME;

    public void parse(DocxArchive archive, DocxPackage.Builder builder) throws IOException {
        CoreProperties core = parseCore(archive);
        if (core != null) {
            builder.coreProperties(core);
        }
        AppProperties app = parseApp(archive);
        if (app != null) {
            builder.appProperties(app);
        }
        CustomProperties custom = parseCustom(archive);
        if (custom != null) {
            builder.customProperties(custom);
        }
    }

    private CoreProperties parseCore(DocxArchive archive) throws IOException {
        if (!archive.exists("docProps/core.xml")) {
            return null;
        }
        try (InputStream input = archive.open("docProps/core.xml")) {
            Document document = XmlUtils.parse(input);
            Element root = document.getDocumentElement();
            CoreProperties.Builder builder = CoreProperties.builder();
            builder.title(XmlUtils.textValue(root, Namespaces.DC_ELEMENTS, "title"));
            builder.subject(XmlUtils.textValue(root, Namespaces.DC_ELEMENTS, "subject"));
            builder.creator(XmlUtils.textValue(root, Namespaces.DC_ELEMENTS, "creator"));
            builder.description(XmlUtils.textValue(root, Namespaces.DC_ELEMENTS, "description"));
            builder.keywords(XmlUtils.textValue(root, Namespaces.CP_CORE_PROPERTIES, "keywords"));
            builder.lastModifiedBy(XmlUtils.textValue(root, Namespaces.CP_CORE_PROPERTIES, "lastModifiedBy"));
            builder.revision(XmlUtils.textValue(root, Namespaces.CP_CORE_PROPERTIES, "revision"));
            builder.category(XmlUtils.textValue(root, Namespaces.CP_CORE_PROPERTIES, "category"));
            builder.contentStatus(XmlUtils.textValue(root, Namespaces.CP_CORE_PROPERTIES, "contentStatus"));
            builder.language(XmlUtils.textValue(root, Namespaces.DC_ELEMENTS, "language"));
            builder.identifier(XmlUtils.textValue(root, Namespaces.DC_ELEMENTS, "identifier"));
            builder.version(XmlUtils.textValue(root, Namespaces.CP_CORE_PROPERTIES, "version"));
            String created = XmlUtils.textValue(root, Namespaces.DC_TERMS, "created");
            if (created != null) {
                builder.created(parseInstant(created));
            }
            String modified = XmlUtils.textValue(root, Namespaces.DC_TERMS, "modified");
            if (modified != null) {
                builder.modified(parseInstant(modified));
            }
            return builder.build();
        }
    }

    private AppProperties parseApp(DocxArchive archive) throws IOException {
        if (!archive.exists("docProps/app.xml")) {
            return null;
        }
        try (InputStream input = archive.open("docProps/app.xml")) {
            Document document = XmlUtils.parse(input);
            Element root = document.getDocumentElement();
            Map<String, String> values = new HashMap<>();
            for (Element element : XmlUtils.childElements(root)) {
                values.put(element.getLocalName(), element.getTextContent());
            }
            return new AppProperties(
                    values.get("Application"),
                    values.get("AppVersion"),
                    values.get("Company"),
                    values.get("Manager"),
                    values.get("PresentationFormat"),
                    parseInteger(values.get("Pages")),
                    parseInteger(values.get("Words")),
                    parseInteger(values.get("Characters")),
                    parseInteger(values.get("Paragraphs")),
                    parseInteger(values.get("Lines")),
                    parseInteger(values.get("Slides")),
                    parseInteger(values.get("Notes")),
                    parseInteger(values.get("TotalTime")),
                    parseBoolean(values.get("Template")),
                    parseBoolean(values.get("SharedDoc")),
                    parseInteger(values.get("DocSecurity")),
                    parseBoolean(values.get("LinksUpToDate")),
                    values);
        }
    }

    private CustomProperties parseCustom(DocxArchive archive) throws IOException {
        if (!archive.exists("docProps/custom.xml")) {
            return null;
        }
        try (InputStream input = archive.open("docProps/custom.xml")) {
            Document document = XmlUtils.parse(input);
            Element root = document.getDocumentElement();
            CustomProperties.Builder builder = CustomProperties.builder();
            for (Element property : XmlUtils.childElements(root)) {
                if (!"property".equals(property.getLocalName())) {
                    continue;
                }
                String name = property.getAttribute("name");
                Element valueEl = XmlUtils.childElements(property).stream().findFirst().orElse(null);
                if (valueEl == null) {
                    continue;
                }
                String local = valueEl.getLocalName();
                ValueType type = mapValueType(local);
                Object value = parseCustomValue(local, valueEl.getTextContent());
                builder.add(new CustomProperty(name, type, value));
            }
            return builder.build();
        }
    }

    private ValueType mapValueType(String localName) {
        return switch (localName) {
            case "lpwstr", "lpstr", "bstr" -> ValueType.STRING;
            case "i1", "i2", "i4", "i8", "int", "ui1", "ui2", "ui4", "ui8", "uint" -> ValueType.INTEGER;
            case "r4" -> ValueType.FLOAT;
            case "r8" -> ValueType.DOUBLE;
            case "decimal" -> ValueType.DECIMAL;
            case "bool" -> ValueType.BOOLEAN;
            case "date" -> ValueType.DATE_TIME;
            case "filetime" -> ValueType.FILETIME;
            case "oblob", "blob" -> ValueType.BINARY;
            default -> ValueType.STRING;
        };
    }

    private Object parseCustomValue(String localName, String value) {
        return switch (localName) {
            case "i1", "i2", "i4", "i8", "int", "ui1", "ui2", "ui4", "ui8", "uint" -> parseInteger(value);
            case "bool" -> parseBoolean(value);
            case "r4", "r8", "decimal" -> parseDouble(value);
            case "date", "filetime" -> parseInstant(value);
            case "oblob", "blob" -> Base64.getDecoder().decode(value);
            default -> value;
        };
    }

    private Instant parseInstant(String value) {
        try {
            return Instant.from(DATE_FORMAT.parse(value));
        } catch (Exception e) {
            return null;
        }
    }

    private Integer parseInteger(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        try {
            return Integer.parseInt(value.trim());
        } catch (NumberFormatException e) {
            return null;
        }
    }

    private Double parseDouble(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        try {
            return Double.parseDouble(value.trim());
        } catch (NumberFormatException e) {
            return null;
        }
    }

    private Boolean parseBoolean(String value) {
        if (value == null) {
            return null;
        }
        String normalized = value.trim().toLowerCase(Locale.ROOT);
        if (normalized.isEmpty()) {
            return null;
        }
        return normalized.equals("true") || normalized.equals("1");
    }
}
