package com.example.docx.parser;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/**
 * XML helper functions with safe defaults for parsing WordprocessingML parts.
 */
public final class XmlUtils {

    private static final DocumentBuilderFactory FACTORY = buildFactory();

    private XmlUtils() {
    }

    private static DocumentBuilderFactory buildFactory() {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        try {
            factory.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false);
        } catch (ParserConfigurationException e) {
            throw new IllegalStateException("Failed to configure XML factory", e);
        }
        factory.setExpandEntityReferences(false);
        return factory;
    }

    public static Document parse(InputStream inputStream) throws IOException {
        try {
            DocumentBuilder builder = FACTORY.newDocumentBuilder();
            return builder.parse(inputStream);
        } catch (Exception e) {
            throw new IOException("Failed to parse XML document", e);
        }
    }

    public static Document parse(byte[] bytes) throws IOException {
        return parse(new ByteArrayInputStream(bytes));
    }

    public static Document parse(String xml) throws IOException {
        return parse(xml.getBytes(StandardCharsets.UTF_8));
    }

    public static List<Element> children(Element parent, String namespaceUri, String localName) {
        List<Element> result = new ArrayList<>();
        NodeList nodes = parent.getChildNodes();
        for (int i = 0; i < nodes.getLength(); i++) {
            Node node = nodes.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                Element element = (Element) node;
                if ((namespaceUri == null || namespaceUri.equals(element.getNamespaceURI()))
                        && (localName == null || localName.equals(element.getLocalName()))) {
                    result.add(element);
                }
            }
        }
        return result;
    }

    public static Optional<Element> firstChild(Element parent, String namespaceUri, String localName) {
        NodeList nodes = parent.getChildNodes();
        for (int i = 0; i < nodes.getLength(); i++) {
            Node node = nodes.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                Element element = (Element) node;
                if ((namespaceUri == null || namespaceUri.equals(element.getNamespaceURI()))
                        && (localName == null || localName.equals(element.getLocalName()))) {
                    return Optional.of(element);
                }
            }
        }
        return Optional.empty();
    }

    public static List<Element> childElements(Element parent) {
        List<Element> result = new ArrayList<>();
        NodeList nodes = parent.getChildNodes();
        for (int i = 0; i < nodes.getLength(); i++) {
            Node node = nodes.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                result.add((Element) node);
            }
        }
        return result;
    }

    public static String attribute(Element element, String name) {
        if (element.hasAttribute(name)) {
            String value = element.getAttribute(name);
            return value.isEmpty() ? null : value;
        }
        int colon = name.indexOf(':');
        if (colon > 0) {
            String prefix = name.substring(0, colon);
            String local = name.substring(colon + 1);
            String namespace = element.lookupNamespaceURI(prefix);
            if (namespace != null && element.hasAttributeNS(namespace, local)) {
                String value = element.getAttributeNS(namespace, local);
                return value.isEmpty() ? null : value;
            }
        }
        return null;
    }

    public static Integer intAttribute(Element element, String name) {
        String value = attribute(element, name);
        if (value == null) {
            return null;
        }
        try {
            return Integer.parseInt(value);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    public static Long longAttribute(Element element, String name) {
        String value = attribute(element, name);
        if (value == null) {
            return null;
        }
        try {
            return Long.parseLong(value);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    public static boolean booleanElement(Element parent, String namespaceUri, String localName) {
        return firstChild(parent, namespaceUri, localName)
                .map(child -> {
                    String value = attribute(child, "val");
                    if (value == null) {
                        value = attribute(child, "w:val");
                    }
                    if (value == null) {
                        return true;
                    }
                    return value.isEmpty() || "true".equalsIgnoreCase(value) || "1".equals(value);
                })
                .orElse(false);
    }

    public static String textValue(Element parent, String namespaceUri, String localName) {
        return firstChild(parent, namespaceUri, localName)
                .map(Node::getTextContent)
                .map(String::trim)
                .filter(value -> !value.isEmpty())
                .orElse(null);
    }
}
