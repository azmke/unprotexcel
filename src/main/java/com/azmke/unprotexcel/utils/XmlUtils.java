package com.azmke.unprotexcel.utils;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class XmlUtils {

    // Parse an XML file and return the Document
    public static Document parseXml(Path xmlPath) throws IOException {
        try (InputStream inputStream = Files.newInputStream(xmlPath)) {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            return builder.parse(inputStream);
        } catch (Exception e) {
            throw new IOException("Error processing " + xmlPath.getFileName() + ": " + e.getMessage(), e);
        }
    }

    // Check for the existence of a specific XML tag
    public static boolean hasXmlTag(Document document, String tagName) {
        NodeList nodes = document.getElementsByTagName(tagName);
        return nodes.getLength() > 0;
    }

    // Remove a specific XML tag
    public static void removeXmlTag(Document document, String tagName) {
        NodeList nodes = document.getElementsByTagName(tagName);
        while (nodes.getLength() > 0) {
            Node parent = nodes.item(0).getParentNode();
            parent.removeChild(nodes.item(0));
        }
    }

    // Write changes to an XML file
    public static void writeXml(Document document, Path xmlPath) throws IOException {
        try (OutputStream outputStream = Files.newOutputStream(xmlPath)) {
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource source = new DOMSource(document);
            StreamResult result = new StreamResult(outputStream);
            transformer.transform(source, result);
        } catch (Exception e) {
            throw new IOException("Error writing to " + xmlPath.getFileName() + ": " + e.getMessage(), e);
        }
    }
}
