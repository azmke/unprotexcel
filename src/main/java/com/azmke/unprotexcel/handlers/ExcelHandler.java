package com.azmke.unprotexcel.handlers;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.stream.Stream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class ExcelHandler extends DocumentHandler {

    // Check if an Excel file is password protected
    public static boolean hasOpenFilePassword(String filePath) {
        try (InputStream is = new FileInputStream(filePath)) {
            POIFSFileSystem fs = new POIFSFileSystem(is);
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor decryptor = Decryptor.getInstance(info);
            return !decryptor.verifyPassword("");
        } catch (Exception e) {
            return false;
        }
    }

    // Check if the Excel file has a modify file password
    public boolean hasModifyFilePassword() throws IOException {
        Path workbookXmlPath = tempDir.resolve("xl/workbook.xml");

        if (!Files.exists(workbookXmlPath)) {
            throw new IOException("workbook.xml not found in the extracted files.");
        }

        try (InputStream inputStream = Files.newInputStream(workbookXmlPath)) {
            // Parse workbook.xml
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document document = builder.parse(inputStream);
    
            // Search for <fileSharing ... /> tag
            NodeList nodes = document.getElementsByTagName("fileSharing");
            return nodes.getLength() > 0; // true, if the tag exists
        } catch (Exception e) {
            throw new IOException("Error processing workbook XML: " + e.getMessage(), e);
        }
    }

    // Remove the modify file password from workbook.xml
    public void removeModifyFilePassword() throws IOException {
        Path workbookXmlPath = tempDir.resolve("xl/workbook.xml");
    
        // Check if the workbook.xml file exists
        if (!Files.exists(workbookXmlPath)) {
            throw new IOException("workbook.xml not found in the extracted files.");
        }
    
        // First, read and process the file
        Document doc;
        try (InputStream inputStream = Files.newInputStream(workbookXmlPath)) {
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            doc = dBuilder.parse(inputStream);
        } catch (Exception e) {
            throw new IOException("Error processing workbook XML during read: " + e.getMessage(), e);
        }
    
        // Remove the <fileSharing> element if it exists
        NodeList nodes = doc.getElementsByTagName("fileSharing");
        while (nodes.getLength() > 0) {
            Node parent = nodes.item(0).getParentNode();
            parent.removeChild(nodes.item(0));
        }
    
        // Now write the changes back to workbook.xml
        try (OutputStream outputStream = Files.newOutputStream(workbookXmlPath)) {
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource source = new DOMSource(doc);
            StreamResult result = new StreamResult(outputStream);
            transformer.transform(source, result);
        } catch (Exception e) {
            throw new IOException("Error processing workbook XML during write: " + e.getMessage(), e);
        }
    }

    // Check if the Excel file has a workbook protection
    public boolean hasWorkbookProtection() throws IOException {
        Path workbookXmlPath = tempDir.resolve("xl/workbook.xml");

        if (!Files.exists(workbookXmlPath)) {
            throw new IOException("workbook.xml not found in the extracted files.");
        }

        try (InputStream inputStream = Files.newInputStream(workbookXmlPath)) {
            // Parse workbook.xml
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document document = builder.parse(inputStream);
    
            // Search for <workbookProtection ... /> tag
            NodeList nodes = document.getElementsByTagName("workbookProtection");
            return nodes.getLength() > 0; // true, if the tag exists
        } catch (Exception e) {
            throw new IOException("Error processing workbook XML: " + e.getMessage(), e);
        }
    }

    // Remove the workbook protection from workbook.xml
    public void removeWorkbookProtection() throws IOException {
        Path workbookXmlPath = tempDir.resolve("xl/workbook.xml");

        if (!Files.exists(workbookXmlPath)) {
            throw new IOException("workbook.xml not found in the extracted files.");
        }

        try (InputStream inputStream = Files.newInputStream(workbookXmlPath);
            OutputStream outputStream = Files.newOutputStream(workbookXmlPath)) {
            // Parse and edit the workbook.xml
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(inputStream);

            // Remove the <fileSharing> element if it exists
            NodeList nodes = doc.getElementsByTagName("workbookProtection");
            while (nodes.getLength() > 0) {
                Node parent = nodes.item(0).getParentNode();
                parent.removeChild(nodes.item(0));
            }

            // Save the changes back to workbook.xml
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource source = new DOMSource(doc);
            StreamResult result = new StreamResult(outputStream);
            transformer.transform(source, result);
        } catch (Exception e) {
            throw new IOException("Error processing workbook XML: " + e.getMessage(), e);
        }
    }

    public void removeSheetProtection() throws IOException {
        // Construct the path to the worksheets directory
        Path worksheetsDir = tempDir.resolve("xl/worksheets");

        // Check if the directory exists
        if (!Files.exists(worksheetsDir) || !Files.isDirectory(worksheetsDir)) {
            throw new IOException("The directory " + worksheetsDir + " does not exist or is not a directory.");
        }

        try (Stream<Path> paths = Files.list(worksheetsDir)) {
            paths.filter(Files::isRegularFile)
                 .filter(path -> path.toString().endsWith(".xml"))
                 .forEach(this::removeSheetProtectionFromFile);
        }
    }

    public void removeSheetProtectionFromFile(Path file) throws IOException {
        // Process the XML document
        try (InputStream inputStream = Files.newInputStream(file); 
             OutputStream outputStream = Files.newOutputStream(file)) {
            
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document document = builder.parse(inputStream);

            NodeList sheetProtectionNodes = document.getElementsByTagName("sheetProtection");
            while (sheetProtectionNodes.getLength() > 0) {
                Node parent = sheetProtectionNodes.item(0).getParentNode();
                parent.removeChild(sheetProtectionNodes.item(0));
            }

            // Save changes to the file
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource source = new DOMSource(document);
            StreamResult result = new StreamResult(outputStream);
            transformer.transform(source, result);
        } catch (Exception e) {
            throw new IOException("Error processing file " + file.getFileName() + ": " + e.getMessage());
        }
    }

    public boolean hasVBAProtection() throws IOException {
        Path vbaProjectPath = tempDir.resolve("xl/vbaProject.bin");

        if (!Files.exists(vbaProjectPath)) {
            return false;
        }

        String dpbString = "DPB=";
        byte[] dpbBytes = dpbString.getBytes(StandardCharsets.UTF_8);

        try (InputStream inputStream = Files.newInputStream(vbaProjectPath)) {
            byte[] buffer = new byte[4096];
            int bytesRead;
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                for (int i = 0; i <= bytesRead - dpbBytes.length; i++) {
                    boolean found = true;
                    for (int j = 0; j < dpbBytes.length; j++) {
                        if (buffer[i + j] != dpbBytes[j]) {
                            found = false;
                            break;
                        }
                    }
                    if (found) {
                        return true;
                    }
                }
            }
        }

        return false;
    }
}
