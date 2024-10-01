package com.azmke.unprotexcel.handlers;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.StandardOpenOption;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.function.Predicate;
import java.util.stream.Stream;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import com.azmke.unprotexcel.utils.XmlUtils;
import org.w3c.dom.Document;

public class ExcelHandler extends DocumentHandler {

    private Boolean initialWorkbookProtectionStatus = null;
    private Boolean initialModifyFilePasswordStatus = null;
    private Boolean initialSheetProtectionStatus = null;
    private Boolean initialRangeProtectionStatus = null;
    private Boolean initialVBAProtectionStatus = null;

    private boolean workbookProtectionRemoved = false;
    private boolean modifyFilePasswordRemoved = false;
    private boolean sheetProtectionRemoved = false;
    private boolean rangeProtectionRemoved = false;
    private boolean vbaProtectionRemoved = false;

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

    public boolean hasModifyFilePassword() throws IOException {
        if (initialModifyFilePasswordStatus == null) {
            initialModifyFilePasswordStatus = checkXmlTag(tempDir.resolve("xl/workbook.xml"), "fileSharing");
        }
        return initialModifyFilePasswordStatus;
    }

    public void removeModifyFilePassword() throws IOException {
        if (initialModifyFilePasswordStatus == Boolean.FALSE || modifyFilePasswordRemoved) {
            return;
        }
        updateXmlRemoveTag(tempDir.resolve("xl/workbook.xml"), "fileSharing");
        modifyFilePasswordRemoved = true;
    }

    public boolean hasWorkbookProtection() throws IOException {
        if (initialWorkbookProtectionStatus == null) {
            initialWorkbookProtectionStatus = checkXmlTag(tempDir.resolve("xl/workbook.xml"), "workbookProtection");
        }
        return initialWorkbookProtectionStatus;
    }

    public void removeWorkbookProtection() throws IOException {
        if (initialWorkbookProtectionStatus == Boolean.FALSE || workbookProtectionRemoved) {
            return;
        }
        updateXmlRemoveTag(tempDir.resolve("xl/workbook.xml"), "workbookProtection");
        workbookProtectionRemoved = true;
    }

    public boolean hasSheetProtection() throws IOException {
        if (initialSheetProtectionStatus == null) {
            initialSheetProtectionStatus = processWorksheetFiles(file -> {
                try {
                    return checkSheetProtection(file);
                } catch (IOException e) {
                    throw new RuntimeException("Error checking sheet protection", e);
                }
            });
        }
        return initialSheetProtectionStatus;
    }

    public void removeSheetProtection() throws IOException {
        if (initialSheetProtectionStatus == Boolean.FALSE || sheetProtectionRemoved) {
            return;
        }
        processWorksheetFiles(file -> {
            try {
                updateSheetRemoveProtection(file);
            } catch (IOException e) {
                throw new RuntimeException("Error removing sheet protection", e);
            }
            return false;
        });
        sheetProtectionRemoved = true;
    }

    public boolean hasRangeProtection() throws IOException {
        if (initialRangeProtectionStatus == null) {
            initialRangeProtectionStatus = processWorksheetFiles(file -> {
                try {
                    return checkRangeProtection(file);
                } catch (IOException e) {
                    throw new RuntimeException("Error checking range protection", e);
                }
            });
        }
        return initialRangeProtectionStatus;
    }

    public void removeRangeProtection() throws IOException {
        if (initialRangeProtectionStatus == Boolean.FALSE || rangeProtectionRemoved) {
            return;
        }
        processWorksheetFiles(file -> {
            try {
                updateRangeRemoveProtection(file);
            } catch (IOException e) {
                throw new RuntimeException("Error removing range protection", e);
            }
            return false;
        });
        rangeProtectionRemoved = true;
    }

    // General method to check for an XML tag in a given file
    private boolean checkXmlTag(Path xmlPath, String tagName) throws IOException {
        if (!Files.exists(xmlPath)) {
            throw new IOException(xmlPath.getFileName() + " not found in the extracted files.");
        }

        Document document = XmlUtils.parseXml(xmlPath);
        return XmlUtils.hasXmlTag(document, tagName);
    }

    // General method to remove an XML tag in a given file
    private void updateXmlRemoveTag(Path xmlPath, String tagName) throws IOException {
        if (!Files.exists(xmlPath)) {
            throw new IOException(xmlPath.getFileName() + " not found in the extracted files.");
        }

        Document document = XmlUtils.parseXml(xmlPath);
        XmlUtils.removeXmlTag(document, tagName);
        XmlUtils.writeXml(document, xmlPath);
    }

    private boolean checkSheetProtection(Path file) throws IOException {
        return checkXmlTag(file, "sheetProtection");
    }

    private void updateSheetRemoveProtection(Path file) throws IOException {
        updateXmlRemoveTag(file, "sheetProtection");
    }

    private boolean checkRangeProtection(Path file) throws IOException {
        return checkXmlTag(file, "protectedRanges");
    }

    private void updateRangeRemoveProtection(Path file) throws IOException {
        updateXmlRemoveTag(file, "protectedRanges");
    }

    private boolean processWorksheetFiles(Predicate<Path> action) throws IOException {
        Path worksheetsDir = tempDir.resolve("xl/worksheets");

        if (!Files.exists(worksheetsDir) || !Files.isDirectory(worksheetsDir)) {
            throw new IOException("The directory " + worksheetsDir + " does not exist or is not a directory.");
        }

        try (Stream<Path> paths = Files.list(worksheetsDir)) {
            return paths.filter(Files::isRegularFile)
                        .filter(path -> path.toString().endsWith(".xml"))
                        .anyMatch(action);
        }
    }

    public boolean hasVBAProtection() throws IOException {
        if (initialVBAProtectionStatus == null) {
            Path vbaProjectPath = tempDir.resolve("xl/vbaProject.bin");

            if (!Files.exists(vbaProjectPath)) {
                initialVBAProtectionStatus = false;
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
                            initialVBAProtectionStatus = true;
                            return true;
                        }
                    }
                }
            }
            initialVBAProtectionStatus = false;
        }
        return initialVBAProtectionStatus;
    }

    public void removeVBAProtection() throws IOException {
        if (initialVBAProtectionStatus == Boolean.FALSE || vbaProtectionRemoved) {
            return;
        }
        Path vbaProjectPath = tempDir.resolve("xl/vbaProject.bin");

        if (!Files.exists(vbaProjectPath)) {
            throw new IOException("The file " + vbaProjectPath + " does not exist.");
        }

        String dpbString = "DPB=";
        String dpxString = "DPx=";
        byte[] dpbBytes = dpbString.getBytes(StandardCharsets.UTF_8);
        byte[] dpxBytes = dpxString.getBytes(StandardCharsets.UTF_8);

        byte[] fileContent = Files.readAllBytes(vbaProjectPath);
        boolean found = false;

        for (int i = 0; i <= fileContent.length - dpbBytes.length; i++) {
            boolean match = true;
            for (int j = 0; j < dpbBytes.length; j++) {
                if (fileContent[i + j] != dpbBytes[j]) {
                    match = false;
                    break;
                }
            }
            if (match) {
                found = true;
                System.arraycopy(dpxBytes, 0, fileContent, i, dpxBytes.length);
            }
        }

        if (found) {
            try (OutputStream outputStream = Files.newOutputStream(vbaProjectPath, StandardOpenOption.WRITE)) {
                outputStream.write(fileContent);
            }
            vbaProtectionRemoved = true;
        }
    }
}