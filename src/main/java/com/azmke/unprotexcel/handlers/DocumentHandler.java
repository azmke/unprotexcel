package com.azmke.unprotexcel.handlers;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class DocumentHandler {
    protected String filePath;
    protected Path tempDir;

    // Open and unpack the document file
    public void open(String filePath) throws IOException {
        this.filePath = filePath;
        this.tempDir = Files.createTempDirectory("unprotexcel");

        // Unzip the document file to the temporary directory
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(filePath))) {
            unzip(zis, tempDir);
        }
    }

    // Unzips the contents of the provided ZipInputStream into the specified directory
    private void unzip(ZipInputStream zis, Path destinationDir) throws IOException {
        ZipEntry entry;
        while ((entry = zis.getNextEntry()) != null) {
            Path newFilePath = destinationDir.resolve(entry.getName());
            if (entry.isDirectory()) {
                Files.createDirectories(newFilePath); // Create directories if the entry is a directory
            } else {
                // Ensure parent directories exist and write the file
                Files.createDirectories(newFilePath.getParent());
                try (OutputStream os = Files.newOutputStream(newFilePath)) {
                    byte[] buffer = new byte[1024];
                    int len;
                    while ((len = zis.read(buffer)) > 0) {
                        os.write(buffer, 0, len); // Write to the output stream
                    }
                }
            }
            zis.closeEntry(); // Close the entry
        }
    }

    // Save the changes by repacking the files into a new zip file at the specified path
    public void saveAs(String newFilePath) throws IOException {
        // Ensure the new file path has the same extension as the original file
        String originalExtension = getFileExtension(filePath);
        if (!newFilePath.endsWith(originalExtension)) {
            newFilePath += originalExtension; // Append the original extension if not present
        }

        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(newFilePath))) {
            zipDirectory(tempDir, zos, tempDir); // Zip the contents of the temp directory
        }
    }

    // Zips the contents of the specified directory into the provided ZipOutputStream
    private void zipDirectory(Path sourceDir, ZipOutputStream zos, Path baseDir) throws IOException {
        Files.walk(sourceDir).filter(path -> !Files.isDirectory(path)).forEach(path -> {
            // Create a ZipEntry for each file
            String zipEntryName = baseDir.relativize(path).toString().replace("\\", "/");
            try {
                zos.putNextEntry(new ZipEntry(zipEntryName)); // Start a new entry in the ZIP file
                Files.copy(path, zos); // Copy the file's contents to the ZIP output stream
                zos.closeEntry(); // Close the current entry
            } catch (IOException e) {
                throw new UncheckedIOException(e); // Handle exceptions
            }
        });
    }

    // Helper method to get the file extension from the file path
    private String getFileExtension(String filePath) {
        int lastIndex = filePath.lastIndexOf('.');
        if (lastIndex > 0 && lastIndex < filePath.length() - 1) {
            return filePath.substring(lastIndex); // Returns the extension including the dot, e.g., ".xlsx"
        }
        return ""; // Return empty string if no extension found
    }

    // Close the document handler and clean up the temporary directory
    public void close() throws IOException {
        if (tempDir != null && Files.exists(tempDir)) {
            deleteDirectory(tempDir);
        }
    }

    // Recursively delete the specified directory and its contents
    private void deleteDirectory(Path directory) throws IOException {
        // Walk through the directory and delete all files and subdirectories
        try (Stream<Path> paths = Files.walk(directory)) {
            paths.sorted(Comparator.reverseOrder()) // Sort paths in reverse order to delete files before directories
                .forEach(path -> {
                    try {
                        Files.delete(path); // Delete each file or directory
                    } catch (IOException e) {
                        System.err.println("Unable to delete " + path + ": " + e.getMessage());
                    }
                });
        }
    }
}
