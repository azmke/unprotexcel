package com.azmke;

import javax.swing.*;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

public class UnlockWorker extends SwingWorker<Void, String> {
    private final String filePath;
    private final LanguageManager languageManager;
    private final FileManager fileManager;

    public UnlockWorker(String filePath, LanguageManager languageManager, FileManager fileManager) {
        this.filePath = filePath;
        this.languageManager = languageManager;
        this.fileManager = fileManager;
    }

    @Override
    protected Void doInBackground() {
        try {
            File originalFile = new File(filePath);
            String originalFileName = originalFile.getName();
            String tempDirPath = Files.createTempDirectory("excel_unprotect").toString();
            File tempDir = new File(tempDirPath);

            // Step 1: Copy Excel file to temporary directory
            App.log(languageManager.getString("message.temp", tempDir.getAbsolutePath()));
            File tempFile = new File(tempDir, originalFileName);
            Files.copy(originalFile.toPath(), tempFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

            // Step 2: Change file extension to .zip
            App.log(languageManager.getString("message.extract"));
            File zipFile = new File(tempDir, originalFileName + ".zip");
            tempFile.renameTo(zipFile);
            File unpackedDir = new File(tempDir, originalFileName + "_unpacked");
            unpackedDir.mkdir();
            FileUtils.unzip(zipFile, unpackedDir);

            // Step 3: Modify XML file in ./xl/worksheets/
            App.log(languageManager.getString("message.xml"));
            File worksheetsDir = new File(unpackedDir, "xl/worksheets");
            for (File xmlFile : worksheetsDir.listFiles((dir, name) -> name.matches("sheet\\d+\\.xml"))) {
                FileUtils.removeSheetProtection(xmlFile);
            }

            // Step 4: Zip all files and set original file extension
            App.log(languageManager.getString("message.zip"));
            File unlockedZipFile = new File(tempDir, originalFileName.replaceFirst("\\.\\w+$", "") + "_unlocked.zip");
            FileUtils.zip(unpackedDir, unlockedZipFile);
            File unlockedFile = new File(tempDir, originalFileName.replaceFirst("\\.\\w+$", "") + "_unlocked" + FileUtils.getFileExtension(originalFile));
            unlockedZipFile.renameTo(unlockedFile);

            // Step 5: Copy file to desired destination
            File destinationFile = fileManager.showSaveDialog(originalFile);
            if (destinationFile != null) {
                Files.copy(unlockedFile.toPath(), destinationFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
                App.log(languageManager.getString("message.success", destinationFile.getAbsolutePath()));
                JOptionPane.showMessageDialog(null, languageManager.getString("dialog.success.text"), languageManager.getString("dialog.success.title"), JOptionPane.INFORMATION_MESSAGE);
            }

        } catch (Exception ex) {
            ex.printStackTrace();
            App.log(languageManager.getString("message.failed", ex.getMessage()));
            JOptionPane.showMessageDialog(null, languageManager.getString("dialog.error.text", ex.getMessage()), languageManager.getString("dialog.error.title"), JOptionPane.ERROR_MESSAGE);
        }
        return null;
    }
}
