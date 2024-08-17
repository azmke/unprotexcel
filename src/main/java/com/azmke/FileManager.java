package com.azmke;

import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.filechooser.FileFilter;

public class FileManager {
    private LanguageManager languageManager;

    public FileManager(LanguageManager languageManager) {
        this.languageManager = languageManager;
    }

    // Method to open the file dialog and update the text field
    public File showOpenDialog(JTextField textField) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(languageManager.getString("dialog.open.title"));

        // Set file filters for Excek files
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fileChooser.setAcceptAllFileFilterUsed(false);

        FileFilter excelFilter = createExcelFilter();
        FileFilter allFilter = createAllFilter();

        fileChooser.addChoosableFileFilter(excelFilter);
        fileChooser.addChoosableFileFilter(allFilter);

        fileChooser.setFileFilter(excelFilter);

        // Show dialog
        int result = fileChooser.showOpenDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();

            // Update the text field with the selected file path
            textField.setText(selectedFile.getAbsolutePath());
            App.log(languageManager.getString("message.fileSelected", selectedFile.getAbsolutePath()));

            return selectedFile;
        }

        return null;
    }

    // Method to open the save dialog for a new file
    public File showSaveDialog(File originalFile) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle(languageManager.getString("dialog.save.title"));

        // Set the initial directory and file name based on the original file path
        if (originalFile.exists()) {
            fileChooser.setCurrentDirectory(originalFile.getParentFile());
            // Create the suggested filename based on the original name and extension
            String originalName = originalFile.getName();
            String originalExtension = FileUtils.getFileExtension(originalFile);
            String suggestedName = originalName.substring(0, originalName.lastIndexOf('.')) + "_unlocked" + originalExtension;
            fileChooser.setSelectedFile(new File(suggestedName));
        }

        // Set file filters for Excel files
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fileChooser.setAcceptAllFileFilterUsed(false);

        FileFilter excelFilter = createExcelFilter();
        FileFilter allFilter = createAllFilter();

        fileChooser.addChoosableFileFilter(excelFilter);
        fileChooser.addChoosableFileFilter(allFilter);

        fileChooser.setFileFilter(excelFilter);

        // Show dialog
        int result = fileChooser.showSaveDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            File destinationFile = fileChooser.getSelectedFile();

            // If the user forgets to specify the extension, use the original file's extension
            if (!destinationFile.getName().contains(".")) {
                String originalExtension = FileUtils.getFileExtension(originalFile);
                if (!originalExtension.isEmpty()) {
                    destinationFile = new File(destinationFile.getAbsolutePath() + originalExtension);
                }
            }

            // Check if the file already exists
            if (destinationFile.exists()) {
                int response = JOptionPane.showConfirmDialog(null,
                    languageManager.getString("dialog.fileExists.text", destinationFile.getName()),
                    languageManager.getString("dialog.fileExists.title"),
                    JOptionPane.YES_NO_OPTION);
                
                if (response == JOptionPane.NO_OPTION) {
                    return showSaveDialog(originalFile); // Open again if not replacing
                }
            }

            // Return destination file
            return destinationFile;
        }

        // No file path selected
        return null;
    }

    private FileFilter createExcelFilter() {
        return new FileFilter() {
            @Override
            public boolean accept(File f) {
                return f.isDirectory() || f.getName().toLowerCase().endsWith(".xls") || f.getName().toLowerCase().endsWith(".xlsx");
            }

            @Override
            public String getDescription() {
                return languageManager.getString("dialog.extension.excel");
            }
        };
    }

    private FileFilter createAllFilter() {
        return new FileFilter() {
            @Override
            public boolean accept(File f) {
                return true;
            }

            @Override
            public String getDescription() {
                return languageManager.getString("dialog.extension.all");
            }
        };
    }
}