package com.azmke.unprotexcel.gui;

import com.azmke.unprotexcel.App;
import com.azmke.unprotexcel.handlers.ExcelHandler;
import com.azmke.unprotexcel.utils.FileManager;
import com.azmke.unprotexcel.utils.LanguageManager;

import com.formdev.flatlaf.FlatDarkLaf;
import com.formdev.flatlaf.FlatLightLaf;

import javax.swing.*;

import java.awt.*;
import java.awt.dnd.DropTarget;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

import java.io.File;
import java.io.IOException;

import java.net.URI;
import java.net.URISyntaxException;

public class GUI {
    private static JFrame frame;
    private static JMenuBar menuBar;
    private static JTextField textField;
    private static JButton browseButton;
    private static JButton scanButton;
    private static JButton unlockButton;
    private static ExcelTable protections;

    private static boolean isDarkTheme = false;

    private static LanguageManager languageManager;
    private static FileManager fileManager;
    private static ExcelHandler excelHandler;

    public GUI() {
        // Initialize language and file managers for localization and file operations
        languageManager = new LanguageManager("en");
        fileManager = new FileManager(languageManager);
        
        // Create the main application frame
        frame = new JFrame(languageManager.getString("app.title"));
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(600, 300);
        frame.setMinimumSize(new Dimension(500, 300));
        frame.setMaximumSize(new Dimension(800, 600));
        
        // Set a modern look and feel for the application
        try {
            updateTheme(isDarkTheme);
        } catch (UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }
    
        // Create and set the menu bar for the frame
        menuBar = createMenuBar();
        frame.setJMenuBar(menuBar);
        
        // Create a main panel with GridBagLayout for flexible component arrangement
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);
        gbc.fill = GridBagConstraints.BOTH;
    
        // Create a text field for file path input with drag-and-drop support
        textField = new JTextField();
        textField.setDropTarget(new DropTarget(textField, new DragAndDropHandler(textField, languageManager)));
        
        // Add the text field to the panel
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.gridwidth = 2;
        gbc.weightx = 1.0;
        panel.add(textField, gbc);
        
        // Create and add a browse button to select files
        browseButton = new JButton(languageManager.getString("button.browse"));
        gbc.gridx = 2;
        gbc.gridy = 0;
        gbc.gridwidth = 1;
        gbc.weightx = 0;
        panel.add(browseButton, gbc);
    
        // Create a table model and table for displaying data
        protections = new ExcelTable(languageManager);
        //protections.setModifyFilePasswordDetected(DetectionStatus.TRUE);
        //protections.setModifyFilePasswordRemoved(RemovalStatus.TRUE);
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.gridwidth = 3;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        panel.add(protections.getTable(), gbc);
    
        // Create horizonzal space
        JPanel panel2 = new JPanel();
        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.gridwidth = 1;
        gbc.weightx = 1.0;
        gbc.weighty = 0;
        panel.add(panel2, gbc);

        // Create a scan button
        scanButton = new JButton(languageManager.getString("button.scan"));
        scanButton.setMinimumSize(browseButton.getPreferredSize());
        gbc.gridx = 1;
        gbc.gridy = 2;
        gbc.gridwidth = 1;
        gbc.weightx = 0;
        gbc.weighty = 0;
        panel.add(scanButton, gbc);

        // Create an unlock button that is initially disabled
        unlockButton = new JButton(languageManager.getString("button.unlock"));
        unlockButton.setMinimumSize(browseButton.getPreferredSize());
        unlockButton.setEnabled(false);
        gbc.gridx = 2;
        gbc.gridy = 2;
        gbc.gridwidth = 1;
        gbc.weightx = 0;
        gbc.weighty = 0;
        //gbc.anchor = GridBagConstraints.EAST;
        panel.add(unlockButton, gbc);
        
        // Action listener for the browse button to open a file dialog
        browseButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                File selectedFile = fileManager.showOpenDialog();
                if (selectedFile != null) {
                    // Update path in text field
                    textField.setText(selectedFile.getAbsolutePath());

                    // Reset detection and removal status
                    protections.resetDetectionStatus();
                    protections.resetRemovalStatus();

                    App.log(languageManager.getString("message.fileSelected", selectedFile.getAbsolutePath()));
                }
            }
        });

        // Action listener for the scan button to handle protection detection
        scanButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent event) {
                String filePath = textField.getText();
                if (filePath.isEmpty()) {
                    // Show error dialog if no file is selected
                    JOptionPane.showMessageDialog(frame, languageManager.getString("dialog.noFile.text"), languageManager.getString("dialog.noFile.title"), JOptionPane.ERROR_MESSAGE);
                    return;
                }

                try {
                    // Disable buttons temporarily
                    browseButton.setEnabled(false);
                    scanButton.setEnabled(false);
                    unlockButton.setEnabled(false);

                    // Reset detection and removal status
                    protections.resetDetectionStatus();
                    protections.resetRemovalStatus();

                    // Safely close existing Excel handler for previous file
                    if (excelHandler != null) {
                        excelHandler.close();
                    }
                    excelHandler = new ExcelHandler();

                    // Check if Excel file needs a password to open and therefore has been encrypted
                    if (ExcelHandler.hasOpenFilePassword(filePath)) {
                        protections.setOpenFilePasswordDetected(DetectionStatus.TRUE);
                        protections.setModifyFilePasswordDetected(DetectionStatus.FAILED);
                        protections.setWorkbookProtectionDetected(DetectionStatus.FAILED);
                        protections.setSheetProtectionDetected(DetectionStatus.FAILED);
                        protections.setRangeProtectionDetected(DetectionStatus.FAILED);
                        protections.setVBAProtectionDetected(DetectionStatus.FAILED);

                        // Show error dialog if a password is required to open the file
                        JOptionPane.showMessageDialog(frame, languageManager.getString("dialog.filePassword.text"), languageManager.getString("dialog.filePassword.title"), JOptionPane.ERROR_MESSAGE);
                        scanButton.setEnabled(true);
                        return;
                    }
                    protections.setOpenFilePasswordDetected(DetectionStatus.FALSE);

                    // Open and extract Excel file
                    excelHandler.open(filePath);

                    // Scan for protections in Excel file
                    protections.setModifyFilePasswordDetected(
                        excelHandler.hasModifyFilePassword()
                            ? DetectionStatus.TRUE
                            : DetectionStatus.FALSE
                    );

                    protections.setWorkbookProtectionDetected(
                        excelHandler.hasWorkbookProtection()
                            ? DetectionStatus.TRUE
                            : DetectionStatus.FALSE
                    );

                    protections.setVBAProtectionDetected(
                        excelHandler.hasVBAProtection()
                            ? DetectionStatus.TRUE
                            : DetectionStatus.FALSE
                    );

                    // Enable unlock button
                    unlockButton.setEnabled(true);
                } catch (IOException e) {
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(frame, languageManager.getString("dialog.error.text", e.getMessage()), languageManager.getString("dialog.error.title"), JOptionPane.ERROR_MESSAGE);
                    return;
                } finally {
                    // Enable scan button
                    browseButton.setEnabled(true);
                    scanButton.setEnabled(true);
                }
            }
        });
    
        // Action listener for the unlock button to handle protection removal
        unlockButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent event) {
                String filePath = textField.getText();
                if (filePath.isEmpty()) {
                    // Show error dialog if no file is selected
                    JOptionPane.showMessageDialog(frame, languageManager.getString("dialog.noFile.text"), languageManager.getString("dialog.noFile.title"), JOptionPane.ERROR_MESSAGE);
                    return;
                }
                
                try {
                    // Disable buttons temporarily
                    browseButton.setEnabled(false);
                    scanButton.setEnabled(false);
                    unlockButton.setEnabled(false);

                    // User should only be here, if file is unencrypted
                    protections.setOpenFilePasswordRemoved(RemovalStatus.FALSE);

                    if (excelHandler.hasModifyFilePassword()) {
                        excelHandler.removeModifyFilePassword();
                        protections.setModifyFilePasswordRemoved(RemovalStatus.TRUE);
                    }

                    if (excelHandler.hasWorkbookProtection()) {
                        excelHandler.removeWorkbookProtection();
                        protections.setWorkbookProtectionRemoved(RemovalStatus.TRUE);
                    }

                    if (excelHandler.hasVBAProtection()) {
                        int response = JOptionPane.showConfirmDialog(frame,
                            languageManager.getString("dialog.vbaProtection.text"),
                            languageManager.getString("dialog.vbaProtection.title"),
                            JOptionPane.YES_NO_OPTION,
                            JOptionPane.WARNING_MESSAGE);
                    
                        if (response == JOptionPane.YES_OPTION) {
                            //protections.removeVBAProtection();
                            protections.setVBAProtectionRemoved(RemovalStatus.TRUE);
                        } else {
                            protections.setVBAProtectionRemoved(RemovalStatus.FALSE);
                        }
                    }
                
                } catch (Exception e) {
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(frame, languageManager.getString("dialog.error.text", e.getMessage()), languageManager.getString("dialog.error.title"), JOptionPane.ERROR_MESSAGE);
                    return;
                } finally {
                    // Enable buttons
                    browseButton.setEnabled(true);
                    scanButton.setEnabled(true);
                    unlockButton.setEnabled(true);
                }
            }
        });

        // Window Listener for cleaning up before exiting
        frame.addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                closeGUI();
            }
        });
    
        // Add the main panel to the frame
        frame.add(panel);
    }

    private static JMenuBar createMenuBar() {
        JMenuBar menuBar = new JMenuBar();
    
        // Menu item: File
        JMenu fileMenu = new JMenu(languageManager.getString("menu.file"));
        JMenuItem openItem = new JMenuItem(languageManager.getString("menu.file.open"));
        openItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                File selectedFile = fileManager.showOpenDialog();
                if (selectedFile != null) {
                    textField.setText(selectedFile.getAbsolutePath());
                    App.log(languageManager.getString("message.fileSelected", selectedFile.getAbsolutePath()));
                }
            }
        });

        JMenuItem exitItem = new JMenuItem(languageManager.getString("menu.file.quit"));
        exitItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                closeGUI();
            }
        });

        fileMenu.add(openItem);
        fileMenu.addSeparator();
        fileMenu.add(exitItem);
    
        // Menu item: View -> Theme
        JMenu viewMenu = new JMenu(languageManager.getString("menu.view"));
    
        JRadioButtonMenuItem lightThemeItem = new JRadioButtonMenuItem(languageManager.getString("menu.view.theme.light"));
        JRadioButtonMenuItem darkThemeItem = new JRadioButtonMenuItem(languageManager.getString("menu.view.theme.dark"));
    
        ButtonGroup themeGroup = new ButtonGroup();
        themeGroup.add(lightThemeItem);
        themeGroup.add(darkThemeItem);
    
        lightThemeItem.addActionListener(e -> {
            try {
                updateTheme(false);
                lightThemeItem.setSelected(true);
                darkThemeItem.setSelected(false);
                App.log(languageManager.getString("message.theme.light"));
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });
    
        darkThemeItem.addActionListener(e -> {
            try {
                updateTheme(true);
                lightThemeItem.setSelected(false);
                darkThemeItem.setSelected(true);
                App.log(languageManager.getString("message.theme.dark"));
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });
    
        if (isDarkTheme) {
            darkThemeItem.setSelected(true);
        } else {
            lightThemeItem.setSelected(true);
        }
    
        // Menu item: View -> Language
        JRadioButtonMenuItem germanLanguageItem = new JRadioButtonMenuItem(languageManager.getString("menu.view.language.de"));
        JRadioButtonMenuItem englishLanguageItem = new JRadioButtonMenuItem(languageManager.getString("menu.view.language.en"));
    
        ButtonGroup langaugeGroup = new ButtonGroup();
        langaugeGroup.add(germanLanguageItem);
        langaugeGroup.add(englishLanguageItem);

        germanLanguageItem.addActionListener(e -> {
            try {
                updateLanguage("de");
                germanLanguageItem.setSelected(true);
                englishLanguageItem.setSelected(false);
                App.log(languageManager.getString("message.language"));
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });
    
        englishLanguageItem.addActionListener(e -> {
            try {
                updateLanguage("en");
                germanLanguageItem.setSelected(false);
                englishLanguageItem.setSelected(true);
                App.log(languageManager.getString("message.language"));
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });

        if (languageManager.getLanguage().equals("de")) {
            germanLanguageItem.setSelected(true);
        } else {
            englishLanguageItem.setSelected(true);
        }
    
        viewMenu.add(lightThemeItem);
        viewMenu.add(darkThemeItem);
        viewMenu.addSeparator();
        viewMenu.add(germanLanguageItem);
        viewMenu.add(englishLanguageItem);
    
        // Menu item: Help
        JMenu helpMenu = new JMenu(languageManager.getString("menu.help"));

        JMenuItem gitItem = new JMenuItem(languageManager.getString("menu.help.git"));
        gitItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent arg0) {
                openUrl("https://github.com/azmke/unprotexcel");
            }
        });

        JMenuItem infoItem = new JMenuItem(languageManager.getString("menu.help.info"));
        infoItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent arg0) {
                JOptionPane.showMessageDialog(frame, languageManager.getString("dialog.info.text"), languageManager.getString("dialog.info.title"), JOptionPane.INFORMATION_MESSAGE);
            }
        });
        
        helpMenu.add(gitItem);
        helpMenu.addSeparator();
        helpMenu.add(infoItem);
    
        menuBar.add(fileMenu);
        menuBar.add(viewMenu);
        menuBar.add(helpMenu);
    
        return menuBar;
    }

    private static void openUrl(String url) {
        try {
            URI uri = new URI(url);
            Desktop desktop = Desktop.getDesktop();
            if (desktop.isSupported(Desktop.Action.BROWSE)) {
                desktop.browse(uri);
            } else {
                System.err.println("Browsing not supported on this platform.");
            }
        } catch (URISyntaxException | IOException e) {
            e.printStackTrace(); // Fehlerbehandlung
        }
    }

    public static void updateTheme(boolean darkTheme) throws UnsupportedLookAndFeelException {
        if (darkTheme) {
            UIManager.setLookAndFeel(new FlatDarkLaf());
        } else {
            UIManager.setLookAndFeel(new FlatLightLaf());
        }
        SwingUtilities.updateComponentTreeUI(frame);
        isDarkTheme = darkTheme;
    }

    public static void updateLanguage(String languageCode) {
        // Update language manager
        languageManager.setLanguage(languageCode);

        // Update menu
        menuBar = createMenuBar();
        frame.setJMenuBar(menuBar);

        // Update buttons
        browseButton.setText(languageManager.getString("button.browse"));
        scanButton.setText(languageManager.getString("button.scan"));
        unlockButton.setText(languageManager.getString("button.unlock"));

        // Update file table
        protections.updateLanguage();

        // Other stuff
        frame.setTitle(languageManager.getString("app.title"));
        frame.revalidate();
        frame.repaint();
    }

    public void showGUI() {
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }
    
    public static void closeGUI() {
        try {
            if (excelHandler != null) {
                excelHandler.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.exit(0);
        }
    }
}
