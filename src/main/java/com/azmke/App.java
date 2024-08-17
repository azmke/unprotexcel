package com.azmke;

import com.formdev.flatlaf.FlatDarkLaf;
import com.formdev.flatlaf.FlatLightLaf;

import javax.swing.*;
import java.awt.*;
import java.awt.dnd.DropTarget;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class App {
    private static LanguageManager languageManager;
    private static FileManager fileManager;
    private static DefaultListModel<String> logModel;

    private static JFrame frame;
    private static JMenuBar menuBar;
    private static JTextField textField;
    private static JButton browseButton;
    private static JButton unlockButton;
    private static JList<String> logList;

    private static boolean isDarkTheme = true;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(App::createAndShowGUI);
    }

    private static void createAndShowGUI() {
        try {
            UIManager.setLookAndFeel(new FlatDarkLaf());
        } catch (UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }

        languageManager = new LanguageManager("en");
        fileManager = new FileManager(languageManager);
    
        frame = new JFrame(languageManager.getString("app.title"));
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(600, 300);
        frame.setMinimumSize(new Dimension(500, 300));
        frame.setMaximumSize(new Dimension(800, 600));
    
        menuBar = createMenuBar(frame);
        frame.setJMenuBar(menuBar);
    
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);
        gbc.fill = GridBagConstraints.BOTH;
    
        textField = new JTextField();
        textField.setDropTarget(new DropTarget(textField, new DragAndDropHandler(textField, languageManager)));
    
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 1.0;
        gbc.gridwidth = 1;
        panel.add(textField, gbc);
    
        browseButton = new JButton(languageManager.getString("button.browse"));
        gbc.gridx = 1;
        gbc.gridy = 0;
        gbc.weightx = 0;
        panel.add(browseButton, gbc);
    
        logModel = new DefaultListModel<>();
        logModel.addElement(languageManager.getString("message.welcome"));
        logList = new JList<>(logModel);
        JScrollPane scrollPane = new JScrollPane(logList);
    
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.gridwidth = 2;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        panel.add(scrollPane, gbc);
    
        unlockButton = new JButton(languageManager.getString("button.unlock"));
        gbc.gridx = 1;
        gbc.gridy = 2;
        gbc.gridwidth = 1;
        gbc.weightx = 0;
        gbc.weighty = 0;
        gbc.anchor = GridBagConstraints.EAST;
        panel.add(unlockButton, gbc);
    
        browseButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                fileManager.showOpenDialog(textField);
            }
        });

        unlockButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String filePath = textField.getText();
                if (filePath.isEmpty()) {
                    JOptionPane.showMessageDialog(frame, languageManager.getString("dialog.noFile.text"), languageManager.getString("dialog.noFile.title"), JOptionPane.ERROR_MESSAGE);
                    return;
                }

                new UnlockWorker(filePath, languageManager, fileManager).execute();
            }
        });

        frame.add(panel);
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }

    private static JMenuBar createMenuBar(JFrame frame) {
        JMenuBar menuBar = new JMenuBar();
    
        // Menu item: File
        JMenu fileMenu = new JMenu(languageManager.getString("menu.file"));
        JMenuItem openItem = new JMenuItem(languageManager.getString("menu.file.open"));
        openItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                fileManager.showOpenDialog(textField);
            }
        });

        JMenuItem exitItem = new JMenuItem(languageManager.getString("menu.file.quit"));
        exitItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                System.exit(0);
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
                App.log(languageManager.getString("message.theme.light"));
                UIManager.setLookAndFeel(new FlatLightLaf());
                SwingUtilities.updateComponentTreeUI(frame);
                isDarkTheme = false;
                lightThemeItem.setSelected(true);
                darkThemeItem.setSelected(false);
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });
    
        darkThemeItem.addActionListener(e -> {
            try {
                App.log(languageManager.getString("message.theme.dark"));
                UIManager.setLookAndFeel(new FlatDarkLaf());
                SwingUtilities.updateComponentTreeUI(frame);
                isDarkTheme = true;
                lightThemeItem.setSelected(false);
                darkThemeItem.setSelected(true);
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
                languageManager.setLanguage("de");
                App.log(languageManager.getString("message.language"));
                germanLanguageItem.setSelected(true);
                englishLanguageItem.setSelected(false);
                updateUI();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });
    
        englishLanguageItem.addActionListener(e -> {
            try {
                languageManager.setLanguage("en");
                App.log(languageManager.getString("message.language"));
                germanLanguageItem.setSelected(false);
                englishLanguageItem.setSelected(true);
                updateUI();
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
        JMenuItem infoItem = new JMenuItem(languageManager.getString("menu.help.info"));
        infoItem.addActionListener(e -> JOptionPane.showMessageDialog(null, languageManager.getString("dialog.info.text"), languageManager.getString("dialog.info.title"), JOptionPane.INFORMATION_MESSAGE));
        helpMenu.add(infoItem);
    
        menuBar.add(fileMenu);
        menuBar.add(viewMenu);
        menuBar.add(helpMenu);
    
        return menuBar;
    }

    public static void updateUI() {
        // Update menu
        menuBar = createMenuBar(frame);
        frame.setJMenuBar(menuBar);

        // Update buttons
        browseButton.setText(languageManager.getString("button.browse"));
        unlockButton.setText(languageManager.getString("button.unlock"));

        // Other stuff
        frame.setTitle(languageManager.getString("app.title"));
        //frame.revalidate();
        //frame.repaint();
    }

    public static void log(String message) {
        logModel.addElement(message);
        logList.ensureIndexIsVisible(logModel.getSize() - 1);
    }
}
