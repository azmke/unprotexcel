package com.azmke.unprotexcel.gui;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumnModel;
import javax.swing.table.DefaultTableCellRenderer;
import java.awt.*;
import java.util.ArrayList;
import java.util.List;

import com.azmke.unprotexcel.utils.LanguageManager;

public class ProtectionsTable extends JScrollPane {
    private JScrollPane scrollPane;
    private JTable table;
    private DefaultTableModel tableModel;
    private List<Protection> protections;
    private LanguageManager languageManager;

    public ProtectionsTable(LanguageManager languageManager) {
        this.languageManager = languageManager;

        protections = new ArrayList<>();

        // Initialize table
        String[] columnNames = getLocalizedColumnNames();
        tableModel = new DefaultTableModel(columnNames, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return false; // cells are not editable
            }
        };
        table = new JTable(tableModel);
        scrollPane = new JScrollPane(table);

        // Set cell renderer
        table.setDefaultRenderer(Object.class, new CustomCellRenderer());

        // Disable column reordering
        table.getTableHeader().setReorderingAllowed(false);

        // Set column widths
        setColumnWidths();
    }

    private String[] getLocalizedColumnNames() {
        return new String[] {
            languageManager.getString("table.protection.title"),
            languageManager.getString("table.detection.title"),
            languageManager.getString("table.removal.title")
        };
    }

    // Set the preferred widths of the columns
    private void setColumnWidths() {
        TableColumnModel columnModel = table.getColumnModel();
        int totalWidth = columnModel.getTotalColumnWidth();

        int columnAWidth = (int) (totalWidth * 0.8);
        int columnBWidth = (int) (totalWidth * 0.1);
        int columnCWidth = (int) (totalWidth * 0.1);

        columnModel.getColumn(0).setPreferredWidth(columnAWidth);
        columnModel.getColumn(1).setPreferredWidth(columnBWidth);
        columnModel.getColumn(2).setPreferredWidth(columnCWidth);
    }

    // Update the table based on the current protection results
    private void updateTable() {
        tableModel.setRowCount(0); // Clear existing rows
        for (Protection protection : protections) {
            Object[] rowData = {
                protection.getName(),
                protection.getDetectionStatus(),
                protection.getRemovalStatus()
            };
            tableModel.addRow(rowData);
        }
    }

    // Method to update the table when the language changes
    public void updateLanguage() {
        tableModel.setColumnIdentifiers(getLocalizedColumnNames());
        updateTable(); // Update to reflect current protections
        setColumnWidths(); // Reset column widths
        table.repaint(); // Repaint the table to reflect changes
    }

    // Method to add a new protection
    public void addProtection(Protection protection) {
        protections.add(protection);
        updateTable();
    }

    public void addProtection(String name) {
        Protection protection = new Protection(name);
        addProtection(protection);
    }

    public void updateProtection(int index, DetectionStatus detectionStatus, RemovalStatus removalStatus) {
        if (index < protections.size()) {
            Protection protection = protections.get(index);
            if (detectionStatus != null) {
                protection.setDetectionStatus(detectionStatus);
            }
            if (removalStatus != null) {
                protection.setRemovalStatus(removalStatus);
            }
            updateTable(); // Update the visual representation
        }
    }

    public void removeProtection(int index) {
        if (index < protections.size()) {
            protections.remove(index);
            updateTable(); // Update the visual representation
        }
    }

    public JScrollPane getTable() {
        return scrollPane;
    }
    
    // Custom cell renderer for coloring cells based on values
    private class CustomCellRenderer extends DefaultTableCellRenderer {
        @Override
        public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
            JLabel label = (JLabel) super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

            // Reset cell color
            label.setOpaque(true);
            label.setBackground(null);
            label.setForeground(null);

            // Se cell text and color based on value
            if (column == 1) { // Detection Status
                if (value == DetectionStatus.TRUE) {
                    label.setText(languageManager.getString("table.detection.true"));
                    label.setBackground(new Color(76, 175, 80)); // Green
                    label.setForeground(Color.WHITE);
                } else if (value == DetectionStatus.FALSE) {
                    label.setText(languageManager.getString("table.detection.false"));
                    label.setForeground(Color.GRAY);
                } else if (value == DetectionStatus.FAILED) {
                    label.setText(languageManager.getString("table.detection.failed"));
                    label.setBackground(new Color(183, 74, 74)); // Red
                    label.setForeground(Color.WHITE);
                } else {
                    label.setText(languageManager.getString("table.detection.null"));
                    label.setForeground(Color.GRAY);
                }
            } else if (column == 2) { // Removal Status
                if (value == RemovalStatus.TRUE) {
                    label.setText(languageManager.getString("table.removal.true"));
                    label.setBackground(new Color(76, 175, 80)); // Green
                    label.setForeground(Color.WHITE);
                } else if (value == RemovalStatus.FALSE) {
                    label.setText(languageManager.getString("table.removal.false"));
                    label.setForeground(Color.GRAY);
                } else if (value == RemovalStatus.FAILED) {
                    label.setText(languageManager.getString("table.removal.failed"));
                    label.setBackground(new Color(183, 74, 74)); // Red
                    label.setForeground(Color.WHITE);
                } else {
                    label.setText(languageManager.getString("table.removal.null"));
                    label.setForeground(Color.GRAY);
                }
            } else {
                label.setText(languageManager.getString(value.toString()));
            }

            return label;
        }
    }
}
