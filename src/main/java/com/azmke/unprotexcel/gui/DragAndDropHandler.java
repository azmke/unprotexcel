package com.azmke.unprotexcel.gui;

import javax.swing.*;

import com.azmke.unprotexcel.App;
import com.azmke.unprotexcel.utils.LanguageManager;

import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTargetAdapter;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.util.List;

public class DragAndDropHandler extends DropTargetAdapter {
    private final JTextField textArea;
    private final LanguageManager languageManager;

    public DragAndDropHandler(JTextField textArea, LanguageManager languageManager) {
        this.textArea = textArea;
        this.languageManager = languageManager;
    }

    @Override
    public void drop(DropTargetDropEvent dtde) {
        try {
            dtde.acceptDrop(DnDConstants.ACTION_COPY);

            @SuppressWarnings("unchecked")
            List<File> droppedFiles = (List<File>) dtde.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);

            for (File file : droppedFiles) {
                if (file.getName().endsWith(".xlsx") || file.getName().endsWith(".xls")) {
                    textArea.setText(file.getAbsolutePath());
                    App.log(languageManager.getString("message.fileSelected", file.getAbsolutePath()));
                } else {
                    String message = languageManager.getString("message.unsupportedFileType", file.getName());
                    JOptionPane.showMessageDialog(null, message, languageManager.getString("dialog.error.title"), JOptionPane.ERROR_MESSAGE);
                    App.log(message);
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}
