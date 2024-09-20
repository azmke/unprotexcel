package com.azmke.unprotexcel.workers;

import javax.swing.*;

import com.azmke.unprotexcel.App;
import com.azmke.unprotexcel.handlers.ExcelHandler;
import com.azmke.unprotexcel.utils.LanguageManager;

public class UnlockWorker extends SwingWorker<Void, String> {
    private final String filePath;
    private final LanguageManager languageManager;

    public UnlockWorker(String filePath, LanguageManager languageManager) {
        this.filePath = filePath;
        this.languageManager = languageManager;
    }

    @Override
    protected Void doInBackground() {
        try {
            ExcelHandler excelHandler = new ExcelHandler();

            excelHandler.open(filePath);

            App.log(String.valueOf(ExcelHandler.hasOpenFilePassword(filePath)));
            App.log(String.valueOf(excelHandler.hasModifyFilePassword()));
            App.log(String.valueOf(excelHandler.hasWorkbookProtection()));
            App.log(String.valueOf(excelHandler.hasVBAProtection()));

            excelHandler.close();

            App.log("Success :)");

        } catch (Exception ex) {
            ex.printStackTrace();
            App.log(languageManager.getString("message.failed", ex.getMessage()));
            JOptionPane.showMessageDialog(null, languageManager.getString("dialog.error.text", ex.getMessage()), languageManager.getString("dialog.error.title"), JOptionPane.ERROR_MESSAGE);
        }
        return null;
    }
}
