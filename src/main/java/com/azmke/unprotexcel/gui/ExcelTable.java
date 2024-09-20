package com.azmke.unprotexcel.gui;

import com.azmke.unprotexcel.utils.LanguageManager;

public class ExcelTable extends ProtectionsTable {

    public ExcelTable(LanguageManager languageManager) {
        super(languageManager);

        addProtection("table.protection.openFile");
        addProtection("table.protection.modifyFile");
        addProtection("table.protection.workbookProtection");
        addProtection("table.protection.sheetProtection");
        addProtection("table.protection.rangeProtection");
        addProtection("table.protection.vbaProject");
    }

    public void setOpenFilePasswordDetected(DetectionStatus detectionStatus) {
        updateProtection(0, detectionStatus, null);
    }

    public void setOpenFilePasswordRemoved(RemovalStatus removalStatus) {
        updateProtection(0, null, removalStatus);
    }

    public void setModifyFilePasswordDetected(DetectionStatus detectionStatus) {
        updateProtection(1, detectionStatus, null);
    }

    public void setModifyFilePasswordRemoved(RemovalStatus removalStatus) {
        updateProtection(1, null, removalStatus);
    }

    public void setWorkbookProtectionDetected(DetectionStatus detectionStatus) {
        updateProtection(2, detectionStatus, null);
    }

    public void setWorkbookProtectionRemoved(RemovalStatus removalStatus) {
        updateProtection(2, null, removalStatus);
    }

    public void setSheetProtectionDetected(DetectionStatus detectionStatus) {
        updateProtection(3, detectionStatus, null);
    }

    public void setSheetProtectionRemoved(RemovalStatus removalStatus) {
        updateProtection(3, null, removalStatus);
    }

    public void setRangeProtectionDetected(DetectionStatus detectionStatus) {
        updateProtection(4, detectionStatus, null);
    }

    public void setRangeProtectionRemoved(RemovalStatus removalStatus) {
        updateProtection(4, null, removalStatus);
    }

    public void setVBAProtectionDetected(DetectionStatus detectionStatus) {
        updateProtection(5, detectionStatus, null);
    }

    public void setVBAProtectionRemoved(RemovalStatus removalStatus) {
        updateProtection(5, null, removalStatus);
    }

    public void resetDetectionStatus() {
        for (int i = 0; i <= 5; i++) {
            updateProtection(i, DetectionStatus.NULL, null);
        }
    }

    public void resetRemovalStatus() {
        for (int i = 0; i <= 5; i++) {
            updateProtection(i, null, RemovalStatus.NULL);
        }
    }
}
