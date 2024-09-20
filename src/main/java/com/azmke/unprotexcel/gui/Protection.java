package com.azmke.unprotexcel.gui;

public class Protection {
    private final String name;
    private DetectionStatus detectionStatus;
    private RemovalStatus removalStatus;

    public Protection(String name) {
        this.name = name;
        this.detectionStatus = DetectionStatus.NULL;
        this.removalStatus = RemovalStatus.NULL;
    }

    public String getName() {
        return name;
    }

    public DetectionStatus getDetectionStatus() {
        return detectionStatus;
    }

    public void setDetectionStatus(DetectionStatus detectionStatus) {
        this.detectionStatus = detectionStatus;
    }

    public RemovalStatus getRemovalStatus() {
        return removalStatus;
    }

    public void setRemovalStatus(RemovalStatus removalStatus) {
        this.removalStatus = removalStatus;
    }
}
