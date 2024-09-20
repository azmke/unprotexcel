package com.azmke.unprotexcel;

import javax.swing.*;

import com.azmke.unprotexcel.gui.GUI;

public class App {
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            GUI gui = new GUI();
            gui.showGUI();
        });
    }

    public static void log(String message) {
        System.out.println(message);
    }
}
