package com.azmke.unprotexcel.utils;

import java.text.MessageFormat;
import java.util.Locale;
import java.util.ResourceBundle;

public class LanguageManager {
    private ResourceBundle resourceBundle;
    private String currentLanguage;

    public LanguageManager() {
        setLanguage("en");
    }

    public LanguageManager(String languageCode) {
        setLanguage(languageCode);
    }

    public void setLanguage(String languageCode) {
        currentLanguage = languageCode;
        Locale locale = Locale.of(languageCode);
        resourceBundle = ResourceBundle.getBundle("messages", locale);
    }

    public String getLanguage() {
        return currentLanguage;
    }

    public String getString(String key) {
        if (resourceBundle.containsKey(key)) {
            return resourceBundle.getString(key);
        } else {
            return key;
        }
    }

    public String getString(String key, Object... params) {
        if (resourceBundle.containsKey(key)) {
            String message = resourceBundle.getString(key);
            return MessageFormat.format(message, params);
        } else {
            return key;
        }
    }
}
