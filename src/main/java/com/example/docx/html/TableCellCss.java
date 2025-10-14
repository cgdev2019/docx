package com.example.docx.html;

public record TableCellCss(String backgroundColor) {

    boolean hasColor() {
        return backgroundColor != null && !backgroundColor.isBlank();
    }

    String declarations() {
        if (!hasColor()) {
            return "";
        }
        return "background-color:" + backgroundColor;
    }
}
