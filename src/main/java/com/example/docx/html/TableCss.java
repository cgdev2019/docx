package com.example.docx.html;

record TableCss(String backgroundColor) {

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
