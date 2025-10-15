package com.example.docx.html;

record TableRowCss(String backgroundColor) {

    boolean hasDeclarations() {
        return backgroundColor != null && !backgroundColor.isBlank();
    }

    String rowDeclarations() {
        if (!hasDeclarations()) {
            return "";
        }
        return "background-color:" + backgroundColor;
    }

    String cellCascadeDeclarations() {
        if (!hasDeclarations()) {
            return "";
        }
        return "background-color:" + backgroundColor;
    }
}
