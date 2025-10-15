package com.example.docx.html;

record TableCss(String backgroundColor, BorderDefinition borders) {

    boolean hasDeclarations() {
        return (backgroundColor != null && !backgroundColor.isBlank())
                || (borders != null && !borders.isEmpty());
    }

    String tableDeclarations() {
        if (!hasDeclarations()) {
            return "";
        }
        StringBuilder builder = new StringBuilder();
        if (backgroundColor != null && !backgroundColor.isBlank()) {
            builder.append("background-color:").append(backgroundColor);
        }
        if (borders != null && !borders.isEmpty()) {
            if (builder.length() > 0) {
                builder.append(';');
            }
            java.util.List<String> items = new java.util.ArrayList<>();
            borders.appendCss(items);
            builder.append(String.join(";", items));
        }
        return builder.toString();
    }

    String cellCascadeDeclarations() {
        if (backgroundColor == null || backgroundColor.isBlank()) {
            return "";
        }
        return "background-color:" + backgroundColor;
    }
}
