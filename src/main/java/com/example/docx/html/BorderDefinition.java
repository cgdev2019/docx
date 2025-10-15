package com.example.docx.html;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

final class BorderDefinition {
    private final BorderEdge top;
    private final BorderEdge right;
    private final BorderEdge bottom;
    private final BorderEdge left;

    private BorderDefinition(BorderEdge top, BorderEdge right, BorderEdge bottom, BorderEdge left) {
        this.top = top;
        this.right = right;
        this.bottom = bottom;
        this.left = left;
    }

    static BorderDefinition empty() {
        return new BorderDefinition(null, null, null, null);
    }

    static BorderDefinition of(BorderEdge top, BorderEdge right, BorderEdge bottom, BorderEdge left) {
        if (top == null && right == null && bottom == null && left == null) {
            return empty();
        }
        return new BorderDefinition(top, right, bottom, left);
    }

    boolean isEmpty() {
        return top == null && right == null && bottom == null && left == null;
    }

    BorderEdge top() {
        return top;
    }

    BorderEdge right() {
        return right;
    }

    BorderEdge bottom() {
        return bottom;
    }

    BorderEdge left() {
        return left;
    }

    BorderDefinition overrideWith(BorderDefinition override) {
        if (override == null || override.isEmpty()) {
            return this;
        }
        BorderEdge newTop = override.top != null ? override.top : this.top;
        BorderEdge newRight = override.right != null ? override.right : this.right;
        BorderEdge newBottom = override.bottom != null ? override.bottom : this.bottom;
        BorderEdge newLeft = override.left != null ? override.left : this.left;
        return new BorderDefinition(newTop, newRight, newBottom, newLeft);
    }

    BorderDefinition fillMissing(BorderDefinition fallback) {
        if (fallback == null || fallback.isEmpty()) {
            return this;
        }
        BorderEdge newTop = this.top != null ? this.top : fallback.top;
        BorderEdge newRight = this.right != null ? this.right : fallback.right;
        BorderEdge newBottom = this.bottom != null ? this.bottom : fallback.bottom;
        BorderEdge newLeft = this.left != null ? this.left : fallback.left;
        return new BorderDefinition(newTop, newRight, newBottom, newLeft);
    }

    void appendCss(List<String> declarations) {
        if (top != null) {
            String css = top.toCss("border-top");
            if (css != null) {
                declarations.add(css);
            }
        }
        if (right != null) {
            String css = right.toCss("border-right");
            if (css != null) {
                declarations.add(css);
            }
        }
        if (bottom != null) {
            String css = bottom.toCss("border-bottom");
            if (css != null) {
                declarations.add(css);
            }
        }
        if (left != null) {
            String css = left.toCss("border-left");
            if (css != null) {
                declarations.add(css);
            }
        }
    }

    BorderDefinition withTop(BorderEdge edge) {
        return new BorderDefinition(edge, right, bottom, left);
    }

    BorderDefinition withRight(BorderEdge edge) {
        return new BorderDefinition(top, edge, bottom, left);
    }

    BorderDefinition withBottom(BorderEdge edge) {
        return new BorderDefinition(top, right, edge, left);
    }

    BorderDefinition withLeft(BorderEdge edge) {
        return new BorderDefinition(top, right, bottom, edge);
    }

    BorderDefinition stripSide(String side) {
        return switch (Objects.requireNonNull(side)) {
            case "top" -> new BorderDefinition(null, right, bottom, left);
            case "right" -> new BorderDefinition(top, null, bottom, left);
            case "bottom" -> new BorderDefinition(top, right, null, left);
            case "left" -> new BorderDefinition(top, right, bottom, null);
            default -> this;
        };
    }

    boolean hasSide(String side) {
        return switch (Objects.requireNonNull(side)) {
            case "top" -> top != null;
            case "right" -> right != null;
            case "bottom" -> bottom != null;
            case "left" -> left != null;
            default -> false;
        };
    }

    BorderEdge edge(String side) {
        return switch (Objects.requireNonNull(side)) {
            case "top" -> top;
            case "right" -> right;
            case "bottom" -> bottom;
            case "left" -> left;
            default -> null;
        };
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (!(o instanceof BorderDefinition that)) {
            return false;
        }
        return Objects.equals(top, that.top)
                && Objects.equals(right, that.right)
                && Objects.equals(bottom, that.bottom)
                && Objects.equals(left, that.left);
    }

    @Override
    public int hashCode() {
        return Objects.hash(top, right, bottom, left);
    }

    @Override
    public String toString() {
        List<String> chunks = new ArrayList<>();
        appendCss(chunks);
        return String.join(";", chunks);
    }

    static final class BorderEdge {
        private final CssLength width;
        private final String style;
        private final String color;

        BorderEdge(CssLength width, String style, String color) {
            this.width = width;
            this.style = style;
            this.color = color == null || color.isBlank() ? null : color;
        }

        boolean isVisible() {
            return width != null && style != null && !"none".equals(style);
        }

        String toCss(String property) {
            if (!isVisible()) {
                return property + ":0";
            }
            String effectiveColor = color == null ? "currentColor" : color;
            return property + ':' + width.css() + ' ' + style + ' ' + effectiveColor;
        }

        CssLength width() {
            return width;
        }

        String style() {
            return style;
        }

        String color() {
            return color;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) {
                return true;
            }
            if (!(o instanceof BorderEdge that)) {
                return false;
            }
            return Objects.equals(width, that.width)
                    && Objects.equals(style, that.style)
                    && Objects.equals(color, that.color);
        }

        @Override
        public int hashCode() {
            return Objects.hash(width, style, color);
        }
    }
}
