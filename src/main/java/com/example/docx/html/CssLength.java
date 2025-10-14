package com.example.docx.html;

import java.util.Objects;

public final class CssLength {
    private final double value;
    private final String unit;

    private CssLength(double value, String unit) {
        this.value = DocxHtmlUtils.roundLength(value);
        this.unit = unit;
    }

    static CssLength points(double value) {
        return new CssLength(value, "pt");
    }

    CssLength negate() {
        return new CssLength(-value, unit);
    }

    String css() {
        if (Math.abs(value) < 0.0005d) {
            return "0";
        }
        return DocxHtmlUtils.formatDecimal(value) + unit;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (!(o instanceof CssLength that)) {
            return false;
        }
        return Double.compare(that.value, value) == 0 && Objects.equals(unit, that.unit);
    }

    @Override
    public int hashCode() {
        return Objects.hash(value, unit);
    }
}
