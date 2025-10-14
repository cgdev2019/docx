package com.example.docx.model.metadata;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

/**
 * Custom properties contained in {@code docProps/custom.xml}.
 */
public final class CustomProperties {

    private final List<CustomProperty> properties;

    public CustomProperties(List<CustomProperty> properties) {
        this.properties = List.copyOf(properties);
    }

    public List<CustomProperty> properties() {
        return properties;
    }

    public CustomProperty property(String name) {
        return properties.stream()
                .filter(prop -> prop.name().equals(name))
                .findFirst()
                .orElse(null);
    }

    public static Builder builder() {
        return new Builder();
    }

    public static final class Builder {
        private final List<CustomProperty> properties = new ArrayList<>();

        public Builder add(CustomProperty property) {
            properties.add(Objects.requireNonNull(property, "property"));
            return this;
        }

        public CustomProperties build() {
            return new CustomProperties(Collections.unmodifiableList(new ArrayList<>(properties)));
        }
    }

    public enum ValueType {
        STRING,
        BOOLEAN,
        INTEGER,
        FLOAT,
        DOUBLE,
        DECIMAL,
        DATE_TIME,
        FILETIME,
        BINARY
    }

    public record CustomProperty(String name, ValueType valueType, Object value) {
        public CustomProperty {
            Objects.requireNonNull(name, "name");
            Objects.requireNonNull(valueType, "valueType");
        }
    }
}
