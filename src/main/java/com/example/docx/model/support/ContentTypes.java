package com.example.docx.model.support;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Represents the {@code [Content_Types].xml} part.
 */
public final class ContentTypes {

    private final Map<String, String> defaults;
    private final Map<String, String> overrides;

    public ContentTypes(Map<String, String> defaults, Map<String, String> overrides) {
        this.defaults = Collections.unmodifiableMap(new LinkedHashMap<>(defaults));
        this.overrides = Collections.unmodifiableMap(new LinkedHashMap<>(overrides));
    }

    public Map<String, String> defaults() {
        return defaults;
    }

    public Map<String, String> overrides() {
        return overrides;
    }

    public String lookup(String partName) {
        return overrides.getOrDefault(partName, null);
    }
}
