package com.example.docx.model.relationship;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

/**
 * Represents a relationship part ({@code _rels/*.rels}).
 */
public final class RelationshipSet {

    private final Map<String, Relationship> relationships;

    public RelationshipSet(Map<String, Relationship> relationships) {
        this.relationships = Collections.unmodifiableMap(new LinkedHashMap<>(relationships));
    }

    public Map<String, Relationship> relationships() {
        return relationships;
    }

    public Optional<Relationship> byId(String id) {
        return Optional.ofNullable(relationships.get(id));
    }

    public static final class Relationship {
        private final String id;
        private final String type;
        private final String target;
        private final String targetMode;

        public Relationship(String id, String type, String target, String targetMode) {
            this.id = id;
            this.type = type;
            this.target = target;
            this.targetMode = targetMode;
        }

        public String id() {
            return id;
        }

        public String type() {
            return type;
        }

        public String target() {
            return target;
        }

        public Optional<String> targetMode() {
            return Optional.ofNullable(targetMode);
        }
    }
}
