package com.example.docx.html;

import com.example.docx.model.relationship.RelationshipSet;

import java.util.Optional;

final class HyperlinkResolver {
    private final RelationshipSet relationships;

    HyperlinkResolver(RelationshipSet relationships) {
        this.relationships = relationships;
    }

    Optional<String> resolve(String relationshipId, String anchor) {
        String target = null;
        if (relationshipId != null && relationships != null) {
            RelationshipSet.Relationship rel = relationships.byId(relationshipId).orElse(null);
            if (rel != null) {
                target = rel.target();
                if (rel.targetMode().filter(mode -> "External".equalsIgnoreCase(mode)).isEmpty()
                        && target != null && !target.startsWith("#") && !target.startsWith("http")) {
                    target = sanitizeInternalTarget(target);
                }
            }
        }
        if (anchor != null && !anchor.isBlank()) {
            if (target == null || target.isBlank()) {
                target = "#" + anchor;
            } else if (!target.contains("#")) {
                target = target + "#" + anchor;
            }
        }
        return Optional.ofNullable(target);
    }

    private String sanitizeInternalTarget(String target) {
        if (target == null) {
            return null;
        }
        if (target.startsWith("/")) {
            return target;
        }
        if (target.startsWith("#")) {
            return target;
        }
        return target.replace('\\', '/');
    }
}
