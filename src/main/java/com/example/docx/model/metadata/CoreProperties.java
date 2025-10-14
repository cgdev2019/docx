package com.example.docx.model.metadata;

import java.time.Instant;
import java.util.Map;
import java.util.Objects;

/**
 * Core document metadata extracted from {@code docProps/core.xml}.
 */
public record CoreProperties(
        String title,
        String subject,
        String creator,
        String description,
        String keywords,
        String lastModifiedBy,
        String revision,
        Instant created,
        Instant modified,
        String category,
        String contentStatus,
        String language,
        String identifier,
        String version,
        Map<String, String> extra
) {
    public CoreProperties {
        extra = Map.copyOf(extra);
    }

    public static Builder builder() {
        return new Builder();
    }

    public static final class Builder {
        private String title;
        private String subject;
        private String creator;
        private String description;
        private String keywords;
        private String lastModifiedBy;
        private String revision;
        private Instant created;
        private Instant modified;
        private String category;
        private String contentStatus;
        private String language;
        private String identifier;
        private String version;
        private Map<String, String> extra = Map.of();

        public Builder title(String value) {
            this.title = value;
            return this;
        }

        public Builder subject(String value) {
            this.subject = value;
            return this;
        }

        public Builder creator(String value) {
            this.creator = value;
            return this;
        }

        public Builder description(String value) {
            this.description = value;
            return this;
        }

        public Builder keywords(String value) {
            this.keywords = value;
            return this;
        }

        public Builder lastModifiedBy(String value) {
            this.lastModifiedBy = value;
            return this;
        }

        public Builder revision(String value) {
            this.revision = value;
            return this;
        }

        public Builder created(Instant value) {
            this.created = value;
            return this;
        }

        public Builder modified(Instant value) {
            this.modified = value;
            return this;
        }

        public Builder category(String value) {
            this.category = value;
            return this;
        }

        public Builder contentStatus(String value) {
            this.contentStatus = value;
            return this;
        }

        public Builder language(String value) {
            this.language = value;
            return this;
        }

        public Builder identifier(String value) {
            this.identifier = value;
            return this;
        }

        public Builder version(String value) {
            this.version = value;
            return this;
        }

        public Builder extra(Map<String, String> value) {
            this.extra = Objects.requireNonNull(value, "value");
            return this;
        }

        public CoreProperties build() {
            return new CoreProperties(title, subject, creator, description, keywords, lastModifiedBy, revision,
                    created, modified, category, contentStatus, language, identifier, version,
                    extra == null ? Map.of() : extra);
        }
    }
}
