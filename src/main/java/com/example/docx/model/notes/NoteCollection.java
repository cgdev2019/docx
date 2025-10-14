package com.example.docx.model.notes;

import com.example.docx.model.document.WordDocument;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * Container for footnotes or endnotes.
 */
public final class NoteCollection {

    public enum Type {FOOTNOTE, ENDNOTE}

    private final Type type;
    private final Map<Integer, Note> notes;

    public NoteCollection(Type type, Map<Integer, Note> notes) {
        this.type = type;
        this.notes = Collections.unmodifiableMap(new LinkedHashMap<>(notes));
    }

    public Type type() {
        return type;
    }

    public Map<Integer, Note> notes() {
        return notes;
    }

    public Optional<Note> noteById(int id) {
        return Optional.ofNullable(notes.get(id));
    }

    public static final class Note {
        private final int id;
        private final String type;
        private final List<WordDocument.Block> content;

        public Note(int id, String type, List<WordDocument.Block> content) {
            this.id = id;
            this.type = type;
            this.content = List.copyOf(content);
        }

        public int id() {
            return id;
        }

        public Optional<String> type() {
            return Optional.ofNullable(type);
        }

        public List<WordDocument.Block> content() {
            return content;
        }
    }

    public static Builder builder(Type type) {
        return new Builder(type);
    }

    public static final class Builder {
        private final Type type;
        private final Map<Integer, Note> notes = new LinkedHashMap<>();

        public Builder(Type type) {
            this.type = type;
        }

        public Builder add(Note note) {
            notes.put(note.id(), note);
            return this;
        }

        public NoteCollection build() {
            return new NoteCollection(type, notes);
        }
    }
}
