package com.example.docx.model.numbering;

import org.w3c.dom.Element;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * Parsed representation of {@code word/numbering.xml}.
 */
public final class NumberingDefinitions {

    private final Map<Integer, AbstractNumbering> abstractNumberings;
    private final Map<Integer, NumberingInstance> numberingInstances;

    public NumberingDefinitions(Map<Integer, AbstractNumbering> abstractNumberings,
                                Map<Integer, NumberingInstance> numberingInstances) {
        this.abstractNumberings = Map.copyOf(abstractNumberings);
        this.numberingInstances = Map.copyOf(numberingInstances);
    }

    public Optional<AbstractNumbering> abstractNumbering(int id) {
        return Optional.ofNullable(abstractNumberings.get(id));
    }

    public Optional<NumberingInstance> numberingInstance(int id) {
        return Optional.ofNullable(numberingInstances.get(id));
    }

    public Map<Integer, AbstractNumbering> abstractNumberings() {
        return abstractNumberings;
    }

    public Map<Integer, NumberingInstance> numberingInstances() {
        return numberingInstances;
    }

    public static final class AbstractNumbering {
        private final int id;
        private final Map<Integer, Level> levels;
        private final Element raw;

        public AbstractNumbering(int id, Map<Integer, Level> levels, Element raw) {
            this.id = id;
            this.levels = Map.copyOf(levels);
            this.raw = raw;
        }

        public int id() {
            return id;
        }

        public Map<Integer, Level> levels() {
            return levels;
        }

        public Optional<Element> raw() {
            return Optional.ofNullable(raw);
        }
    }

    public static final class NumberingInstance {
        private final int id;
        private final int abstractNumId;
        private final Map<Integer, LevelOverride> levelOverrides;
        private final Element raw;

        public NumberingInstance(int id, int abstractNumId, Map<Integer, LevelOverride> levelOverrides, Element raw) {
            this.id = id;
            this.abstractNumId = abstractNumId;
            this.levelOverrides = Map.copyOf(levelOverrides);
            this.raw = raw;
        }

        public int id() {
            return id;
        }

        public int abstractNumId() {
            return abstractNumId;
        }

        public Map<Integer, LevelOverride> levelOverrides() {
            return levelOverrides;
        }

        public Optional<Element> raw() {
            return Optional.ofNullable(raw);
        }
    }

    public static final class Level {
        private final int level;
        private final String numberingFormat;
        private final String levelText;
        private final Integer start;
        private final Integer restart;
        private final Element raw;

        public Level(int level, String numberingFormat, String levelText, Integer start, Integer restart, Element raw) {
            this.level = level;
            this.numberingFormat = numberingFormat;
            this.levelText = levelText;
            this.start = start;
            this.restart = restart;
            this.raw = raw;
        }

        public int level() {
            return level;
        }

        public Optional<String> numberingFormat() {
            return Optional.ofNullable(numberingFormat);
        }

        public Optional<String> levelText() {
            return Optional.ofNullable(levelText);
        }

        public Optional<Integer> start() {
            return Optional.ofNullable(start);
        }

        public Optional<Integer> restart() {
            return Optional.ofNullable(restart);
        }

        public Optional<Element> raw() {
            return Optional.ofNullable(raw);
        }
    }

    public static final class LevelOverride {
        private final int level;
        private final Integer startOverride;
        private final Level levelDefinition;

        public LevelOverride(int level, Integer startOverride, Level levelDefinition) {
            this.level = level;
            this.startOverride = startOverride;
            this.levelDefinition = levelDefinition;
        }

        public int level() {
            return level;
        }

        public Optional<Integer> startOverride() {
            return Optional.ofNullable(startOverride);
        }

        public Optional<Level> levelDefinition() {
            return Optional.ofNullable(levelDefinition);
        }
    }

    public static Builder builder() {
        return new Builder();
    }

    public static final class Builder {
        private final Map<Integer, AbstractNumbering> abstractNumberings = new LinkedHashMap<>();
        private final Map<Integer, NumberingInstance> numberingInstances = new LinkedHashMap<>();

        public Builder addAbstractNumbering(AbstractNumbering numbering) {
            abstractNumberings.put(numbering.id(), numbering);
            return this;
        }

        public Builder addNumberingInstance(NumberingInstance instance) {
            numberingInstances.put(instance.id(), instance);
            return this;
        }

        public NumberingDefinitions build() {
            return new NumberingDefinitions(abstractNumberings, numberingInstances);
        }
    }
}
