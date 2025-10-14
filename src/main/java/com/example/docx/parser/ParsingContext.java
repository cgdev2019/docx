package com.example.docx.parser;

/**
 * Aggregates parser instances to enforce consistent handling and strict validation.
 */
public final class ParsingContext {

    final RunParser runParser;
    final ParagraphParser paragraphParser;
    final TableParser tableParser;
    final BlockParser blockParser;

    public ParsingContext() {
        this.runParser = new RunParser(this);
        this.paragraphParser = new ParagraphParser(this);
        this.tableParser = new TableParser(this);
        this.blockParser = new BlockParser(this);
    }
}
