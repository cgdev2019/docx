package com.example.docx.parser;

import com.example.docx.io.DocxArchive;
import com.example.docx.model.DocxPackage;
import com.example.docx.model.document.WordDocument;
import com.example.docx.model.notes.NoteCollection;
import com.example.docx.model.relationship.RelationshipSet;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.IOException;
import java.io.InputStream;

/**
 * Parses footnote and endnote parts.
 */
public final class NotesParser {

    private final RelationshipsParser relationshipsParser = new RelationshipsParser();
    private final ParsingContext context;

    public NotesParser(ParsingContext context) {
        this.context = context;
    }

    public NoteCollection parse(DocxArchive archive, String partName, NoteCollection.Type type, DocxPackage.Builder builder) throws IOException {
        if (!archive.exists(partName)) {
            return null;
        }
        RelationshipSet rels = relationshipsParser.parse(archive, toRels(partName));
        if (rels != null) {
            builder.relationshipForPart(partName, rels);
        }
        try (InputStream input = archive.open(partName)) {
            Document document = XmlUtils.parse(input);
            Element root = document.getDocumentElement();
            NoteCollection.Builder notes = NoteCollection.builder(type);
            for (Element noteEl : XmlUtils.childElements(root)) {
                if (!Namespaces.WORD_MAIN.equals(noteEl.getNamespaceURI())) {
                    throw ParserSupport.unknownElement("note", noteEl);
                }
                String idAttr = noteEl.getAttributeNS(Namespaces.WORD_MAIN, "id");
                if (idAttr == null || idAttr.isEmpty()) {
                    continue;
                }
                int id = Integer.parseInt(idAttr);
                String noteType = noteEl.getAttributeNS(Namespaces.WORD_MAIN, "type");
                noteType = noteType.isEmpty() ? null : noteType;
                notes.add(new NoteCollection.Note(id, noteType, context.blockParser.parseChildBlocks(noteEl, rels)));
            }
            return notes.build();
        }
    }

    private String toRels(String path) {
        int index = path.lastIndexOf('/');
        String directory = index == -1 ? "" : path.substring(0, index + 1);
        String fileName = index == -1 ? path : path.substring(index + 1);
        return directory + "_rels/" + fileName + ".rels";
    }
}
