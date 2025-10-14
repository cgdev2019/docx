package com.example.docx.model.metadata;

import java.util.Map;

/**
 * Extended properties extracted from {@code docProps/app.xml}.
 */
public record AppProperties(
        String application,
        String applicationVersion,
        String company,
        String manager,
        String presentationFormat,
        Integer pages,
        Integer words,
        Integer characters,
        Integer paragraphs,
        Integer lines,
        Integer slides,
        Integer notes,
        Integer totalTime,
        Boolean template,
        Boolean sharedDoc,
        Integer docSecurity,
        Boolean linksUpToDate,
        Map<String, String> extra
) {
    public AppProperties {
        extra = Map.copyOf(extra);
    }
}
