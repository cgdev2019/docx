package com.example.docx.model.support;

import java.util.Arrays;

/**
 * Binary media contained in the DOCX package (images, embedded objects, etc.).
 */
public final class MediaFile {
    private final String partName;
    private final String contentType;
    private final byte[] data;

    public MediaFile(String partName, String contentType, byte[] data) {
        this.partName = partName;
        this.contentType = contentType;
        this.data = Arrays.copyOf(data, data.length);
    }

    public String partName() {
        return partName;
    }

    public String contentType() {
        return contentType;
    }

    public byte[] data() {
        return Arrays.copyOf(data, data.length);
    }
}
