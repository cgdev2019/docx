package com.example.docx;

/**
 * Runtime exception thrown when the library fails to read or interpret a DOCX file.
 */
public class DocxException extends RuntimeException {

    public DocxException(String message) {
        super(message);
    }

    public DocxException(String message, Throwable cause) {
        super(message, cause);
    }
}
