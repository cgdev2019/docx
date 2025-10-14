package com.example.docx.io;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;
import java.util.HashSet;
import java.util.Objects;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * Abstraction over the physical DOCX container (zip archive or extracted directory).
 */
public interface DocxArchive extends AutoCloseable {

    InputStream open(String partName) throws IOException;

    boolean exists(String partName);

    Set<String> list(String prefix) throws IOException;

    @Override
    void close() throws IOException;

    static DocxArchive open(Path path) throws IOException {
        Objects.requireNonNull(path, "path");
        if (Files.isDirectory(path)) {
            return new DirectoryArchive(path);
        }
        return new ZipDocxArchive(path);
    }

    final class DirectoryArchive implements DocxArchive {
        private final Path root;

        public DirectoryArchive(Path root) {
            this.root = root;
        }

        @Override
        public InputStream open(String partName) throws IOException {
            Path resolved = root.resolve(partName.replace('/', java.io.File.separatorChar));
            return Files.newInputStream(resolved);
        }

        @Override
        public boolean exists(String partName) {
            Path resolved = root.resolve(partName.replace('/', java.io.File.separatorChar));
            return Files.exists(resolved);
        }

        @Override
        public Set<String> list(String prefix) throws IOException {
            Set<String> result = new HashSet<>();
            Path resolved = root.resolve(prefix.replace('/', java.io.File.separatorChar));
            if (!Files.exists(resolved)) {
                return Collections.emptySet();
            }
            try (var stream = java.nio.file.Files.walk(resolved)) {
                stream.forEach(path -> {
                    if (java.nio.file.Files.isRegularFile(path)) {
                        String relative = root.relativize(path).toString().replace(java.io.File.separatorChar, '/');
                        result.add(relative);
                    }
                });
            }
            return result;
        }

        @Override
        public void close() {
            // nothing to close
        }
    }

    final class ZipDocxArchive implements DocxArchive {
        private final ZipFile zipFile;

        public ZipDocxArchive(Path path) throws IOException {
            this.zipFile = new ZipFile(path.toFile(), ZipFile.OPEN_READ, StandardCharsets.UTF_8);
        }

        @Override
        public InputStream open(String partName) throws IOException {
            ZipEntry entry = zipFile.getEntry(partName);
            if (entry == null) {
                throw new IOException("Missing part: " + partName);
            }
            return zipFile.getInputStream(entry);
        }

        @Override
        public boolean exists(String partName) {
            return zipFile.getEntry(partName) != null;
        }

        @Override
        public Set<String> list(String prefix) {
            Set<String> result = new HashSet<>();
            zipFile.stream()
                    .map(ZipEntry::getName)
                    .filter(name -> name.startsWith(prefix))
                    .forEach(result::add);
            return result;
        }

        @Override
        public void close() throws IOException {
            zipFile.close();
        }
    }
}
