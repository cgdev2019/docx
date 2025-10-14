# DOCX Reader

Library Java 21 capable of inspecting the full content of Office Open XML `DOCX` packages without relying on external dependencies.

## Build & tests

```bash
mvn test
```

## Usage

```java
var reader = new DocxReader();
DocxPackage pkg = reader.read(Path.of("samples", "demo.docx"));
String title = pkg.coreProperties().flatMap(CoreProperties::title).orElse("(untitled)");
```

See `DocxReaderTest` for additional usage examples.
