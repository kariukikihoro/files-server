package com.muriithi.filesserver.renderviaweb;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

/**
 * Main Office Document Renderer that delegates to specific renderers
 */
public class OfficeDocumentRenderer {

    private static final Logger logger = LoggerFactory.getLogger(OfficeDocumentRenderer.class);

    private static final Set<String> WORD_EXTENSIONS =
            new HashSet<>(Arrays.asList(".docx", ".doc"));

    private static final Set<String> EXCEL_EXTENSIONS =
            new HashSet<>(Arrays.asList(".xlsx", ".xls"));

    private static final Set<String> SUPPORTED_EXTENSIONS = new HashSet<>();

    private static final Set<String> SUPPORTED_FORMATS =
            new HashSet<>(Arrays.asList("html"));

    static {
        SUPPORTED_EXTENSIONS.addAll(WORD_EXTENSIONS);
        SUPPORTED_EXTENSIONS.addAll(EXCEL_EXTENSIONS);
    }

    private final ExcelDocumentRenderer excelRenderer;
    private final WordDocumentRenderer wordRenderer;

    public OfficeDocumentRenderer() {
        this.excelRenderer = new ExcelDocumentRenderer();
        this.wordRenderer = new WordDocumentRenderer();
    }

    /**
     * Renders a document to HTML by delegating to the appropriate renderer
     */
    public byte[] renderDocument(byte[] fileContent, String fileName, String targetFormat) throws DocumentRenderException {
        validateInputs(fileContent, fileName, targetFormat);

        String extension = extractExtension(fileName);

        try {
            logger.info("Starting conversion of {} to {}", fileName, targetFormat);

            if (WORD_EXTENSIONS.contains(extension)) {
                return wordRenderer.renderWordDocument(fileContent, extension, fileName);
            } else if (EXCEL_EXTENSIONS.contains(extension)) {
                return excelRenderer.renderExcelDocument(fileContent, fileName);
            }

            throw new DocumentRenderException(
                    String.format("Unsupported file extension: %s", extension));

        } catch (Exception e) {
            logger.error("Failed to convert {} to {}: {}", fileName, targetFormat, e.getMessage(), e);
            throw new DocumentRenderException("Document conversion failed", e);
        }
    }

    private void validateInputs(byte[] fileContent, String fileName, String targetFormat) throws DocumentRenderException {
        if (fileContent == null || fileContent.length == 0) {
            throw new DocumentRenderException("File content cannot be null or empty");
        }
        if (fileName == null || fileName.trim().isEmpty()) {
            throw new DocumentRenderException("File name cannot be null or empty");
        }
        if (targetFormat == null || !SUPPORTED_FORMATS.contains(targetFormat.toLowerCase())) {
            throw new DocumentRenderException("Unsupported target format: " + targetFormat);
        }

        String extension = extractExtension(fileName);
        if (!SUPPORTED_EXTENSIONS.contains(extension)) {
            throw new DocumentRenderException("Unsupported file extension: " + extension);
        }
    }

    private String extractExtension(String fileName) {
        int lastDot = fileName.lastIndexOf('.');
        if (lastDot == -1) {
            return "";
        }
        return fileName.substring(lastDot).toLowerCase();
    }

    public static class DocumentRenderException extends Exception {
        public DocumentRenderException(String message) {
            super(message);
        }

        public DocumentRenderException(String message, Throwable cause) {
            super(message, cause);
        }
    }

    public static String extractDocumentName(String fileName) {

        if (fileName == null || fileName.trim().isEmpty()) {
            return "Document";
        }
        int lastDot = fileName.lastIndexOf('.');

        String nameWithoutExt = lastDot > 0 ? fileName.substring(0, lastDot) : fileName;

        String formatted = nameWithoutExt.replace("_", " ").replace("-", " ");

        String[] words = formatted.split("\\s+");
        StringBuilder result = new StringBuilder();

        for (String word : words) {

            if (word.length() > 0) {

                if (result.length() > 0) result.append(" ");

                result.append(word.substring(0, 1).toUpperCase())
                        .append(word.substring(1).toLowerCase());
            }
        }

        return result.toString();
    }

    public static byte[] wrapWithModernStyling(byte[] htmlContent, String fileName, String baseStyle) throws Exception {
        String content = new String(htmlContent, "UTF-8");

        String documentName = extractDocumentName(fileName);

        String nameHeader = "<div class='document-name-header'><div class='document-name'>" +
                escapeHtml(documentName) + "</div></div>";

        if (content.contains("<head>")) {
            content = content.replace("<head>", "<head>" + baseStyle);
        } else if (content.contains("<html>")) {
            content = content.replace("<html>", "<html><head>" + baseStyle + "</head>");
        } else {
            content = "<!DOCTYPE html><html><head>" + baseStyle + "</head><body>" + content + "</body></html>";
        }

        if (content.contains("<body>")) {
            content = content.replace("<body>", "<body><div class='document-container'>" + nameHeader);

            content = content.replace("</body>", "</div></body>");
        }

        return content.getBytes("UTF-8");
    }

    public static String escapeHtml(String input) {
        if (input == null) return "";

        return input.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&#x27;")
                .replace("\n", "<br>");
    }
}
