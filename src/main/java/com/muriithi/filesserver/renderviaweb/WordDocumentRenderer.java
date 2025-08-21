package com.muriithi.filesserver.renderviaweb;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import static com.muriithi.filesserver.renderviaweb.OfficeDocumentRenderer.*;

public class WordDocumentRenderer {

    private static final Logger logger = LoggerFactory.getLogger(OfficeDocumentRenderer.class);

    private static final Set<String> WORD_EXTENSIONS =
            new HashSet<>(Arrays.asList(".docx", ".doc"));

    private static final Set<String> EXCEL_EXTENSIONS =
            new HashSet<>(Arrays.asList(".xlsx", ".xls"));

    private static final Set<String> SUPPORTED_EXTENSIONS = new HashSet<>();

    private static final Set<String> SUPPORTED_FORMATS =
            new HashSet<>(Arrays.asList("html"));

    private static final String BASE_STYLE;

    static {

        StringBuilder sb = new StringBuilder();

        sb.append("<style>")

                .append("* { box-sizing: border-box; }")
                .append("body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; ")
                .append("line-height: 1.6; color: #2c3e50; margin: 0; padding: 0; background: linear-gradient(135deg, #f5f7fa 0%, #e8eaed 100%); min-height: 100vh; }")

                /**
                 * Main container with modern design
                 * */

                .append(".document-container { max-width: 98%; margin: 10px auto; padding: 20px; ")
                .append("background: rgba(255, 255, 255, 0.95); backdrop-filter: blur(10px); border-radius: 12px; ")
                .append("box-shadow: 0 10px 30px rgba(0,0,0,0.08), 0 0 0 1px rgba(255,255,255,0.2); }")

                /**
                 * Document name header styling
                 * */

                .append(".document-name-header { ")
                .append("background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); ")
                .append("padding: 15px 20px; ")
                .append("border-radius: 8px; ")
                .append("margin-bottom: 15px; ")
                .append("border-left: 4px solid #A90C2B; ")
                .append("}")
                .append(".document-name { ")
                .append("font-size: 1.1em; ")
                .append("font-weight: 600; ")
                .append("color: #2c3e50; ")
                .append("}")

                /**
                 * Typography improvements
                 * */

                .append(".document-title { font-size: 2.2em; font-weight: 700; color: #2c3e50; text-align: center; ")
                .append("margin-bottom: 25px; background: linear-gradient(135deg, #A90C2B   0%, #8B0A23 100%); ")
                .append("-webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }")

                .append(".paragraph { margin: 12px 0; font-size: 1.1em; line-height: 1.8; }")


                /**
                 * Simple table styling for Word documents
                 * */

                .append(".document-name-header { ")
                .append("background: #A90C2B ; ")
                .append("color: white; ")
                .append("text-align: center; ")
                .append("padding: clamp(12px, 3vw, 20px); ")
                .append("border-radius: 8px; ")
                .append("margin: 0 auto 20px; ")
                .append("width: 100%; ")
                .append("max-width: 800px; ")
                .append("border-left: none; ")
                .append("box-shadow: 0 2px 8px rgba(0,0,0,0.1); ")
                .append("}")
                .append(".document-name { ")
                .append("font-size: clamp(1.1rem, 4vw, 1.5rem); ")
                .append("font-weight: 600; ")
                .append("color: white; ")
                .append("letter-spacing: 0.5px; ")
                .append("line-height: 1.3; ")
                .append("}")

                /**
                 * Responsive design
                 * */

                .append("@media (max-width: 768px) {")
                .append("  .document-container { margin: 5px; padding: 10px; }")
                .append("  .document-title { font-size: 1.8em; }")
                .append("  .document-name-header { padding: 10px 15px; }")
                .append("  .document-name { font-size: 1em; }")
                .append("  table { font-size: 0.9em; }")
                .append("  th, td { padding: 6px 8px; }")
                .append("}")

                .append("</style>");

        BASE_STYLE = sb.toString();
    }

    static {
        SUPPORTED_EXTENSIONS.addAll(WORD_EXTENSIONS);
        SUPPORTED_EXTENSIONS.addAll(EXCEL_EXTENSIONS);
    }


    public byte[] renderWordDocument(byte[] content, String extension, String fileName) throws Exception {
        return ".docx".equals(extension) ? convertDocxToHtml(content, fileName) : convertDocToHtml(content, fileName);
    }

    private byte[] convertDocxToHtml(byte[] docxContent, String fileName) throws Exception {

        InputStream is = new ByteArrayInputStream(docxContent);

        try {

            XWPFDocument document = new XWPFDocument(is);

            try {

                StringBuilder html = new StringBuilder();
                html.append("<!DOCTYPE html><html><head><meta charset='UTF-8'>")
                        .append("<title>Document</title>")
                        .append(BASE_STYLE)
                        .append("</head><body>");

                html.append("<div class='document-container'>");
                html.append("<h1 class='document-title'>Word Document</h1>");


                String documentName = extractDocumentName(fileName);
                html.append("<div class='document-name-header'>");
                html.append("<div class='document-name'>").append(escapeHtml(documentName)).append("</div>");
                html.append("</div>");


                List<IBodyElement> elements = document.getBodyElements();
                for (IBodyElement element : elements) {
                    if (element instanceof XWPFParagraph) {
                        processParagraph(html, (XWPFParagraph) element);
                    } else if (element instanceof XWPFTable) {
                        processWordTable(html, (XWPFTable) element);
                    }
                }

                html.append("</div></body></html>");
                return html.toString().getBytes("UTF-8");
            } finally {
            }
        } finally {

            try {

                is.close();
            } catch (Exception e) {

                logger.warn("Error closing input stream", e);
            }
        }
    }

    private void processParagraph(StringBuilder html, XWPFParagraph paragraph) {
        if (paragraph.getText().trim().isEmpty()) {
            return;
        }

        String style = " class='paragraph'";
        html.append("<p").append(style).append(">");

        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);
            if (text != null) {
                text = escapeHtml(text);
                text = applyRunFormatting(text, run);
                html.append(text);
            }
        }

        html.append("</p>");
    }

    private String applyRunFormatting(String text, XWPFRun run) {
        if (run.isBold()) {
            text = "<strong>" + text + "</strong>";
        }
        if (run.isItalic()) {
            text = "<em>" + text + "</em>";
        }
        if (run.getUnderline() != null && run.getUnderline() != UnderlinePatterns.NONE) {
            text = "<u>" + text + "</u>";
        }

        return text;
    }

    private void processWordTable(StringBuilder html, XWPFTable table) {
        html.append("<table>");

        boolean isFirstRow = true;
        for (XWPFTableRow row : table.getRows()) {
            html.append("<tr>");

            for (XWPFTableCell cell : row.getTableCells()) {
                String tag = isFirstRow ? "th" : "td";
                html.append("<").append(tag).append(">");

                StringBuilder cellContent = new StringBuilder();
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    String cellText = paragraph.getText().trim();
                    if (!cellText.isEmpty()) {
                        if (cellContent.length() > 0) {
                            cellContent.append(" ");
                        }
                        cellContent.append(cellText);
                    }
                }

                String fullText = cellContent.toString().trim();
                html.append(escapeHtml(fullText));
                html.append("</").append(tag).append(">");
            }

            html.append("</tr>");
            isFirstRow = false;
        }

        html.append("</table>");
    }

    private byte[] convertDocToHtml(byte[] docContent, String fileName) throws Exception {
        InputStream is = new ByteArrayInputStream(docContent);
        try {
            HWPFDocument doc = new HWPFDocument(is);
            try {
                Document htmlDocument = DocumentBuilderFactory.newInstance()
                        .newDocumentBuilder()
                        .newDocument();

                WordToHtmlConverter converter = new WordToHtmlConverter(htmlDocument);
                converter.processDocument(doc);

                ByteArrayOutputStream out = new ByteArrayOutputStream();
                try {
                    DOMSource domSource = new DOMSource(converter.getDocument());
                    StreamResult streamResult = new StreamResult(out);

                    Transformer transformer = TransformerFactory.newInstance().newTransformer();
                    transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
                    transformer.setOutputProperty(OutputKeys.INDENT, "yes");
                    transformer.setOutputProperty(OutputKeys.METHOD, "html");
                    transformer.transform(domSource, streamResult);
                    return wrapWithModernStyling(out.toByteArray(), fileName, BASE_STYLE);
                } finally {
                    try {
                        out.close();
                    } catch (Exception e) {
                        logger.warn("Error closing output stream", e);
                    }
                }
            } finally {
            }
        } finally {
            try {
                is.close();
            } catch (Exception e) {
                logger.warn("Error closing input stream", e);
            }
        }
    }
}
