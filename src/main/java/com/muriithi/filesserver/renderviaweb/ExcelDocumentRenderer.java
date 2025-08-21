package com.muriithi.filesserver.renderviaweb;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

import static com.muriithi.filesserver.renderviaweb.OfficeDocumentRenderer.escapeHtml;
import static com.muriithi.filesserver.renderviaweb.OfficeDocumentRenderer.extractDocumentName;

/**
 * Office Document Renderer compatible with POI 3.10.1
 */
public class ExcelDocumentRenderer {

    private static final Logger logger = LoggerFactory.getLogger(ExcelDocumentRenderer.class);

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
                .append("margin-bottom: 25px; background: linear-gradient(135deg, #D0715B  0%, #8B0A23 100%); ")
                .append("-webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }")

                .append(".paragraph { margin: 12px 0; font-size: 1.1em; line-height: 1.8; }")

                /**
                 * Modern sheet tabs with improved UX
                 * */

                .append(".sheet-tabs-container { margin: 20px 0 15px 0; display: flex; flex-wrap: wrap; gap: 6px; }")
                .append(".sheet-tab { background: linear-gradient(135deg, #D0715B  0%, #8B0A23 100%); color: white; ")
                .append("padding: 10px 20px; border-radius: 20px; font-weight: 600; font-size: 0.9em; ")
                .append("box-shadow: 0 3px 12px rgba(169, 12, 43, 0.3); cursor: pointer; transition: all 0.3s ease; ")
                .append("border: none; user-select: none; }")
                .append(".sheet-tab:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(169, 12, 43, 0.4); }")
                .append(".sheet-tab.active { background: linear-gradient(135deg, #8B0A23 0%, #D0715B  100%); }")

                /**
                 *  Sheet content with show/hide functionality
                 * */

                .append(".sheet-content { display: none; margin-top: 15px; }")
                .append(".sheet-content.active { display: block !important; }")

                /**
                 *  Excel-like table styling with full content visibility
                 * */

                .append(".table-container { margin: 15px 0; background: white; border-radius: 8px; ")
                .append("overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.08); border: 1px solid #e1e5e9; }")

                .append(".table-header { background: linear-gradient(135deg, #D0715B  0%, #8B0A23 100%); ")
                .append("color: white; padding: 12px 15px; font-weight: 600; font-size: 1em; }")

                .append(".table-wrapper { overflow: auto; max-height: 70vh; background: white; }")

                /**
                 *  Excel-like table with proper structure
                 * */

                .append("table { border-collapse: separate; border-spacing: 0; width: auto; min-width: 100%; font-size: 0.85em; ")
                .append("font-family: 'Segoe UI', Arial, sans-serif; }")

                /**
                 * Column headers with Excel-like appearance
                 * */

                .append("th { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); ")
                .append("font-weight: 600; padding: 8px 4px; text-align: center; border: 1px solid #d0d7de; ")
                .append("position: sticky; top: 0; z-index: 10; white-space: nowrap; font-size: 0.8em; ")
                .append("color: #495057; min-width: 60px; max-width: 200px; }")

                /**
                 *  Row headers (A, B, C... style)
                 * */

                .append("th.row-header { background: linear-gradient(135deg, #f1f3f4 0%, #e8eaed 100%); ")
                .append("font-weight: 600; text-align: center; color: #5f6368; min-width: 40px; max-width: 40px; ")
                .append("position: sticky; left: 0; z-index: 11; }")

                /**
                 * Cell styling with dynamic width
                 * */

                .append("td { padding: 6px 8px; border: 1px solid #d0d7de; vertical-align: middle; ")
                .append("white-space: pre-wrap; word-wrap: break-word; min-width: 80px; max-width: 300px; ")
                .append("background: white; position: relative; }")

                /**
                 * Row header cells
                 * */

                .append("td.row-header { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); ")
                .append("font-weight: 500; text-align: center; color: #5f6368; min-width: 40px; max-width: 40px; ")
                .append("position: sticky; left: 0; z-index: 10; border-right: 2px solid #d0d7de; }")

                .append("tr:nth-child(even) td:not(.row-header) { background-color: #fafbfc; }")
                .append("tr:hover td:not(.row-header) { background-color: #e8f0fe; }")

                /**
                 *  Cell type specific styling
                 *  **/

                .append("td.number { text-align: right; font-family: 'Consolas', 'Monaco', monospace; }")
                .append("td.date { color: #1a73e8; }")
                .append("td.boolean { text-align: center; font-weight: bold; color: #137333; }")
                .append("td.formula { background-color: #fef7e0; font-family: 'Consolas', 'Monaco', monospace; color: #b7651c; }")
                .append("td.error { background-color: #fce8e6; color: #d73027; font-weight: bold; text-align: center; }")

                /**
                 * Auto-sizing cells
                 * **/

                .append("td.auto-width { width: auto; white-space: nowrap; }")
                .append("td.text-cell { white-space: pre-wrap; word-break: break-word; }")

                /**
                 *  Empty state styling
                 *  **/

                .append(".empty-sheet { text-align: center; padding: 50px 20px; color: #6c757d; }")

                /**
                 * Summary stats
                 * **/

                .append(".sheet-summary { background: #f8f9fa; padding: 12px 15px; border-radius: 6px; ")
                .append("margin-bottom: 15px; font-size: 0.85em; color: #6c757d; display: flex; justify-content: space-between; ")
                .append("flex-wrap: wrap; gap: 15px; border-left: 4px solid #D0715B ; }")
                .append(".summary-item { display: flex; align-items: center; gap: 5px; }")
                .append(".summary-item strong { color: #D0715B ; }")

                /**
                 * Column resize handles
                 * **/

                .append("th { position: relative; }")
                .append("th::after { content: ''; position: absolute; top: 0; right: 0; width: 5px; height: 100%; ")
                .append("cursor: col-resize; background: transparent; }")
                .append("th:hover::after { background: rgba(169, 12, 43, 0.3); }")

                /**
                 * Responsive design
                 * **/

                .append("@media (max-width: 768px) {")
                .append("  .document-container { margin: 5px; padding: 10px; }")
                .append("  .sheet-tabs-container { justify-content: flex-start; }")
                .append("  .sheet-tab { padding: 8px 16px; font-size: 0.8em; }")
                .append("  table { font-size: 0.75em; }")
                .append("  th, td { padding: 4px 6px; }")
                .append("  .document-title { font-size: 1.8em; }")
                .append("  .table-wrapper { max-height: 50vh; }")
                .append("  .document-name-header { padding: 10px 15px; }")
                .append("  .document-name { font-size: 1em; }")
                .append("}")

                .append("</style>");



        BASE_STYLE = sb.toString();
    }

    static {
        SUPPORTED_EXTENSIONS.addAll(WORD_EXTENSIONS);
        SUPPORTED_EXTENSIONS.addAll(EXCEL_EXTENSIONS);
    }

    private final DecimalFormat numberFormat = new DecimalFormat("#,##0.##");
    private final SimpleDateFormat dateFormat = new SimpleDateFormat("MMM dd, yyyy HH:mm");

    public byte[] renderExcelDocument(byte[] content, String fileName) throws Exception {
        return convertExcelToHtml(content, fileName);
    }

    private byte[] convertExcelToHtml(byte[] excelContent, String fileName) throws Exception {
        InputStream is = new ByteArrayInputStream(excelContent);
        try {
            Workbook workbook = WorkbookFactory.create(is);
            try {
                StringBuilder html = new StringBuilder();
                html.append("<!DOCTYPE html><html><head><meta charset='UTF-8'>")
                        .append("<title>Excel Document</title>")
                        .append(BASE_STYLE)
                        .append("</head><body>");

                html.append("<div class='document-container'>");
                html.append("<h1 class='document-title'>Excel Workbook</h1>");

                String documentName = extractDocumentName(fileName);
                html.append("<div class='document-name-header'>");
                html.append("<div class='document-name'>").append(escapeHtml(documentName)).append("</div>");
                html.append("</div>");

                if (workbook.getNumberOfSheets() > 1) {

                    html.append("<div class='sheet-tabs-container'>");

                    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

                        Sheet sheet = workbook.getSheetAt(i);

                        html.append("<button class='sheet-tab")
                                .append(i == 0 ? " active" : "")
                                .append("' data-sheet-index='").append(i).append("'>")
                                .append(escapeHtml(sheet.getSheetName()))
                                .append("</button>");

                    }
                    html.append("</div>");
                }

                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    html.append("<div class='sheet-content").append(i == 0 ? " active" : "").append("'>");
                    processSheet(html, sheet, i);
                    html.append("</div>");
                }

                html.append("</div>");

                html.append("<script>")
                        .append("function showSheet(index) {")
                        .append("  console.log('Switching to sheet:', index);")
                        .append("  var contents = document.querySelectorAll('.sheet-content');")
                        .append("  var tabs = document.querySelectorAll('.sheet-tab');")
                        .append("  ")
                        .append("  for (var i = 0; i < contents.length; i++) {")
                        .append("    if (i === index) {")
                        .append("      contents[i].style.display = 'block';")
                        .append("      contents[i].classList.add('active');")
                        .append("    } else {")
                        .append("      contents[i].style.display = 'none';")
                        .append("      contents[i].classList.remove('active');")
                        .append("    }")
                        .append("  }")
                        .append("  ")
                        .append("  for (var i = 0; i < tabs.length; i++) {")
                        .append("    if (i === index) {")
                        .append("      tabs[i].classList.add('active');")
                        .append("    } else {")
                        .append("      tabs[i].classList.remove('active');")
                        .append("    }")
                        .append("  }")
                        .append("}")
                        .append("")
                        .append("var tabs = document.querySelectorAll('.sheet-tab');")
                        .append("for (var i = 0; i < tabs.length; i++) {")
                        .append("  (function(index) {")
                        .append("    tabs[index].onclick = function() {")
                        .append("      var sheetIndex = parseInt(this.getAttribute('data-sheet-index'));")
                        .append("      if (!isNaN(sheetIndex)) {")
                        .append("        showSheet(sheetIndex);")
                        .append("      }")
                        .append("    };")
                        .append("  })(i);")
                        .append("}")
                        .append("")
                        .append("if (tabs.length > 0) {")
                        .append("  showSheet(0);")
                        .append("}")
                        .append("</script>")
                        .append("</body></html>");
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

    private void processSheet(StringBuilder html, Sheet sheet, int sheetIndex) {
        if (sheet.getPhysicalNumberOfRows() == 0) {
            html.append("<div class='empty-sheet'>");
            html.append("<div style='font-size: 3em; margin-bottom: 20px; opacity: 0.3;'>📄</div>");
            html.append("<h3>This sheet is empty</h3>");
            html.append("<p>No data found in this worksheet</p>");
            html.append("</div>");
            return;
        }

        int totalRows = sheet.getPhysicalNumberOfRows();
        int maxCols = getMaxColumnCount(sheet);
        int nonEmptyRows = 0;
        for (Row row : sheet) {
            if (!isEmptyRow(row)) nonEmptyRows++;
        }

        html.append("<div class='sheet-summary'>");
        html.append("<div class='summary-item'><strong>Rows:</strong> ").append(nonEmptyRows).append("</div>");
        html.append("<div class='summary-item'><strong>Columns:</strong> ").append(maxCols).append("</div>");
        html.append("<div class='summary-item'><strong>Sheet:</strong> ").append(escapeHtml(sheet.getSheetName())).append("</div>");
        html.append("</div>");

        html.append("<div class='table-container'>");
        html.append("<div class='table-header'>").append(escapeHtml(sheet.getSheetName())).append("</div>");
        html.append("<div class='table-wrapper'>");
        html.append("<table>");

        html.append("<thead><tr>");
        html.append("<th class='row-header'></th>");

        for (int col = 0; col < maxCols; col++) {
            html.append("<th>").append(getExcelColumnName(col)).append("</th>");
        }

        html.append("</tr></thead>");

        html.append("<tbody>");

        int rowNum = 1;

        final int MAX_ROWS = 2000;

        int processedRows = 0;

        for (Row row : sheet) {

            if (isEmptyRow(row)) {
                rowNum++;
                continue;
            }

            if (processedRows >= MAX_ROWS) {

                html.append("<tr><td class='row-header'>...</td>");
                html.append("<td colspan='").append(maxCols)
                        .append("' style='text-align: center; padding: 15px; background: #fff3cd; color: #856404; font-weight: bold;'>")
                        .append("⚠️ Data truncated - Showing first ").append(MAX_ROWS)
                        .append(" rows of ").append(totalRows).append(" total rows</td></tr>");
                break;
            }

            html.append("<tr>");

            html.append("<td class='row-header'>").append(rowNum).append("</td>");

            for (int cellIndex = 0; cellIndex < maxCols; cellIndex++) {
                Cell cell = row.getCell(cellIndex);

                html.append("<td");

                if (cell != null) {
                    String cellClass = getCellCssClass(cell);
                    if (!cellClass.isEmpty()) {
                        html.append(" class='").append(cellClass).append("'");
                    }
                }

                html.append(">");

                if (cell != null) {
                    String cellValue = formatCellValue(cell);
                    html.append(escapeHtml(cellValue));
                } else {
                    html.append("&nbsp;");
                }

                html.append("</td>");
            }

            html.append("</tr>");
            rowNum++;
            processedRows++;
        }

        html.append("</tbody>");
        html.append("</table>");
        html.append("</div>");
        html.append("</div>");
    }

    private String getExcelColumnName(int columnIndex) {

        StringBuilder result = new StringBuilder();

        while (columnIndex >= 0) {

            result.insert(0, (char) ('A' + columnIndex % 26));
            columnIndex = columnIndex / 26 - 1;
        }
        return result.toString();
    }

    private String getCellCssClass(Cell cell) {
        switch (cell.getCellType()) {
            case NUMERIC:
                return DateUtil.isCellDateFormatted(cell) ? "date" : "number";
            case BOOLEAN:
                return "boolean";
            case FORMULA:
                return "formula";
            case ERROR:
                return "error";
            default:
                return "";
        }
    }

    private int getMaxColumnCount(Sheet sheet) {
        int maxCols = 0;
        for (Row row : sheet) {
            maxCols = Math.max(maxCols, row.getLastCellNum());
        }
        return Math.min(maxCols, 100); // Increased column limit for better Excel compatibility
    }

    private boolean isEmptyRow(Row row) {
        if (row == null) return true;

        for (Cell cell : row) {
            if (cell != null
                    && cell.getCellType() != CellType.BLANK
                    && !cell.toString().trim().isEmpty()) {
                return false;
            }
        }
        return true;
    }

    private String formatCellValue(Cell cell) {
        try {
            switch (cell.getCellType()) {
                case STRING:
                    String stringValue = cell.getStringCellValue();
                    return stringValue.replace("\n", "\n").replace("\r", "");

                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return dateFormat.format(cell.getDateCellValue());
                    } else {
                        double numValue = cell.getNumericCellValue();
                        if (numValue == Math.floor(numValue)
                                && numValue <= Long.MAX_VALUE
                                && numValue >= Long.MIN_VALUE) {
                            return String.valueOf((long) numValue);
                        } else {
                            return numberFormat.format(numValue);
                        }
                    }

                case BOOLEAN:
                    return cell.getBooleanCellValue() ? "TRUE" : "FALSE";

                case FORMULA:
                    try {
                        switch (cell.getCachedFormulaResultType()) {
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    return dateFormat.format(cell.getDateCellValue());
                                } else {
                                    double numValue = cell.getNumericCellValue();
                                    if (numValue == Math.floor(numValue)
                                            && numValue <= Long.MAX_VALUE
                                            && numValue >= Long.MIN_VALUE) {
                                        return String.valueOf((long) numValue);
                                    } else {
                                        return numberFormat.format(numValue);
                                    }
                                }
                            case STRING:
                                return cell.getStringCellValue();
                            case BOOLEAN:
                                return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                            default:
                                return "=" + cell.getCellFormula();
                        }
                    } catch (Exception e) {
                        return "=" + cell.getCellFormula();
                    }

                case ERROR:
                    return "#ERROR!";

                case BLANK:
                default:
                    return "";
            }
        } catch (Exception e) {
            logger.warn("Error formatting cell value: {}", e.getMessage());
            return "#ERROR!";
        }
    }
}