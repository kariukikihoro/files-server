package com.muriithi.filesserver.renderviaweb;

import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.text.DecimalFormat;
import java.util.*;
import java.util.regex.Pattern;

import static com.muriithi.filesserver.renderviaweb.OfficeDocumentRenderer.escapeHtml;
import static com.muriithi.filesserver.renderviaweb.OfficeDocumentRenderer.extractDocumentName;

/**
 * CSV Document Renderer for web display
 */
@Slf4j
@Component
public class CsvDocumentRenderer {

    private static final Set<String> CSV_EXTENSIONS =
            new HashSet<>(Arrays.asList(".csv", ".txt"));

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
                 */
                .append(".document-container { max-width: 98%; margin: 10px auto; padding: 20px; ")
                .append("background: rgba(255, 255, 255, 0.95); backdrop-filter: blur(10px); border-radius: 12px; ")
                .append("box-shadow: 0 10px 30px rgba(0,0,0,0.08), 0 0 0 1px rgba(255,255,255,0.2); }")

                /**
                 * Document name header styling
                 */
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
                 */
                .append(".document-title { font-size: 2.2em; font-weight: 700; color: #2c3e50; text-align: center; ")
                .append("margin-bottom: 25px; background: linear-gradient(135deg, #D0715B  0%, #8B0A23 100%); ")
                .append("-webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }")

                /**
                 * CSV-specific controls
                 */
                .append(".csv-controls { margin: 20px 0 15px 0; display: flex; flex-wrap: wrap; gap: 15px; align-items: center; }")
                .append(".control-group { display: flex; align-items: center; gap: 8px; }")
                .append(".control-label { font-weight: 600; color: #495057; font-size: 0.9em; }")
                .append(".control-input { padding: 8px 12px; border: 2px solid #e1e5e9; border-radius: 6px; ")
                .append("font-size: 0.9em; transition: border-color 0.3s ease; }")
                .append(".control-input:focus { outline: none; border-color: #D0715B; }")
                .append(".control-button { background: linear-gradient(135deg, #D0715B  0%, #8B0A23 100%); color: white; ")
                .append("padding: 8px 16px; border-radius: 6px; font-weight: 600; font-size: 0.9em; ")
                .append("border: none; cursor: pointer; transition: all 0.3s ease; }")
                .append(".control-button:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(169, 12, 43, 0.3); }")

                /**
                 * Table styling similar to Excel renderer
                 */
                .append(".table-container { margin: 15px 0; background: white; border-radius: 8px; ")
                .append("overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.08); border: 1px solid #e1e5e9; }")

                .append(".table-header { background: linear-gradient(135deg, #D0715B  0%, #8B0A23 100%); ")
                .append("color: white; padding: 12px 15px; font-weight: 600; font-size: 1em; }")

                .append(".table-wrapper { overflow: auto; max-height: 70vh; background: white; }")

                /**
                 * CSV table with proper structure
                 */
                .append("table { border-collapse: separate; border-spacing: 0; width: auto; min-width: 100%; font-size: 0.85em; ")
                .append("font-family: 'Segoe UI', Arial, sans-serif; }")

                /**
                 * Column headers
                 */
                .append("th { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); ")
                .append("font-weight: 600; padding: 8px 12px; text-align: left; border: 1px solid #d0d7de; ")
                .append("position: sticky; top: 0; z-index: 10; white-space: nowrap; font-size: 0.85em; ")
                .append("color: #495057; min-width: 100px; max-width: 300px; }")

                /**
                 * Row headers (1, 2, 3... style)
                 */
                .append("th.row-header { background: linear-gradient(135deg, #f1f3f4 0%, #e8eaed 100%); ")
                .append("font-weight: 600; text-align: center; color: #5f6368; min-width: 50px; max-width: 50px; ")
                .append("position: sticky; left: 0; z-index: 11; }")

                /**
                 * Cell styling
                 */
                .append("td { padding: 8px 12px; border: 1px solid #d0d7de; vertical-align: top; ")
                .append("white-space: pre-wrap; word-wrap: break-word; min-width: 100px; max-width: 300px; ")
                .append("background: white; position: relative; }")

                /**
                 * Row header cells
                 */
                .append("td.row-header { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); ")
                .append("font-weight: 500; text-align: center; color: #5f6368; min-width: 50px; max-width: 50px; ")
                .append("position: sticky; left: 0; z-index: 10; border-right: 2px solid #d0d7de; }")

                .append("tr:nth-child(even) td:not(.row-header) { background-color: #fafbfc; }")
                .append("tr:hover td:not(.row-header) { background-color: #e8f0fe; }")

                /**
                 * Cell type specific styling
                 */
                .append("td.number { text-align: right; font-family: 'Consolas', 'Monaco', monospace; color: #1565c0; }")
                .append("td.date { color: #1a73e8; }")
                .append("td.email { color: #7b1fa2; }")
                .append("td.url { color: #1976d2; text-decoration: underline; }")
                .append("td.large-text { max-width: 400px; }")

                /**
                 * CSV summary stats
                 */
                .append(".csv-summary { background: #f8f9fa; padding: 12px 15px; border-radius: 6px; ")
                .append("margin-bottom: 15px; font-size: 0.85em; color: #6c757d; display: flex; justify-content: space-between; ")
                .append("flex-wrap: wrap; gap: 15px; border-left: 4px solid #D0715B ; }")
                .append(".summary-item { display: flex; align-items: center; gap: 5px; }")
                .append(".summary-item strong { color: #D0715B ; }")

                /**
                 * Empty state styling
                 */
                .append(".empty-csv { text-align: center; padding: 50px 20px; color: #6c757d; }")

                /**
                 * Pagination controls
                 */
                .append(".pagination-container { margin: 15px 0; display: flex; justify-content: center; align-items: center; gap: 10px; }")
                .append(".pagination-button { background: #f8f9fa; border: 1px solid #d0d7de; padding: 8px 12px; ")
                .append("border-radius: 6px; cursor: pointer; transition: all 0.3s ease; }")
                .append(".pagination-button:hover { background: #e9ecef; }")
                .append(".pagination-button.active { background: linear-gradient(135deg, #D0715B  0%, #8B0A23 100%); color: white; border-color: #8B0A23; }")
                .append(".pagination-info { color: #6c757d; font-size: 0.9em; }")

                /**
                 * Responsive design
                 */
                .append("@media (max-width: 768px) {")
                .append("  .document-container { margin: 5px; padding: 10px; }")
                .append("  .csv-controls { flex-direction: column; align-items: stretch; }")
                .append("  .control-group { justify-content: space-between; }")
                .append("  table { font-size: 0.75em; }")
                .append("  th, td { padding: 4px 8px; min-width: 80px; max-width: 200px; }")
                .append("  .document-title { font-size: 1.8em; }")
                .append("  .table-wrapper { max-height: 50vh; }")
                .append("}")

                .append("</style>");

        BASE_STYLE = sb.toString();
    }

    private final DecimalFormat numberFormat = new DecimalFormat("#,##0.##");

    private static final Pattern EMAIL_PATTERN = Pattern.compile(
            "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$"
    );

    private static final Pattern URL_PATTERN = Pattern.compile(
            "^(https?://)?(www\\.)?[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}(/.*)?$"
    );

    private static final Pattern NUMBER_PATTERN = Pattern.compile(
            "^-?\\d+(\\.\\d+)?$"
    );

    public boolean supports(String fileName, String format) {
        if (fileName == null || format == null) {
            return false;
        }

        String extension = getFileExtension(fileName);
        return CSV_EXTENSIONS.contains(extension.toLowerCase()) &&
                SUPPORTED_FORMATS.contains(format.toLowerCase());
    }

    private String getFileExtension(String fileName) {
        int lastDotIndex = fileName.lastIndexOf('.');
        return lastDotIndex > 0 ? fileName.substring(lastDotIndex) : "";
    }

    public byte[] renderCsvDocument(byte[] content, String fileName) throws Exception {
        return convertCsvToHtml(content, fileName);
    }

    private byte[] convertCsvToHtml(byte[] csvContent, String fileName) throws Exception {
        log.warn(":::::::::::  Converting CSV to HTML...");
        String csvText = new String(csvContent, "UTF-8");
        List<String[]> rows = parseCsv(csvText);

        StringBuilder html = new StringBuilder();
        html.append("<!DOCTYPE html><html><head><meta charset='UTF-8'>")
                .append("<title>CSV Document</title>")
                .append(BASE_STYLE)
                .append("</head><body>");

        html.append("<div class='document-container'>");
        html.append("<h1 class='document-title'>CSV Document</h1>");

        String documentName = extractDocumentName(fileName);
        html.append("<div class='document-name-header'>");
        html.append("<div class='document-name'>").append(escapeHtml(documentName)).append("</div>");
        html.append("</div>");

        processCsvData(html, rows, fileName);

        html.append("</div>");

        html.append("<script>")
                .append("var allRows = [];")
                .append("var filteredRows = [];")
                .append("var currentPage = 0;")
                .append("var rowsPerPage = 100;")
                .append("var totalRows = 0;")
                .append("")
                .append("function initializeCsv() {")
                .append("  var tableBody = document.querySelector('tbody');")
                .append("  if (tableBody) {")
                .append("    var rows = tableBody.querySelectorAll('tr');")
                .append("    for (var i = 0; i < rows.length; i++) {")
                .append("      allRows.push(rows[i]);")
                .append("    }")
                .append("    filteredRows = allRows.slice();")
                .append("    totalRows = allRows.length;")
                .append("    updatePagination();")
                .append("  }")
                .append("}")
                .append("")
                .append("function filterTable() {")
                .append("  var searchInput = document.getElementById('searchInput');")
                .append("  if (!searchInput) return;")
                .append("  ")
                .append("  var searchTerm = searchInput.value.toLowerCase();")
                .append("  filteredRows = [];")
                .append("  ")
                .append("  for (var i = 0; i < allRows.length; i++) {")
                .append("    var row = allRows[i];")
                .append("    var cells = row.querySelectorAll('td');")
                .append("    var found = false;")
                .append("    ")
                .append("    for (var j = 0; j < cells.length; j++) {")
                .append("      if (cells[j].textContent.toLowerCase().indexOf(searchTerm) !== -1) {")
                .append("        found = true;")
                .append("        break;")
                .append("      }")
                .append("    }")
                .append("    ")
                .append("    if (found) {")
                .append("      filteredRows.push(row);")
                .append("    }")
                .append("  }")
                .append("  ")
                .append("  currentPage = 0;")
                .append("  updatePagination();")
                .append("}")
                .append("")
                .append("function changeRowsPerPage() {")
                .append("  var select = document.getElementById('rowsPerPageSelect');")
                .append("  if (select) {")
                .append("    rowsPerPage = parseInt(select.value);")
                .append("    currentPage = 0;")
                .append("    updatePagination();")
                .append("  }")
                .append("}")
                .append("")
                .append("function changePage(page) {")
                .append("  currentPage = page;")
                .append("  updatePagination();")
                .append("}")
                .append("")
                .append("function updatePagination() {")
                .append("  var tableBody = document.querySelector('tbody');")
                .append("  if (!tableBody) return;")
                .append("  ")
                .append("  for (var i = 0; i < allRows.length; i++) {")
                .append("    allRows[i].style.display = 'none';")
                .append("  }")
                .append("  ")
                .append("  var startIndex = currentPage * rowsPerPage;")
                .append("  var endIndex = Math.min(startIndex + rowsPerPage, filteredRows.length);")
                .append("  ")
                .append("  for (var i = startIndex; i < endIndex; i++) {")
                .append("    filteredRows[i].style.display = '';")
                .append("  }")
                .append("  ")
                .append("  var paginationInfo = document.querySelector('.pagination-info');")
                .append("  if (paginationInfo) {")
                .append("    var totalPages = Math.ceil(filteredRows.length / rowsPerPage);")
                .append("    paginationInfo.textContent = 'Page ' + (currentPage + 1) + ' of ' + totalPages + ")
                .append("      ' (' + filteredRows.length + ' rows)';")
                .append("  }")
                .append("  ")
                .append("  updatePaginationButtons();")
                .append("}")
                .append("")
                .append("function updatePaginationButtons() {")
                .append("  var container = document.querySelector('.pagination-container');")
                .append("  if (!container) return;")
                .append("  ")
                .append("  var buttons = container.querySelectorAll('.pagination-button');")
                .append("  for (var i = 0; i < buttons.length; i++) {")
                .append("    buttons[i].remove();")
                .append("  }")
                .append("  ")
                .append("  var totalPages = Math.ceil(filteredRows.length / rowsPerPage);")
                .append("  if (totalPages <= 1) return;")
                .append("  ")
                .append("  var paginationInfo = container.querySelector('.pagination-info');")
                .append("  ")
                .append("  if (currentPage > 0) {")
                .append("    var prevBtn = document.createElement('button');")
                .append("    prevBtn.className = 'pagination-button';")
                .append("    prevBtn.textContent = '← Previous';")
                .append("    prevBtn.onclick = function() { changePage(currentPage - 1); };")
                .append("    container.insertBefore(prevBtn, paginationInfo);")
                .append("  }")
                .append("  ")
                .append("  if (currentPage < totalPages - 1) {")
                .append("    var nextBtn = document.createElement('button');")
                .append("    nextBtn.className = 'pagination-button';")
                .append("    nextBtn.textContent = 'Next →';")
                .append("    nextBtn.onclick = function() { changePage(currentPage + 1); };")
                .append("    container.appendChild(nextBtn);")
                .append("  }")
                .append("}")
                .append("")
                .append("initializeCsv();")
                .append("</script>")
                .append("</body></html>");

        return html.toString().getBytes("UTF-8");
    }

    private void processCsvData(StringBuilder html, List<String[]> rows, String fileName) {
        if (rows.isEmpty()) {
            html.append("<div class='empty-csv'>");
            html.append("<div style='font-size: 3em; margin-bottom: 20px; opacity: 0.3;'>📊</div>");
            html.append("<h3>This CSV is empty</h3>");
            html.append("<p>No data found in this file</p>");
            html.append("</div>");
            return;
        }
        int totalRows = rows.size();
        int maxCols = 0;
        int nonEmptyRows = 0;

        for (String[] row : rows) {
            maxCols = Math.max(maxCols, row.length);
            if (!isEmptyRow(row)) nonEmptyRows++;
        }

        html.append("<div class='csv-controls'>");
        html.append("<div class='control-group'>");
        html.append("<label class='control-label' for='searchInput'>Search:</label>");
        html.append("<input type='text' id='searchInput' class='control-input' placeholder='Filter rows...' oninput='filterTable()'>");
        html.append("</div>");
        html.append("<div class='control-group'>");
        html.append("<label class='control-label' for='rowsPerPageSelect'>Rows per page:</label>");
        html.append("<select id='rowsPerPageSelect' class='control-input' onchange='changeRowsPerPage()'>");
        html.append("<option value='50'>50</option>");
        html.append("<option value='100' selected>100</option>");
        html.append("<option value='250'>250</option>");
        html.append("<option value='500'>500</option>");
        html.append("</select>");
        html.append("</div>");
        html.append("</div>");

        html.append("<div class='csv-summary'>");
        html.append("<div class='summary-item'><strong>Total Rows:</strong> ").append(nonEmptyRows).append("</div>");
        html.append("<div class='summary-item'><strong>Columns:</strong> ").append(maxCols).append("</div>");
        html.append("<div class='summary-item'><strong>File:</strong> ").append(escapeHtml(fileName)).append("</div>");
        html.append("</div>");

        html.append("<div class='pagination-container'>");
        html.append("<div class='pagination-info'></div>");
        html.append("</div>");

        html.append("<div class='table-container'>");
        html.append("<div class='table-header'>CSV Data</div>");
        html.append("<div class='table-wrapper'>");
        html.append("<table>");

        if (!rows.isEmpty()) {
            html.append("<thead><tr>");
            html.append("<th class='row-header'>#</th>");

            String[] firstRow = rows.get(0);
            for (int col = 0; col < maxCols; col++) {
                String headerText;
                if (col < firstRow.length && firstRow[col] != null && !firstRow[col].trim().isEmpty()) {
                    headerText = firstRow[col].trim();
                } else {
                    headerText = "Column " + (col + 1);
                }
                html.append("<th>").append(escapeHtml(headerText)).append("</th>");
            }

            html.append("</tr></thead>");
        }

        html.append("<tbody>");

        final int MAX_INITIAL_ROWS = 2000;
        int processedRows = 0;
        int rowNum = 1;

        int startRow = hasHeaders(rows) ? 1 : 0;

        for (int rowIndex = startRow; rowIndex < rows.size(); rowIndex++) {
            String[] row = rows.get(rowIndex);

            if (isEmptyRow(row)) {

                log.warn("===========   No data found in this file");
                continue;
            }

            if (processedRows >= MAX_INITIAL_ROWS) {
                html.append("<tr style='display: none;'><td class='row-header'>...</td>");
                html.append("<td colspan='").append(maxCols)
                        .append("' style='text-align: center; padding: 15px; background: #fff3cd; color: #856404; font-weight: bold;'>")
                        .append("⚠️ Additional data available - Use pagination controls to navigate</td></tr>");
                break;
            }

            html.append("<tr>");
            html.append("<td class='row-header'>").append(rowNum).append("</td>");

            for (int cellIndex = 0; cellIndex < maxCols; cellIndex++) {
                String cellValue = "";
                if (cellIndex < row.length && row[cellIndex] != null) {
                    cellValue = row[cellIndex];
                }

                html.append("<td");

                String cellClass = getCellCssClass(cellValue);
                if (!cellClass.isEmpty()) {
                    html.append(" class='").append(cellClass).append("'");
                }

                html.append(">");

                if (cellValue.isEmpty()) {
                    html.append("&nbsp;");
                } else {
                    html.append(escapeHtml(cellValue));
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

        html.append("<div class='pagination-container'>");
        html.append("<div class='pagination-info'></div>");
        html.append("</div>");
    }

    private String getCellCssClass(String cellValue) {
        if (cellValue == null || cellValue.trim().isEmpty()) {
            return "";
        }

        String trimmed = cellValue.trim();

        if (NUMBER_PATTERN.matcher(trimmed).matches()) {
            return "number";
        }

        if (EMAIL_PATTERN.matcher(trimmed).matches()) {
            return "email";
        }

        if (URL_PATTERN.matcher(trimmed).matches()) {
            return "url";
        }

        if (isDateLike(trimmed)) {
            return "date";
        }

        if (trimmed.length() > 100) {
            return "large-text";
        }

        return "";
    }

    private boolean isDateLike(String value) {

        return value.matches("\\d{1,2}/\\d{1,2}/\\d{4}") ||
                value.matches("\\d{4}-\\d{2}-\\d{2}") ||
                value.matches("\\d{1,2}-\\d{1,2}-\\d{4}");
    }

    private boolean hasHeaders(List<String[]> rows) {
        if (rows.isEmpty()) return false;

        String[] firstRow = rows.get(0);

        int textCount = 0;
        int numberCount = 0;

        for (String cell : firstRow) {
            if (cell != null && !cell.trim().isEmpty()) {
                if (NUMBER_PATTERN.matcher(cell.trim()).matches()) {
                    numberCount++;
                } else {
                    textCount++;
                }
            }
        }

        return textCount > numberCount && rows.size() > 1;
    }

    private boolean isEmptyRow(String[] row) {
        if (row == null) return true;

        for (String cell : row) {
            if (cell != null && !cell.trim().isEmpty()) {
                return false;
            }
        }
        return true;
    }

    private List<String[]> parseCsv(String csvText) {
        List<String[]> rows = new ArrayList<>();

        if (csvText == null || csvText.trim().isEmpty()) {
            return rows;
        }

        String[] lines = csvText.split("\\r?\\n");

        for (String line : lines) {
            if (line.trim().isEmpty()) {
                continue;
            }

            String[] cells = parseCsvLine(line);
            rows.add(cells);
        }

        return rows;
    }

    private String[] parseCsvLine(String line) {
        List<String> cells = new ArrayList<>();
        StringBuilder currentCell = new StringBuilder();
        boolean inQuotes = false;
        boolean previousCharWasQuote = false;

        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);

            if (c == '"') {
                if (inQuotes) {
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') {

                        currentCell.append('"');
                        i++;
                        previousCharWasQuote = false;
                    } else {

                        inQuotes = false;
                        previousCharWasQuote = true;
                    }
                } else {

                    inQuotes = true;
                    previousCharWasQuote = false;
                }
            } else if (c == ',' && !inQuotes) {

                cells.add(currentCell.toString());
                currentCell = new StringBuilder();
                previousCharWasQuote = false;
            } else {
                currentCell.append(c);
                previousCharWasQuote = false;
            }
        }

        cells.add(currentCell.toString());

        return cells.toArray(new String[0]);
    }
}
