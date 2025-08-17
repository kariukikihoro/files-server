package com.muriithi.filesserver.serve;

import jakarta.servlet.http.HttpServletRequest;
import org.apache.commons.text.StringEscapeUtils;

public class PdfHtmlGenerator {

    public static String createPdfViewerHtml(String filename, String type, HttpServletRequest request, String fileUrl) {
        String safeFilename = StringEscapeUtils.escapeHtml4(filename);

        StringBuilder html = new StringBuilder();

        html.append("<!DOCTYPE html>")
                .append("<html lang=\"en\">")
                .append("<head>")
                .append("<meta charset=\"UTF-8\">")
                .append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">")
                .append("<title>").append(safeFilename).append(" - PDF Viewer</title>")
                .append("<style>")
                .append("*{margin:0;padding:0;box-sizing:border-box;}")
                .append("body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background-color:#f8fafc;color:#1f2937;height:100vh;display:flex;flex-direction:column;}")
                .append(".header{background:linear-gradient(135deg,#a1003d 0%,#d32f2f 100%);padding:14px 20px;color:white;box-shadow:0 2px 10px rgba(0,0,0,0.2);display:flex;justify-content:space-between;align-items:center;flex-shrink:0;}")
                .append(".filename{font-size:18px;font-weight:600;color:white;max-width:60%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin:0;}")
                .append(".controls{display:flex;gap:10px;align-items:center;}")
                .append(".btn{padding:8px 16px;border:none;border-radius:8px;font-size:14px;cursor:pointer;transition:all 0.2s;text-decoration:none;display:inline-flex;align-items:center;gap:6px;box-shadow:0 2px 6px rgba(0,0,0,0.15);}")
                .append(".btn-primary{background-color:white;color:#a1003d;font-weight:600;}")
                .append(".btn-primary:hover{background-color:#fce7ef;color:#7a002d;}")
                .append(".viewer-container{flex:1;position:relative;overflow:hidden;background-color:#ffffff;}")
                .append(".pdf-embed{width:100%;height:100%;border:none;background:white;}")
                .append(".error-message{display:none;background-color:#fef2f2;border:1px solid #fecaca;color:#dc2626;padding:20px;margin:20px;border-radius:6px;text-align:center;}")
                .append("</style>")
                .append("</head>")
                .append("<body>")
                .append("<div class=\"header\">")
                .append("<h1 class=\"filename\">📄 ").append(safeFilename).append("</h1>")
                .append("<div class=\"controls\"><a href=\"").append(fileUrl).append("\" class=\"btn btn-primary\" download>⬇️ Download</a></div>")
                .append("</div>")
                .append("<div class=\"viewer-container\">")
                .append("<div class=\"error-message\" id=\"pdf-error\">")
                .append("<strong>⚠️ PDF Viewer Error</strong><br>")
                .append("Your browser doesn't support PDF viewing or the file couldn't be loaded.<br>")
                .append("<a href=\"").append(fileUrl).append("\" download>Click here to download the file</a>.")
                .append("</div>")
                .append("<object class=\"pdf-embed\" data=\"").append(fileUrl).append("#toolbar=1&navpanes=1&scrollbar=1&view=FitH\" type=\"application/pdf\" onerror=\"handlePdfError()\">")
                .append("<embed src=\"").append(fileUrl).append("#toolbar=1&navpanes=1&scrollbar=1&view=FitH\" type=\"application/pdf\" onerror=\"handlePdfError()\" />")
                .append("</object>")
                .append("</div>")
                .append("<script>")
                .append("function handlePdfError(){document.getElementById('pdf-error').style.display='block';var embeds=document.getElementsByClassName('pdf-embed');for(var i=0;i<embeds.length;i++){embeds[i].style.display='none';}}")
                .append("document.addEventListener('keydown',function(e){if((e.ctrlKey||e.metaKey)&&e.key==='s'){e.preventDefault();window.location.href='").append(fileUrl).append("';}});")
                .append("</script>")
                .append("</body>")
                .append("</html>");

        return html.toString();
    }

}
