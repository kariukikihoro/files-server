package com.muriithi.filesserver.serve;

import jakarta.servlet.http.HttpServletRequest;
import org.apache.commons.text.StringEscapeUtils;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

public class OfficeHtmlGenerator {

    public static String createOfficeMultiViewerHtml(String filename, String type, HttpServletRequest request,
                                                     boolean isLargeFile, String publicFileUrl, String officeViewerUrl) {
        String safeFilename = StringEscapeUtils.escapeHtml4(filename);
        String fileExtension = filename.substring(filename.lastIndexOf(".") + 1).toUpperCase();

        String icon;
        switch (fileExtension.toLowerCase()) {
            case "doc":
            case "docx":
                icon = "📝";
                break;
            case "xls":
            case "xlsx":
                icon = "📊";
                break;
            case "ppt":
            case "pptx":
                icon = "📽️";
                break;
            case "rtf":
                icon = "📄";
                break;
            default:
                icon = "📄";
                break;
        }

        String largeFileWarning = "";
        if (isLargeFile) {
            largeFileWarning =
                    "<div class=\"warning-message\" style=\"background-color:#fff4f6;border-left:4px solid #a1003d;" +
                            "color:#a1003d;padding:12px;margin:16px 0;border-radius:4px;\">" +
                            "<strong>⚠️ Note:</strong> This file is large (>25MB) and may load slowly." +
                            "</div>";
        }

        // Viewer URL
        String officeEmbedUrl = String.format("%s?src=%s",
                officeViewerUrl,
                URLEncoder.encode(publicFileUrl, StandardCharsets.UTF_8));

        StringBuilder html = new StringBuilder();

        html.append("<!DOCTYPE html>")
                .append("<html lang=\"en\">")
                .append("<head>")
                .append("<meta charset=\"UTF-8\">")
                .append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">")
                .append("<title>").append(safeFilename).append(" - Document Viewer</title>")
                .append("<style>")
                .append("*{margin:0;padding:0;box-sizing:border-box;}body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background-color:#f8fafc;height:100vh;display:flex;flex-direction:column;}")
                .append(".header{background:linear-gradient(135deg,#a1003d 0%,#d32f2f 100%);color:white;padding:16px 20px;box-shadow:0 2px 10px rgba(0,0,0,0.1);display:flex;justify-content:space-between;align-items:center;}")
                .append(".filename{font-size:18px;font-weight:600;max-width:60%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}")
                .append(".btn{padding:8px 16px;border:none;border-radius:6px;cursor:pointer;text-decoration:none;display:inline-flex;align-items:center;gap:6px;transition:all 0.2s;}")
                .append(".btn-primary{background-color:rgba(255,255,255,0.2);color:white;}")
                .append(".btn-primary:hover{background-color:rgba(255,255,255,0.3);}")
                .append(".btn-success{background-color:#a1003d;color:white;font-size:16px;padding:12px 24px;}")
                .append(".btn-success:hover{background-color:#7a002d;}")
                .append(".viewer-tabs{background-color:#f3f4f6;border-bottom:1px solid #d1d5db;padding:0 20px;}")
                .append(".tabs{display:flex;gap:0;overflow-x:auto;}")
                .append(".tab{padding:12px 16px;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.2s;white-space:nowrap;font-weight:500;color:#6b7280;}")
                .append(".tab:hover{background-color:#f9fafb;color:#374151;}")
                .append(".tab.active{background-color:white;border-bottom-color:#a1003d;color:#a1003d;}")
                .append(".content{flex:1;position:relative;overflow:hidden;}")
                .append(".viewer-panel{display:none;height:100%;position:relative;}")
                .append(".viewer-panel.active{display:block;}")
                .append(".viewer-iframe{width:100%;height:100%;border:none;background-color:white;}")
                .append(".loading{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center;color:#6b7280;}")
                .append(".loading-spinner{width:40px;height:40px;border:4px solid #e5e7eb;border-top:4px solid #a1003d;border-radius:50%;animation:spin 1s linear infinite;margin:0 auto 16px;}")
                .append("@keyframes spin{0%{transform:rotate(0deg);}100%{transform:rotate(360deg);}}")
                .append(".error-message{background-color:#fef2f2;border:1px solid #fecaca;color:#dc2626;padding:20px;margin:20px;border-radius:6px;text-align:center;display:none;}")
                .append(".viewer-info{padding:40px 20px;text-align:center;max-width:600px;margin:0 auto;}")
                .append("</style>")
                .append("</head>")
                .append("<body>")
                .append("<div class=\"header\">")
                .append("<h1 class=\"filename\">").append(safeFilename).append(" ").append(icon).append("</h1>")
                .append("<div><a href=\"").append(publicFileUrl).append("\" class=\"btn btn-primary\" download>⬇️ Download</a></div>")
                .append("</div>")
                .append(largeFileWarning)
                .append("<div class=\"viewer-tabs\"><div class=\"tabs\">")
                .append("<div class=\"tab active\" onclick=\"switchTab('office-online')\">🌐 Office Online</div>")
                .append("<div class=\"tab\" onclick=\"switchTab('download')\">⬇️ Download</div>")
                .append("</div></div>")
                .append("<div class=\"content\">")
                .append("<div class=\"viewer-panel active\" id=\"office-online\">")
                .append("<div class=\"loading\" id=\"office-loading\"><div class=\"loading-spinner\"></div><p>Loading with Office Online...</p></div>")
                .append("<div class=\"error-message\" id=\"office-error\"><strong>⚠️ Viewer Error</strong><br>Unable to load the document in the viewer. The file may be password-protected or the viewer service is temporarily unavailable.<br><br><a href=\"").append(publicFileUrl).append("\" class=\"btn btn-primary\" download>⬇️ Download to view locally</a></div>")
                .append("<iframe class=\"viewer-iframe\" id=\"office-iframe\" style=\"display:none;\" src=\"").append(officeEmbedUrl).append("\" onload=\"handleViewerLoad()\" onerror=\"handleViewerError()\"></iframe>")
                .append("</div>")
                .append("<div class=\"viewer-panel\" id=\"download\">")
                .append("<div class=\"viewer-info\">")
                .append("<h3>⬇️ Download Document</h3><p>Download to view with your preferred application.</p>")
                .append("<div style=\"margin:24px 0;\"><a href=\"").append(publicFileUrl).append("\" class=\"btn btn-success\" download>⬇️ Download ").append(fileExtension).append(" Document</a></div>")
                .append("<div style=\"background-color:#f8fafc;padding:20px;border-radius:8px;text-align:left;max-width:400px;margin:0 auto;\">")
                .append("<h4 style=\"color:#374151;margin-bottom:12px;\">📋 File Information</h4>")
                .append("<p><strong>Type:</strong> ").append(fileExtension).append(" Document</p>")
                .append("<p><strong>Filename:</strong> ").append(safeFilename).append("</p>")
                .append("<p><strong>Note:</strong> Some Office files may require authentication to view online.</p>")
                .append("</div></div></div></div>")
                .append("<script>")
                .append("var currentTab='office-online';var viewerLoadTimeout;")
                .append("function switchTab(tabId){document.querySelector('.tab.active').classList.remove('active');document.querySelector('.viewer-panel.active').classList.remove('active');document.querySelector('.tab[onclick*=\"'+tabId+'\"]').classList.add('active');document.getElementById(tabId).classList.add('active');currentTab=tabId;if(tabId==='office-online'){loadOfficeViewer();}}")
                .append("function loadOfficeViewer(){document.getElementById('office-loading').style.display='block';document.getElementById('office-iframe').style.display='none';document.getElementById('office-error').style.display='none';viewerLoadTimeout=setTimeout(function(){if(document.getElementById('office-loading').style.display!=='none'){handleViewerError();}},15000);setTimeout(function(){document.getElementById('office-iframe').style.display='block';},1000);}")
                .append("function handleViewerLoad(){clearTimeout(viewerLoadTimeout);document.getElementById('office-loading').style.display='none';document.getElementById('office-iframe').style.display='block';console.log('Office viewer loaded successfully');}")
                .append("function handleViewerError(){clearTimeout(viewerLoadTimeout);document.getElementById('office-loading').style.display='none';document.getElementById('office-error').style.display='block';console.warn('Office viewer failed to load');setTimeout(function(){if(currentTab==='office-online'){switchTab('download');}},3000);}")
                .append("document.addEventListener('DOMContentLoaded',function(){loadOfficeViewer();});")
                .append("document.addEventListener('keydown',function(e){if((e.ctrlKey||e.metaKey)&&e.key==='s'){e.preventDefault();window.location.href='").append(publicFileUrl).append("';}});")
                .append("</script>")
                .append("</body></html>");

        return html.toString();
    }
}
