package com.muriithi.filesserver.serve;

import com.auxilii.msgparser.Message;
import com.auxilii.msgparser.MsgParser;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;

import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
public class FileServiceImpl implements FileService {

    private static final Logger log = LoggerFactory.getLogger(FileServiceImpl.class);
    private final UtilityMethodsService utilityMethodsService;

    @Value("${file.storage.base-path:./files}")
    private String basePath;


    @Value("${viewer.office.url:https://view.officeapps.live.com/op/embed.aspx}")
    private String officeViewerUrl;

    private final Map<String, String> typeToFolder = Map.of(
            "documents", "documents",
            "images", "images",
            "videos", "videos",
            "office", "office",
            "pdfs", "pdfs",
            "text", "text"
    );

    @Override
    public Map<String, Object> getAvailableTypes() {
        Map<String, Object> response = new HashMap<>();
        Map<String, List<String>> typeFiles = new HashMap<>();

        for (String type : typeToFolder.keySet()) {
            List<String> files = getFilesByType(type);
            typeFiles.put(type, files);
        }

        response.put("availableTypes", typeToFolder.keySet());
        response.put("filesByType", typeFiles);
        response.put("basePath", basePath);

        return response;
    }

    @Override
    public List<String> getFilesByType(String type) {
        if (!typeToFolder.containsKey(type)) {
            log.warn("Invalid file type requested: {}", type);
            return new ArrayList<>();
        }

        try {
            Path folderPath = Paths.get(basePath, typeToFolder.get(type));

            if (!Files.exists(folderPath)) {
                Files.createDirectories(folderPath);
                log.info("Created directory: {}", folderPath);
                return new ArrayList<>();
            }

            return Files.list(folderPath)
                    .filter(Files::isRegularFile)
                    .map(path -> path.getFileName().toString())
                    .sorted()
                    .collect(Collectors.toList());

        } catch (IOException e) {
            log.error("Error reading folder for type: {}", type, e);
            return new ArrayList<>();
        }
    }

    @Override
    public byte[] getFileContent(String type, String filename) throws IOException {
        if (!typeToFolder.containsKey(type)) {
            throw new IllegalArgumentException("Invalid file type: " + type);
        }

        Path filePath = Paths.get(basePath, typeToFolder.get(type), filename);

        if (!Files.exists(filePath) || !Files.isRegularFile(filePath)) {
            throw new IOException("File not found: " + filename);
        }

        return Files.readAllBytes(filePath);
    }

    @Override
    public boolean fileExists(String type, String filename) {
        if (!typeToFolder.containsKey(type)) {
            return false;
        }

        Path filePath = Paths.get(basePath, typeToFolder.get(type), filename);
        return Files.exists(filePath) && Files.isRegularFile(filePath);
    }

    @Override
    public String getContentType(String filename) {
        if (filename == null || !filename.contains(".")) {
            return "application/octet-stream";
        }

        String extension = filename.substring(filename.lastIndexOf(".") + 1).toLowerCase();

        switch (extension) {

            case "pdf":
                return "application/pdf";
            case "doc":
                return "application/msword";
            case "docx":
                return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            case "xls":
                return "application/vnd.ms-excel";
            case "xlsx":
                return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            case "ppt":
                return "application/vnd.ms-powerpoint";
            case "pptx":
                return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
            case "txt":
                return "text/plain";
            case "rtf":
                return "application/rtf";
            case "msg":
                return "application/vnd.ms-outlook";
            case "jpg":
            case "jpeg":
                return "image/jpeg";
            case "png":
                return "image/png";
            case "gif":
                return "image/gif";
            case "bmp":
                return "image/bmp";
            case "svg":
                return "image/svg+xml";
            case "mp4":
                return "video/mp4";
            case "avi":
                return "video/avi";
            case "mov":
                return "video/quicktime";
            case "wmv":
                return "video/x-ms-wmv";
            case "csv":
                return "text/csv";
            case "json":
                return "application/json";
            case "xml":
                return "application/xml";
            case "html":
            case "htm":
                return "text/html";
            case "css":
                return "text/css";
            case "js":
                return "application/javascript";
            default:
                return "application/octet-stream";
        }
    }

    @Override
    public boolean isOfficeFile(String filename) {
        if (filename == null || !filename.contains(".")) {
            return false;
        }

        String ext = filename.substring(filename.lastIndexOf(".")).toLowerCase();
        return ext.equals(".doc") || ext.equals(".docx") ||
                ext.equals(".xls") || ext.equals(".xlsx") ||
                ext.equals(".ppt") || ext.equals(".pptx") ||
                ext.equals(".rtf");
    }

    @Override
    public String determineTypeFromFilename(String filename) {
        if (filename == null || !filename.contains(".")) {
            return "documents";
        }

        String ext = filename.substring(filename.lastIndexOf(".") + 1).toLowerCase();

        switch (ext) {

            case "jpg": case "jpeg": case "png": case "gif": case "bmp": case "svg":
                return "images";

            case "mp4": case "avi": case "mov": case "wmv":
                return "videos";

            case "doc": case "docx": case "xls": case "xlsx": case "ppt": case "pptx": case "rtf":
                return "office";

            case "pdf":
                return "pdfs";

            case "txt": case "csv": case "log": case "json": case "xml": case "html": case "htm": case "css": case "js":
                return "text";

            default:
                return "documents";
        }
    }

    @Override
    public void serveOfficeFileInline(String type, String filename, HttpServletRequest request,
                                       HttpServletResponse response) throws IOException {
        try {
            long fileSize = getFileSize(type, filename);
            boolean isLargeFile = fileSize > 25 * 1024 * 1024; // 25MB

            if (filename.toLowerCase().endsWith(".pdf")) {
                servePdfWithViewer(type, filename, request, response);
                return;
            }

            String token = utilityMethodsService.generateToken(filename, type);
            String publicUrl = utilityMethodsService.getPublicFileUrlWithToken(filename, token, request);

            serveOfficeFileWithViewerOptions(type, filename, request, response, isLargeFile, publicUrl);

        } catch (Exception e) {
            log.error("Error serving office file inline: {} for client IP: {}", filename, request.getRemoteAddr(), e);
            serveFallbackDownload(type, filename, response);
        }
    }

    private void servePdfWithViewer(String type, String filename, HttpServletRequest request,
                                    HttpServletResponse response) throws IOException {
        String htmlContent = PdfHtmlGenerator.createPdfViewerHtml(filename, type, request, utilityMethodsService.getPublicFileUrl(type, filename, request));
        serveHtmlContent(htmlContent, filename, response);
        log.info("PDF served with PDF.js viewer: {}", filename);
    }

    private void serveOfficeFileWithViewerOptions(String type, String filename, HttpServletRequest request,
                                                  HttpServletResponse response, boolean isLargeFile, String publicFileUrl)
            throws IOException {
        String htmlContent = OfficeHtmlGenerator.createOfficeMultiViewerHtml(filename, type, request, isLargeFile,
                publicFileUrl, officeViewerUrl);
        serveHtmlContent(htmlContent, filename, response);
    }


    private void serveHtmlContent(String htmlContent, String filename, HttpServletResponse response)
            throws IOException {
        response.setContentType("text/html; charset=UTF-8");
        byte[] htmlBytes = htmlContent.getBytes(StandardCharsets.UTF_8);
        response.setContentLength(htmlBytes.length);
        response.setHeader("Content-Disposition",
                "inline; filename=\"" + URLEncoder.encode(filename + "_viewer.html", StandardCharsets.UTF_8) + "\"");
        response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");

        try (java.io.PrintWriter writer = response.getWriter()) {
            writer.write(htmlContent);
        }
    }

    @Override
    public void serveFallbackDownload(String type, String filename, HttpServletResponse response)
            throws IOException {
        byte[] content = getFileContent(type, filename);
        if (content == null) {
            log.error("File not found: {} in type: {}", filename, type);
            response.sendError(HttpStatus.NOT_FOUND.value(), "File not found");
            return;
        }

        String contentType = getContentType(filename);

        response.setContentType(contentType);
        response.setContentLength(content.length);
        response.setHeader("Content-Disposition",
                "attachment; filename=\"" + URLEncoder.encode(filename, StandardCharsets.UTF_8) + "\"");

        try (OutputStream out = response.getOutputStream()) {
            out.write(content);
        }

        log.info("Office file served as fallback download: {}", filename);
    }

    @Override
    public void serveMsgFile(String type, String filename, HttpServletResponse response)
            throws IOException {

        byte[] content = getFileContent(type, filename);
        if (content == null) {
            response.sendError(HttpServletResponse.SC_NOT_FOUND, "File not found");
            return;
        }

        // If the file is already an .eml, skip conversion and stream directly
        if (filename.toLowerCase().endsWith(".eml")) {
            response.setContentType("message/rfc822");
            response.setHeader("Content-Disposition", "inline; filename=\"" + filename + "\"");

            try (OutputStream out = response.getOutputStream()) {
                out.write(content);
            }
            return;
        }

        // Otherwise handle .msg → .eml conversion
        try (InputStream in = new ByteArrayInputStream(content)) {
            MsgParser parser = new MsgParser();
            Message msg = parser.parseMsg(in);

            Session session = Session.getDefaultInstance(new Properties());
            MimeMessage mimeMessage = new MimeMessage(session);

            if (msg.getFromEmail() != null) {
                mimeMessage.setFrom(new InternetAddress(msg.getFromEmail(), msg.getFromName()));
            }
            if (msg.getToEmail() != null) {
                mimeMessage.setRecipients(javax.mail.Message.RecipientType.TO,
                        InternetAddress.parse(msg.getToEmail()));
            }
            mimeMessage.setSubject(msg.getSubject());
            mimeMessage.setText(msg.getBodyText()); // Use setContent(msg.getBodyHTML(), "text/html") if needed

            response.setContentType("message/rfc822");
            response.setHeader("Content-Disposition",
                    "inline; filename=\"" + filename.replace(".msg", ".eml") + "\"");

            try (OutputStream out = response.getOutputStream()) {
                mimeMessage.writeTo(out);
            }
        } catch (MessagingException e) {
            throw new IOException("Failed to convert .msg to .eml", e);
        }
    }

    @Override
    public void serveRegularFile(String type, String filename, HttpServletResponse response)
            throws IOException {
        byte[] content = getFileContent(type, filename);
        if (content == null) {
            log.error("File not found: {} in type: {}", filename, type);
            response.sendError(HttpStatus.NOT_FOUND.value(), "File not found");
            return;
        }

        String contentType = getContentType(filename);

        response.setContentType(contentType);
        response.setContentLength(content.length);
        response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");
        response.setHeader("Content-Disposition",
                "inline; filename=\"" + URLEncoder.encode(filename, StandardCharsets.UTF_8) + "\"");

        try (OutputStream out = response.getOutputStream()) {
            out.write(content);
        }

        log.info("Regular file served: {}", filename);
    }

    private long getFileSize(String type, String filename) throws IOException {
        byte[] content = getFileContent(type, filename);
        if (content == null) {
            throw new IOException("File not found: " + filename);
        }
        return content.length;
    }
}
