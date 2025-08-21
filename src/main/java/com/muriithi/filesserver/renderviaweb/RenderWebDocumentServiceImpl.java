package com.muriithi.filesserver.renderviaweb;

import com.auxilii.msgparser.Message;
import com.auxilii.msgparser.MsgParser;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.IOUtils;
import org.apache.coyote.BadRequestException;
import org.springframework.beans.factory.annotation.Autowired;
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
import java.util.Properties;


@Service
@Slf4j
public class RenderWebDocumentServiceImpl implements RenderWebDocumentService {

    @Autowired
    private CsvDocumentRenderer csvDocumentRenderer;


    @Override
    public void renderThumbNailLocally(byte[] fileContent, String fileName, String fileContentType, HttpServletResponse response) throws BadRequestException, IOException {

        /***local testing
         UtilityMethodsService.FileTestData fileTestData = utilityMethodsService.getLocalTestData();
         renderThumbnail(response, fileTestData.fileName, fileTestData.fileContent, fileTestData.fileContentType);
         */
        log.info(":::::::::::  rendering thumbnail from sybrin case fileName  ==>  {} \n," +
                "###### contentType ===> {}, fileSize ===> {}, response {}", fileName, fileContentType, getFileSize(fileContent), response.toString());
        renderThumbnail(response, fileName, fileContent, fileContentType);
    }

    private void renderThumbnail(HttpServletResponse response, String fileName,
                                 byte[] fileContent, String fileContentType) throws IOException {

        if (fileContent == null || fileContent.length == 0) {
            response.sendError(HttpStatus.NOT_FOUND.value(), "Empty document");
            return;
        }

        try {
            if (isOfficeFile(fileName)) {

                log.info("===== rendering an office file ({}) ===", fileName);

                try {

                    OfficeDocumentRenderer renderer = new OfficeDocumentRenderer();
                    byte[] renderedContent;
                    String contentType;
                    try {

                        renderedContent = renderer.renderDocument(fileContent, fileName, "html");
                        contentType = "text/html; charset=UTF-8";
                    } catch (Exception htmlException) {

                        try {

                            renderedContent = renderer.renderDocument(fileContent, fileName, "pdf");
                            contentType = "application/pdf";
                        } catch (Exception pdfException) {
                            renderedContent = fileContent;
                            contentType = ContentTypeHelper.getContentType(fileName);
                        }
                    }
                    response.setContentType(contentType);
                    response.setContentLength(renderedContent.length);

                    try (OutputStream out = response.getOutputStream()) {

                        out.write(renderedContent);
                    }
                } catch (Exception e) {

                    e.printStackTrace();
                    response.sendError(HttpServletResponse.SC_INTERNAL_SERVER_ERROR, "Error rendering document: " + e.getMessage());
                }
            }

            if (fileName.toLowerCase().endsWith(".csv")) {

                log.info("===== rendering a csv file ({}) ===", fileName);

                try {

                    byte[] htmlContent = csvDocumentRenderer.renderCsvDocument(fileContent, fileName);

                    response.setContentType("text/html");
                    response.setCharacterEncoding("UTF-8");
                    response.setContentLength(htmlContent.length);
                    response.getOutputStream().write(htmlContent);
                    response.getOutputStream().flush();

                } catch (Exception e) {
                    log.error("Error rendering CSV document {}. Falling back to download.", fileName, e);
                    serveFallbackDownload(fileContent, fileName, fileContentType, response);
                }
            }

            if (fileName.toLowerCase().endsWith(".msg") || fileName.toLowerCase().endsWith(".eml")) {
                serveMsgFile(fileContent, fileName, response);
            }

            serveRegularFile(fileContent, fileName, fileContentType, response);

        } catch (Exception e) {
            log.error("Unexpected error rendering {}. Using fallback download.", fileName, e);
            serveFallbackDownload(fileContent, fileName, fileContentType, response);
        }
    }


    private void serveFallbackDownload(byte[] fileContent, String fileName, String fileContentType, HttpServletResponse response)
            throws IOException {

        if (fileContent == null) {
            log.error("File not found: {} ", fileName);
            response.sendError(HttpStatus.NOT_FOUND.value(), "File not found");
            return;
        }

        response.setContentType(fileContentType);
        response.setContentLength(fileContent.length);
        response.setHeader("Content-Disposition",
                "attachment; fileName=\"" + URLEncoder.encode(fileName, "UTF-8") + "\"");

        try (OutputStream out = response.getOutputStream()) {
            out.write(fileContent);
        }

        log.info("Office file served as fallback download: {}", fileName);
    }


    private void serveMsgFile(byte[] fileContent, String fileName, HttpServletResponse response) throws IOException {

        log.info("===== rendering a message file ({}) ===", fileName);

        if (fileContent == null || fileContent.length == 0) {
            response.sendError(HttpStatus.NOT_FOUND.value(), "Empty MSG file");
            return;
        }

        if (fileName.toLowerCase().endsWith(".eml")) {
            response.setContentType("message/rfc822");
            response.setHeader("Content-Disposition", "inline; fileName=\"" + fileName + "\"");

            try (OutputStream out = response.getOutputStream()) {
                out.write(fileContent);
            }
            return;
        }

        try (InputStream in = new ByteArrayInputStream(fileContent)) {
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
            mimeMessage.setText(msg.getBodyText());

            response.setContentType("message/rfc822");
            response.setHeader("Content-Disposition",
                    "inline; fileName=\"" + fileName.replace(".msg", ".eml") + "\"");

            try (OutputStream out = response.getOutputStream()) {
                mimeMessage.writeTo(out);
            }
        } catch (MessagingException e) {

            e.printStackTrace();
            throw new IOException("Failed to convert .msg to .eml", e);
        }
    }


    private void serveRegularFile(byte[] fileContent, String fileName, String fileContentType, HttpServletResponse response) throws IOException {

        log.info("===== rendering regular file ({}) ===", fileName);

        try {

            if (fileContent == null || fileContent.length == 0) {
                response.sendError(HttpStatus.NOT_FOUND.value(), "Empty document");
                return;
            }

            if (fileName.toLowerCase().endsWith(".jpeg") || fileName.toLowerCase().endsWith(".jpg") || fileName.toLowerCase().endsWith(".png")) {
                //currently, sybrin is converting images to pdf
                fileContentType = ContentTypeHelper.getContentType("fileName.pdf");
            }

            // Set universal headers
            response.setContentType(fileContentType);
            response.setContentLength(fileContent.length);
            response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
            response.setHeader("Pragma", "no-cache");
            response.setHeader("Expires", "0");
            response.setHeader("Content-Disposition",
                    "inline; fileName=\"" + URLEncoder.encode(fileName, "UTF-8").replace("+", "%20") + "\"");

            try (OutputStream out = response.getOutputStream()) {
                out.write(fileContent);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private boolean isOfficeFile(String fileName) {

        if (fileName == null || !fileName.contains(".")) {
            return false;
        }

        String ext = fileName.substring(fileName.lastIndexOf(".")).toLowerCase();
        return ext.equals(".doc") || ext.equals(".docx") ||
                ext.equals(".xls") || ext.equals(".xlsx") ||
                ext.equals(".ppt") || ext.equals(".pptx") ||
                ext.equals(".rtf");
    }

    private long getFileSize(byte[] fileContent) throws IOException {
        if (fileContent == null) {
            log.error("====== Invalid document file content, null content parsed");
            throw new IOException("File not found or empty: ");
        }
        return fileContent.length;
    }

}
