package com.muriithi.filesserver.serve;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.concurrent.*;

@RestController
@RequestMapping("/api/files")
@RequiredArgsConstructor
@CrossOrigin(origins = "*", methods = {RequestMethod.GET, RequestMethod.HEAD, RequestMethod.OPTIONS})
public class FileController {


    private final FileService fileService;
    private final UtilityMethodsService utilityMethodsService;

    private static final Logger log = LoggerFactory.getLogger(FileController.class);
    private static final long MAX_FILE_SIZE = 100 * 1024 * 1024; // 100MB

    @Value("${file.token.expiry.minutes:5}")
    private int tokenExpiryMinutes;

    @RequestMapping(value = "/**", method = RequestMethod.OPTIONS)
    public ResponseEntity<Void> handleOptions(HttpServletResponse response) {
        utilityMethodsService.setCorsHeaders(response);
        return ResponseEntity.ok().build();
    }

    @GetMapping("/types")
    public ResponseEntity<Map<String, Object>> getAvailableTypes() {
        try {
            Map<String, Object> response = fileService.getAvailableTypes();
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            log.error("Error getting available types", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(Map.of("error", "Failed to retrieve available types"));
        }
    }

    @GetMapping("/list")
    public ResponseEntity<List<String>> getFilesByType(@RequestParam String type) {
        if (!utilityMethodsService.isValidType(type)) {
            return ResponseEntity.badRequest().build();
        }

        try {
            List<String> files = fileService.getFilesByType(type);
            return ResponseEntity.ok(files);
        } catch (Exception e) {
            log.error("Error listing files for type: {}", type, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    @GetMapping("/serve")
    public void serveFile(@RequestParam String type,
                          @RequestParam String filename,
                          HttpServletRequest request,
                          HttpServletResponse response) throws IOException {
        if (type == null || filename == null) {
            response.sendError(HttpStatus.BAD_REQUEST.value(), "Type and filename are required");
            return;
        }
        if (!utilityMethodsService.isValidInput(type, filename)) {
            response.sendError(HttpStatus.BAD_REQUEST.value(), "Invalid parameters");
            return;
        }

        log.info("Serving file: {} from type: {} for client IP: {}", filename, type, request.getRemoteAddr());
        utilityMethodsService.setCorsHeaders(response);

        try {
            byte[] content = fileService.getFileContent(type, filename);
            if (content == null) {
                log.error("File not found: {} in type: {}", filename, type);
                response.sendError(HttpStatus.NOT_FOUND.value(), "File not found");
                return;
            }

            long fileSize = content.length;
            if (fileSize > MAX_FILE_SIZE) {
                log.warn("File too large: {} ({} bytes)", filename, fileSize);
                response.sendError(HttpStatus.PAYLOAD_TOO_LARGE.value(), "File too large");
                return;
            }

            if (fileService.isOfficeFile(filename)) {
                fileService.serveOfficeFileInline(type, filename, request, response);
                return;
            }

            if ((filename.toLowerCase().endsWith(".msg")|| filename.toLowerCase().endsWith(".eml"))) {
                fileService.serveMsgFile(type, filename, response);
                return;
            }

            fileService.serveRegularFile(type, filename, response);

        } catch (Exception e) {
            log.error("Error serving file: {} for client IP: {}", filename, request.getRemoteAddr(), e);
            response.sendError(HttpStatus.INTERNAL_SERVER_ERROR.value(), "Error serving file");
        }
    }

    @GetMapping("/download")
    public void downloadFile(@RequestParam String type,
                             @RequestParam String filename,
                             HttpServletRequest request,
                             HttpServletResponse response) throws IOException {
        if (type == null || filename == null) {
            response.sendError(HttpStatus.BAD_REQUEST.value(), "Type and filename are required");
            return;
        }
        if (!utilityMethodsService.isValidInput(type, filename)) {
            response.sendError(HttpStatus.BAD_REQUEST.value(), "Invalid parameters");
            return;
        }

        log.info("Download requested for file: {} from type: {} for client IP: {}", filename, type, request.getRemoteAddr());
        utilityMethodsService.setCorsHeaders(response);

        try {
            byte[] content = fileService.getFileContent(type, filename);
            if (content == null) {
                log.error("File not found: {} in type: {}", filename, type);
                response.sendError(HttpStatus.NOT_FOUND.value(), "File not found");
                return;
            }

            String contentType = fileService.getContentType(filename);

            response.setContentType(contentType);
            response.setContentLength(content.length);
            response.setHeader("Content-Disposition",
                    "attachment; filename=\"" + URLEncoder.encode(filename, StandardCharsets.UTF_8) + "\"");

            try (OutputStream out = response.getOutputStream()) {
                out.write(content);
            }

            log.info("File downloaded successfully: {}", filename);

        } catch (Exception e) {
            log.error("Error downloading file: {} for client IP: {}", filename, request.getRemoteAddr(), e);
            response.sendError(HttpStatus.INTERNAL_SERVER_ERROR.value(), "Error downloading file");
        }
    }

    @GetMapping("/exists")
    public ResponseEntity<Map<String, Object>> checkFileExists(@RequestParam String type,
                                                               @RequestParam String filename) {
        if (type == null || filename == null) {
            return ResponseEntity.badRequest().body(Map.of("error", "Type and filename are required"));
        }
        if (!utilityMethodsService.isValidInput(type, filename)) {
            return ResponseEntity.badRequest().body(Map.of("error", "Invalid parameters"));
        }

        try {
            boolean exists = fileService.fileExists(type, filename);
            Map<String, Object> response = Map.of(
                    "exists", exists,
                    "type", type,
                    "filename", filename
            );
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            log.error("Error checking file existence: {} in type: {}", filename, type, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(Map.of("error", "Failed to check file existence"));
        }
    }

    @GetMapping("/public-check")
    public ResponseEntity<Map<String, Object>> checkPublicAccess(@RequestParam String type,
                                                                 @RequestParam String filename,
                                                                 HttpServletRequest request) {
        if (type == null || filename == null) {
            return ResponseEntity.badRequest().body(Map.of("error", "Type and filename are required"));
        }
        if (!utilityMethodsService.isValidInput(type, filename)) {
            return ResponseEntity.badRequest().body(Map.of("error", "Invalid parameters"));
        }

        try {
            String fileUrl = utilityMethodsService.getPublicFileUrl(type, filename, request);
            URL testUrl = new URL(fileUrl);

            HttpURLConnection connection = (HttpURLConnection) testUrl.openConnection();
            connection.setRequestMethod("HEAD");
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(5000);

            int responseCode = connection.getResponseCode();
            boolean isPublic = responseCode == 200;

            return ResponseEntity.ok(Map.of(
                    "filename", filename,
                    "isPublic", isPublic,
                    "publicUrl", fileUrl,
                    "message", isPublic ? "File is publicly accessible" : "File is not publicly accessible"
            ));
        } catch (Exception e) {
            log.error("Error checking public access for file: {} for client IP: {}", filename, request.getRemoteAddr(), e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(Map.of(
                            "error", "Verification failed",
                            "message", e.getMessage()
                    ));
        }
    }

    /* ================== Token Management with Security Enhancements ================== */

    @GetMapping("/public-view")
    public void publicFileView(@RequestParam String token,
                               @RequestParam String filename,
                               HttpServletRequest request,
                               HttpServletResponse response) throws IOException {
        if (token == null || filename == null) {
            response.sendError(HttpStatus.BAD_REQUEST.value(), "Token and filename are required");
            return;
        }
        if (!utilityMethodsService.isValidFilename(filename)) {
            response.sendError(HttpStatus.BAD_REQUEST.value(), "Invalid filename");
            return;
        }

        TokenInfo tokenInfo = utilityMethodsService.validateToken(token, filename);
        if (tokenInfo == null) {
            response.sendError(HttpStatus.FORBIDDEN.value(), "Invalid or expired token");
            return;
        }

        utilityMethodsService.setCorsHeaders(response);

        try {
            byte[] content = fileService.getFileContent(tokenInfo.type, filename);
            if (content == null) {
                log.error("File not found: {} in type: {}", filename, tokenInfo.type);
                response.sendError(HttpStatus.NOT_FOUND.value(), "File not found");
                return;
            }

            String contentType = fileService.getContentType(filename);

            response.setContentType(contentType);
            response.setContentLength(content.length);
            response.setHeader("Content-Disposition",
                    "inline; filename=\"" + URLEncoder.encode(filename, StandardCharsets.UTF_8) + "\"");

            try (OutputStream out = response.getOutputStream()) {
                out.write(content);
            }
        } catch (Exception e) {
            log.error("Error serving public file: {} for client IP: {}", filename, request.getRemoteAddr(), e);
            response.sendError(HttpStatus.INTERNAL_SERVER_ERROR.value(), "Error serving file");
        }
    }

    @GetMapping("/generate-token")
    public ResponseEntity<Map<String, Object>> generateAccessToken(@RequestParam String type,
                                                                   @RequestParam String filename,
                                                                   HttpServletRequest request) {
        if (type == null || filename == null) {
            return ResponseEntity.badRequest().body(Map.of("error", "Type and filename are required"));
        }
        if (!utilityMethodsService.isValidInput(type, filename)) {
            return ResponseEntity.badRequest().body(Map.of("error", "Invalid parameters"));
        }

        if (!fileService.fileExists(type, filename)) {
            return ResponseEntity.notFound().build();
        }

        String clientIp = utilityMethodsService.getClientIp(request);
        if (utilityMethodsService.isRateLimited(clientIp)) {
            log.warn("Rate limit exceeded for IP: {}", clientIp);
            return ResponseEntity.status(HttpStatus.TOO_MANY_REQUESTS)
                    .body(Map.of("error", "Too many requests. Try again later."));
        }

        try {
            String token = utilityMethodsService.generateToken(filename, type);
            return ResponseEntity.ok(Map.of(
                    "token", token,
                    "expiresIn",  tokenExpiryMinutes + " minutes",
                    "publicUrl", utilityMethodsService.getPublicFileUrlWithToken(filename, token, request)
            ));
        } catch (Exception e) {
            log.error("Error generating token for file: {} for client IP: {}", filename, clientIp, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(Map.of("error", "Failed to generate token"));
        }
    }

}