package com.muriithi.filesserver.upload;

import lombok.RequiredArgsConstructor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("/api/upload")
@RequiredArgsConstructor
public class UploadController {

    private static final Logger log = LoggerFactory.getLogger(UploadController.class);

    private final UploadService uploadService;

    @PostMapping
    public ResponseEntity<Map<String, Object>> uploadFile(
            @RequestParam("file") MultipartFile file,
            @RequestParam(required = false) String type) {

        log.info("Upload request received for file: {}, type: {}",
                file.getOriginalFilename(), type);

        try {
            Map<String, Object> response;

            if (type != null && !type.trim().isEmpty()) {
                response = uploadService.uploadFile(type, file);
            } else {
                response = uploadService.uploadFileAutoDetect(file);
            }

            return ResponseEntity.ok(response);

        } catch (IllegalArgumentException e) {
            log.error("Invalid upload request: {}", e.getMessage());
            Map<String, Object> errorResponse = new HashMap<>();
            errorResponse.put("success", false);
            errorResponse.put("error", e.getMessage());
            return ResponseEntity.badRequest().body(errorResponse);

        } catch (Exception e) {
            log.error("Upload failed for file: {}", file.getOriginalFilename(), e);
            Map<String, Object> errorResponse = new HashMap<>();
            errorResponse.put("success", false);
            errorResponse.put("error", "Upload failed: " + e.getMessage());
            return ResponseEntity.internalServerError().body(errorResponse);
        }
    }

    @PostMapping("/batch")
    public ResponseEntity<Map<String, Object>> uploadMultipleFiles(
            @RequestParam("files") MultipartFile[] files,
            @RequestParam(required = false) String type) {

        log.info("Batch upload request received for {} files, type: {}", files.length, type);

        Map<String, Object> response = new HashMap<>();
        Map<String, Object> results = new HashMap<>();
        int successCount = 0;
        int failureCount = 0;

        for (MultipartFile file : files) {
            try {
                Map<String, Object> fileResult;

                if (type != null && !type.trim().isEmpty()) {
                    fileResult = uploadService.uploadFile(type, file);
                } else {
                    fileResult = uploadService.uploadFileAutoDetect(file);
                }

                results.put(file.getOriginalFilename(), fileResult);
                successCount++;

            } catch (Exception e) {
                log.error("Failed to upload file: {}", file.getOriginalFilename(), e);
                Map<String, Object> errorResult = new HashMap<>();
                errorResult.put("success", false);
                errorResult.put("error", e.getMessage());
                results.put(file.getOriginalFilename(), errorResult);
                failureCount++;
            }
        }

        response.put("totalFiles", files.length);
        response.put("successCount", successCount);
        response.put("failureCount", failureCount);
        response.put("results", results);

        return ResponseEntity.ok(response);
    }
}