package com.muriithi.filesserver.upload;

import com.muriithi.filesserver.serve.FileService;
import lombok.RequiredArgsConstructor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

@Service
@RequiredArgsConstructor
public class UploadServiceImpl implements UploadService {

    private static final Logger log = LoggerFactory.getLogger(UploadServiceImpl.class);

    @Value("${file.storage.base-path:./files}")
    private String basePath;

    private final FileService fileService;

    private final Map<String, String> typeToFolder = Map.of(
            "documents", "documents",
            "images", "images",
            "videos", "videos",
            "office", "office",
            "pdfs", "pdfs",
            "text", "text"
    );

    @Override
    public Map<String, Object> uploadFile(String type, MultipartFile file) throws IOException {

        if (!typeToFolder.containsKey(type)) {
            throw new IllegalArgumentException("Invalid file type: " + type);
        }

        if (file.isEmpty()) {
            throw new IllegalArgumentException("File is empty");
        }

        String filename = file.getOriginalFilename();
        if (filename == null || filename.trim().isEmpty()) {
            throw new IllegalArgumentException("Invalid filename");
        }

        String folder = typeToFolder.get(type);
        Path uploadPath = Paths.get(basePath, folder);

        Files.createDirectories(uploadPath);

        Path filePath = uploadPath.resolve(filename);
        Files.write(filePath, file.getBytes());

        log.info("File uploaded successfully: {} to {}", filename, folder);

        Map<String, Object> response = new HashMap<>();
        response.put("success", true);
        response.put("filename", filename);
        response.put("type", type);
        response.put("folder", folder);
        response.put("size", file.getSize());
        response.put("contentType", fileService.getContentType(filename));
        response.put("path", filePath.toString());

        return response;
    }

    @Override
    public Map<String, Object> uploadFileAutoDetect(MultipartFile file) throws IOException {

        String filename = file.getOriginalFilename();
        String detectedType = fileService.determineTypeFromFilename(filename);

        log.info("Auto-detected type '{}' for file: {}", detectedType, filename);

        return uploadFile(detectedType, file);
    }
}