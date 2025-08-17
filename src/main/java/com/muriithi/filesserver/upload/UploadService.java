package com.muriithi.filesserver.upload;

import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.Map;

public interface UploadService {

    Map<String, Object> uploadFile(String type, MultipartFile file) throws IOException;

    Map<String, Object> uploadFileAutoDetect(MultipartFile file) throws IOException;
}
