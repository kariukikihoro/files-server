package com.muriithi.filesserver.serve;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface FileService {

    Map<String, Object> getAvailableTypes();

    List<String> getFilesByType(String type);

    byte[] getFileContent(String type, String filename) throws IOException;

    boolean fileExists(String type, String filename);

    String getContentType(String filename);

    boolean isOfficeFile(String filename);

    String determineTypeFromFilename(String filename);

    void serveOfficeFileInline(String type, String filename, HttpServletRequest request, HttpServletResponse response) throws IOException;

    void serveFallbackDownload(String type, String filename, HttpServletResponse response) throws IOException;

    void serveMsgFile(String type, String filename, HttpServletResponse response) throws IOException;

    void serveRegularFile(String type, String filename, HttpServletResponse response) throws IOException;
}
