package com.muriithi.filesserver.renderviaweb;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.coyote.BadRequestException;

import java.io.IOException;

public interface RenderWebDocumentService {

    void renderThumbNailLocally(byte[] fileContent, String fileName, String fileContentType, HttpServletResponse response) throws BadRequestException, IOException;
}
