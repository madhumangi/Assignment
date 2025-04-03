package com.claritysystems.XLS.controller;

import com.claritysystems.XLS.service.EWNWorkStreamService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping("/ewn")
public class EWNWorkStreamController {

    @Autowired
    private EWNWorkStreamService ewNworkstreamService;

    @PostMapping("/newRules")
    public ResponseEntity<ByteArrayResource> newRules(@RequestParam("inputFile") MultipartFile inputFile){

        try {
            byte[] processedFile = ewNworkstreamService.processXlsFile(inputFile);
            ByteArrayResource resource = new ByteArrayResource(processedFile);

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=EWNworkstreamAutomationOutput.xlsx");
            return ResponseEntity.ok()
                    .headers(headers)
                    .contentLength(processedFile.length)
                    .contentType(org.springframework.http.MediaType.APPLICATION_OCTET_STREAM)
                    .body(resource);
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
}
