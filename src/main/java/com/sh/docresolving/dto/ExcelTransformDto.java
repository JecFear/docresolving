package com.sh.docresolving.dto;

import lombok.Data;
import org.springframework.web.multipart.MultipartFile;

@Data
public class ExcelTransformDto {

    private String fileIn;

    private MultipartFile multipartFile;

    private String fileout;

    private PrintSetup printSetup;
}
