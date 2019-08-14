package com.sh.docresolving.controller;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.sh.docresolving.dto.PrintSetup;
import com.sh.docresolving.service.ConvertRecordService;
import com.sh.docresolving.service.ExcelResolvingService;
import com.sh.docresolving.service.FastDFSService;
import com.sh.docresolving.utils.Excel2Pdf;
import io.swagger.annotations.ApiOperation;
import net.shouhouzn.lims.ms.umps.service.FeignUpmsService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;

@RestController
@RequestMapping("/excel-convert")
public class ExcelResolvingController {

    @Autowired
    private FastDFSService fastDFSService;
    @Autowired
    private ConvertRecordService convertRecordService;
    @Autowired
    private ExcelResolvingService excelResolvingService;
    @Autowired
    private FeignUpmsService feignUpmsService;

    @ApiOperation(value = "excel转pdf", notes = "jacob", response = ResponseEntity.class)
    @RequestMapping(value = "/excel-to-pdf", method = RequestMethod.POST)
    public ResponseEntity Excel2Pdf(@RequestParam("file") MultipartFile multipartFile,@RequestParam("printSetup") String json,@RequestHeader(value = "token",required = false) String token){
       try {
           ObjectMapper objectMapper = new ObjectMapper();
           PrintSetup printSetup = objectMapper.readValue(json, PrintSetup.class);
           String fileOutDir = Thread.currentThread().getContextClassLoader().getResource("fileOut").getPath();
           String fileOutDirTruePath = Excel2Pdf.checkFileOutPathAndOut(fileOutDir);
           String fileOutPath = fileOutDirTruePath+File.separator+System.currentTimeMillis()+".xlsx";
           File xlsxFile = new File(fileOutPath);
           multipartFile.transferTo(xlsxFile);
           String fastOutUrl = excelResolvingService.excelToPdf(fileOutPath,null,printSetup);
           if(!StringUtils.hasText(fastOutUrl)){
                throw new IllegalStateException("fastDfs上传失败，请重试或联系管理员!");
           }
           return ResponseEntity.ok(fastOutUrl);
       }catch (Exception e) {
           e.printStackTrace();
           return ResponseEntity.badRequest().body(e.getMessage());
       }
    }
}
