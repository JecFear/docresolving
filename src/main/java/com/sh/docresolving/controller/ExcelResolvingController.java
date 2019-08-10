package com.sh.docresolving.controller;

import com.sh.docresolving.dto.ExcelTransformDto;
import com.sh.docresolving.service.ConvertRecordService;
import com.sh.docresolving.service.ExcelResolvingService;
import com.sh.docresolving.service.FastDFSService;
import com.sh.docresolving.utils.Excel2Pdf;
import com.sh.docresolving.utils.FastDFSClient;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;

@RestController("/excel-convert")
public class ExcelResolvingController {

    @Autowired
    private FastDFSService fastDFSService;
    @Autowired
    private ConvertRecordService convertRecordService;
    @Autowired
    private ExcelResolvingService excelResolvingService;

    @ApiOperation(value = "excel转pdf", notes = "jacob", response = ResponseEntity.class)
    @RequestMapping(value = "/excel-to-pdf", method = RequestMethod.POST)
    public ResponseEntity Excel2Pdf(@RequestBody ExcelTransformDto excelTransformDto){
       try {
           String localFilePath = fastDFSService.downloadFile(excelTransformDto.getFileIn(),Excel2Pdf.checkFileOutPathAndOut(Thread.currentThread().getContextClassLoader().getResource("fileOut").getPath()));
           excelTransformDto.setFileIn(localFilePath);
           String fastOutUrl = excelResolvingService.excelToPdf(excelTransformDto);
           convertRecordService.saveConvertRecordByExcel2Pdf(excelTransformDto,fastOutUrl);
           if(!StringUtils.hasText(fastOutUrl)){
                throw new IllegalStateException("fastDfs上传失败，请重试或联系管理员!");
           }
           return ResponseEntity.ok(fastOutUrl);
       }catch (Exception e) {
           return ResponseEntity.badRequest().body(e.getMessage());
       }
    }
}
