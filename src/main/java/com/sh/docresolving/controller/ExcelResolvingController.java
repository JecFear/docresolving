package com.sh.docresolving.controller;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.sh.docresolving.dto.PrintSetup;
import com.sh.docresolving.service.ExcelResolvingService;
import com.sh.docresolving.service.FastDFSService;
import com.sh.docresolving.utils.Excel2Pdf;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/excel-convert")
public class ExcelResolvingController {

    @Autowired
    private ExcelResolvingService excelResolvingService;
    @Autowired
    private FastDFSService fastDFSService;

    @ApiOperation(value = "excel转pdf", notes = "jacob", response = ResponseEntity.class)
    @RequestMapping(value = "/excel-to-pdf", method = RequestMethod.POST)
    public String excel2Pdf(@RequestParam("file") String filein, @RequestParam("json") String json) throws Exception{
       try {
           System.out.println("请求进入..................");
           ObjectMapper objectMapper = new ObjectMapper();
           PrintSetup printSetup = objectMapper.readValue(json, PrintSetup.class);
           String fastOutUrl = excelResolvingService.excelToPdf(filein,null,printSetup);
           if(!StringUtils.hasText(fastOutUrl)){
               throw new IllegalStateException("fastDfs上传失败，请重试或联系管理员!");
           }
           System.out.println(fastOutUrl);
           return fastOutUrl;
       }catch (Exception e) {
           e.printStackTrace();
           throw e;
       }
    }
}
