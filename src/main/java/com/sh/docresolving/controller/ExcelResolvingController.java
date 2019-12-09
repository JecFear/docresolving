package com.sh.docresolving.controller;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.sh.docresolving.dto.PrintConvertDto;
import com.sh.docresolving.dto.PrintSetup;
import com.sh.docresolving.service.ExcelResolvingService;
import com.sh.docresolving.service.FastDFSService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Api("excel处理API")
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

    @ApiOperation(value = "excel转pdf(固定格式,携带本地地址和fsds地址等)", notes = "jacob", response = ResponseEntity.class)
    @RequestMapping(value = "/excel-to-pdf-format", method = RequestMethod.POST)
    public PrintConvertDto excel2PdfFormat(@RequestParam("file") String filein, @RequestParam("json") String json) throws Exception{
        try {
            System.out.println("请求进入..................");
            ObjectMapper objectMapper = new ObjectMapper();
            PrintSetup printSetup = objectMapper.readValue(json, PrintSetup.class);
            PrintConvertDto printConvertDto = excelResolvingService.excelToPdfReturnPrintConvertDto(filein,null,printSetup);
            System.out.println(printConvertDto);
            return printConvertDto;
        }catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }

    @ApiOperation(value = "excel转pdf/多个合并", notes = "jacob", response = ResponseEntity.class)
    @RequestMapping(value = "/excel-to-pdf-multiple", method = RequestMethod.POST)
    public String excel2PdfMultiple(@RequestBody ArrayList<PrintSetup> jsons, @RequestParam("files") List<String> filesIn) throws Exception{
        try {
            System.out.println("文件数量:"+ filesIn.size());
            System.out.println("请求进入..................");
            return excelResolvingService.excelToPdfMultiple(filesIn,jsons);
        }catch (Exception e){
            e.printStackTrace();
            throw e;
        }
    }
}
