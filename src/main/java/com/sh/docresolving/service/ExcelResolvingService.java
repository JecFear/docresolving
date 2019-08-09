package com.sh.docresolving.service;

import com.sh.docresolving.dto.ExcelTransformDto;
import com.sh.docresolving.utils.Excel2Pdf;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;

import java.io.*;

@Service
public class ExcelResolvingService {

    @Autowired
    private FastDFSService fastDFSService;

    public String excelToPdf(ExcelTransformDto excelTransformDto) throws Exception{
        String outFileName = System.currentTimeMillis()+".pdf";
        String fileOut = StringUtils.hasText(excelTransformDto.getFileout())?excelTransformDto.getFileout(): Excel2Pdf.checkFileOutPathAndOut(Thread.currentThread().getContextClassLoader().getResource("fileOut").getPath()+ File.separator+outFileName);
        Excel2Pdf.excel2Pdf(excelTransformDto.getFileIn(),fileOut , excelTransformDto.getPrintSetup());
        File file = new File(fileOut);
        Assert.isTrue(file.exists(),"未能成功转换出PDF，请联系管理员查询原因!");
        String fastOutUrl = "";
        try {
            fastOutUrl = fastDFSService.uploadFile(file);
        }catch (Exception e) {
            e.printStackTrace();
            /*fastOutUrl = fastDFSService.uploadFile(file);*/
        }
        if(StringUtils.hasText(fastOutUrl)) file.delete();
        return fastOutUrl;
    }
}
