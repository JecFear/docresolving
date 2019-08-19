package com.sh.docresolving.service;

import com.sh.docresolving.dto.ExcelTransformDto;
import com.sh.docresolving.dto.PrintSetup;
import com.sh.docresolving.utils.Excel2Pdf;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;
import java.util.Date;

@Service
public class ExcelResolvingService {

    @Autowired
    private FastDFSService fastDFSService;

    public String excelToPdf(String fileIn, String fileOut, PrintSetup printSetup) throws Exception{
        fileIn = fastDFSService.downloadFile(fileIn,"C:\\excel");
        String outFileName = System.currentTimeMillis()+".pdf";
        fileOut = StringUtils.hasText(fileOut)?fileOut: Excel2Pdf.checkFileOutPathAndOut("C:\\pdf"+ File.separator+outFileName);
        File file = new File(fileOut);
        Excel2Pdf.excel2Pdf(fileIn,fileOut , printSetup);
        Assert.isTrue(file.exists(),"未能成功转换出PDF，请联系管理员查询原因!");
        String fastOutUrl = "";
        try {
            fastOutUrl = fastDFSService.uploadFile(file);
            System.out.println("上传成功..................");
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            //if(StringUtils.hasText(fastOutUrl)) file.delete();
            return fastOutUrl;
        }
    }
}
