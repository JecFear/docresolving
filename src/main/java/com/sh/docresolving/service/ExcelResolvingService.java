package com.sh.docresolving.service;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.*;
import com.sh.docresolving.dto.PrintConvertDto;
import com.sh.docresolving.dto.PrintSetup;
import com.sh.docresolving.utils.Excel2Pdf;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

@Service
public class ExcelResolvingService {

    @Autowired
    private FastDFSService fastDFSService;
    @Autowired
    private PdfResolvingService pdfResolvingService;

    public String excelToPdf(String fileIn, String fileOut, PrintSetup printSetup) throws Exception{
        fileIn = fastDFSService.downloadFile(fileIn,"C:\\excel");
        String outFileName = System.currentTimeMillis()+".pdf";
        fileOut = StringUtils.hasText(fileOut)?fileOut:"C:\\pdf"+ File.separator+outFileName;
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
            //if(StringUtils.hasText(fastOutUrl)) new File(fileOut).delete();
            return fastOutUrl;
        }
    }

    public String excelToPdfMultiple(List<String> filesIn,List<PrintSetup> printSetups) throws Exception{
        List<Future<String>> futureList = new ArrayList<>();
        ExecutorService executor = Executors.newFixedThreadPool(filesIn.size());
        final List<String> finalFilesIn = filesIn;
        for(String fileIn:finalFilesIn){
            Callable<String> task = new Callable<String>() {
                @Override
                public String call() throws Exception {
                    String innerFileIn = fastDFSService.downloadFile(fileIn,"C:\\excel");
                    //获取fastdfs名
                    int fileSep = innerFileIn.lastIndexOf(File.separator);
                    String excelName = innerFileIn.substring(fileSep+1);
                    int a = excelName.lastIndexOf(".");
                    String excelNameWithNoSuffix = excelName.substring(0,a);
                    int index = finalFilesIn.indexOf(fileIn);
                    String outFileName = System.currentTimeMillis()+excelNameWithNoSuffix+".pdf";
                    String fileOut = "C:\\pdf"+ File.separator+outFileName;
                    Excel2Pdf.excel2Pdf(innerFileIn,fileOut,printSetups.get(index));
                    return fileOut;
                }
            };
            Future<String> future = executor.submit(task);
            futureList.add(future);
        }
        executor.shutdown();
        Set<String> fileOuts = new HashSet<>();
        for (Future<String> future:futureList){
            try {
                fileOuts.add(future.get());
            }catch (Exception e){
                e.printStackTrace();
                throw e;
            }
        }
        System.out.println("fileOuts:"+fileOuts);
        String mergeFileOut = pdfResolvingService.pdfMerge(new ArrayList<>(fileOuts));
        return mergeFileOut;
    }

    public PrintConvertDto excelToPdfReturnPrintConvertDto(String fileIn, String fileOut, PrintSetup printSetup)throws Exception{
        PrintConvertDto printConvertDto = new PrintConvertDto();
        printConvertDto.setOriginalUrl(fileIn);
        fileIn = fastDFSService.downloadFile(fileIn,"C:\\excel");
        printConvertDto.setOriginalPath(fileIn);
        String outFileName = System.currentTimeMillis()+"_p"+".pdf";
        fileOut = StringUtils.hasText(fileOut)?fileOut:"C:\\pdf"+ File.separator+outFileName;
        printConvertDto.setConvertPath(fileOut);
        File file = new File(fileOut);
        Excel2Pdf.excel2Pdf(fileIn,fileOut , printSetup);
        Assert.isTrue(file.exists(),"未能成功转换出PDF，请联系管理员查询原因!");
        String fastOutUrl = "";
        try {
            fastOutUrl = fastDFSService.uploadFile(file);
            printConvertDto.setConvertPath(fastOutUrl);
            System.out.println("上传成功..................");
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            //if(StringUtils.hasText(fastOutUrl)) new File(fileOut).delete();
            return printConvertDto;
        }
    }
}
