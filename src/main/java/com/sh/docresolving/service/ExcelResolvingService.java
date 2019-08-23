package com.sh.docresolving.service;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import com.itextpdf.text.pdf.parser.PdfTextExtractor;
import com.sh.docresolving.dto.ExcelTransformDto;
import com.sh.docresolving.dto.PrintSetup;
import com.sh.docresolving.utils.Excel2Pdf;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;

import javax.print.Doc;
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
        fileOut = StringUtils.hasText(fileOut)?fileOut:"C:\\pdf"+ File.separator+outFileName;
        File file = new File(fileOut);
        Excel2Pdf.excel2Pdf(fileIn,fileOut , printSetup);
        Assert.isTrue(file.exists(),"未能成功转换出PDF，请联系管理员查询原因!");
        String ousFileName = "";
        if(printSetup.needPageNum()){
            ousFileName = System.currentTimeMillis()+".pdf";
            addPageEvent(fileOut,"C:\\pdf"+ File.separator+ousFileName,printSetup.pageNumStart());
            file = new File("C:\\pdf"+ File.separator+ousFileName);
        }
        String fastOutUrl = "";
        try {
            fastOutUrl = fastDFSService.uploadFile(file);
            System.out.println("上传成功..................");
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if(StringUtils.hasText(fastOutUrl)&&StringUtils.hasText(ousFileName)) new File(fileOut).delete();
            return fastOutUrl;
        }
    }

    protected void addPageEvent(String fileOut,String ousFile,Integer pageNumStart) throws IOException,DocumentException{
        ///////////////////////////
        PdfReader pdfReader = new PdfReader(fileOut);
        PdfStamper pdfStamper = new PdfStamper(pdfReader,new FileOutputStream(ousFile));
        BaseFont baseFont = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
        int pageNum = pdfReader.getNumberOfPages();
        for(int i=1;i<=pageNum;i++){
            if(i>=pageNumStart){
                PdfContentByte pdfContentByte = pdfStamper.getOverContent(i);
                PdfDocument document = pdfContentByte.getPdfDocument();
                String text = "第 " + (i-pageNumStart+1) + " 页"+"  共 "+(pageNum-pageNumStart+1)+" 页";
                pdfContentByte.beginText();
                pdfContentByte.setFontAndSize(baseFont , 10);
                float bottom = document.bottom(0-document.bottom()+6);
                float left = 0;
                float right = pdfStamper.getImportedPage(pdfReader,i).getBoundingBox().getWidth();
                pdfContentByte.showTextAligned(PdfContentByte.ALIGN_CENTER, text, (right + left) / 2, bottom, 0);
                pdfContentByte.endText();
            }
        }
        pdfStamper.close();
        pdfReader.close();
    }
}
