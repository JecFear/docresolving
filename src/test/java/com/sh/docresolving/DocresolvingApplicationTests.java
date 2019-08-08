package com.sh.docresolving;

import com.itextpdf.text.pdf.BaseFont;
import com.sh.docresolving.entity.PrintSetup;
import com.sh.docresolving.utils.Excel2Pdf;
import com.sh.docresolving.utils.ExcelToPdf;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.awt.*;
import java.io.File;

@RunWith(SpringRunner.class)
@SpringBootTest
public class DocresolvingApplicationTests {

    @Test
    public void contextLoads() {
    }

    @Test
    public void HSSFWORKBOOKTEST() throws Exception{
        String fileIn = "sample1/download.xlsx";
        String uri = this.getClass().getResource(fileIn).getPath();
        String fileOut = uri.replaceAll(".xls$|.xlsx$",".pdf");
        ExcelToPdf.convert(uri,fileOut);
    }

    @Test
    public void HSSFWORKBOOKTESTss() throws Exception{
        String fileIn = "F:\\docresolving\\target\\test-classes\\com\\sh\\docresolving\\sample1\\download.xlsx";
        /*String uri = this.getClass().getResource(fileIn).getPath();
        System.out.println(fileIn);*/
        String fileOut = fileIn.replaceAll(".xls$|.xlsx$",".pdf");
        PrintSetup printSetup = new PrintSetup();
        printSetup.put("sheet1",true);
        printSetup.put("sheet2",false);
        printSetup.put("sheet3",true);
        printSetup.put("Sheet5",false);
        Excel2Pdf.excel2Pdf(fileIn,fileOut,printSetup);
    }

    @Test
    public void sss() throws Exception{
        String fontName = "arial";
        String ttfDir =  Thread.currentThread().getContextClassLoader().getResource("ttf").getPath();
        ttfDir = replaceFirstLineChar(ttfDir);
        String fileDir = ttfDir+File.separator+fontName+".ttf";
        File file = new File(fileDir);
        Font font = java.awt.Font.createFont(0,file);
        String aa = font.getFontName();
        String bb = font.getName();
        System.out.println(aa+":"+bb);
        boolean a = file.exists();
        System.out.println(a);
    }

    @Test
    public void printOutFontFileName() throws Exception{
        String folderDir = Thread.currentThread().getContextClassLoader().getResource("ttf").getPath();
        folderDir = replaceFirstLineChar(folderDir);
        File folder = new File(folderDir);
        String[] filelist = folder.list();
        System.out.println(filelist.length);
        for (int i = 0; i < filelist.length; i++) {
            File ttfFile = new File(folderDir + "\\" + filelist[i]);
            if(!ttfFile.getName().contains("ttf")) continue;
            Font font = java.awt.Font.createFont(0,ttfFile);
            System.out.println(ttfFile.getName()+":"+font.getName());
        }
    }

    public String replaceFirstLineChar(String str){
        //BaseFont.create

        return str.substring(1);
    }

    @Test
    public void ssss(){
        PrintSetup printSetup = new PrintSetup();
        printSetup.put("sss",true);
        System.out.println(printSetup.get("sss"));
    }
}
