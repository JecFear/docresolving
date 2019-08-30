package com.sh.docresolving;

import com.sh.docresolving.utils.ExcelToPdf;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;

@RunWith(SpringRunner.class)
@SpringBootTest
public class DocresolvingApplicationTests {

//    @Test
//    public void contextLoads() {
//    }
//
//    @Test
//    public void HSSFWORKBOOKTEST() throws Exception{
//        String fileIn = "sample1/download.xlsx";
//        String uri = this.getClass().getResource(fileIn).getPath();
//        String fileOut = uri.replaceAll(".xls$|.xlsx$",".pdf");
//        ExcelToPdf.convert(uri,fileOut);
//    }
//
//    @Test
//    public void HSSFWORKBOOKTESTss() throws Exception{
//        String fileIn = "F:\\docresolving\\target\\test-classes\\com\\sh\\docresolving\\sample1\\111.xlsx";
//        /*String uri = this.getClass().getResource(fileIn).getPath();
//        System.out.println(fileIn);*/
//        String fileOut = fileIn.replaceAll(".xls$|.xlsx$",".pdf");
//        ExcelToPdf.excel2Pdf(fileIn,fileOut);
//    }
}
