package com.sh.docresolving;

import com.itextpdf.text.DocumentException;
import com.sh.docresolving.utils.Excel2Pdf;
import com.sh.docresolving.utils.ExcelObject;
import com.sh.docresolving.utils.ExcelToPdf;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;

@RunWith(SpringRunner.class)
@SpringBootTest
public class DocresolvingApplicationTests {

    @Test
    public void contextLoads() {
    }

    @Test
    public void HSSFWORKBOOKTEST() throws Exception{
        String fileIn = "sample1/case4.xlsx";
        String uri = this.getClass().getResource(fileIn).getPath();
        String fileOut = uri.replaceAll(".xls$|.xlsx$",".pdf");
        ExcelToPdf.convert(uri,fileOut);
    }


    @Test
    public void testCase5() throws IOException, DocumentException {
        String fileIn = "sample1/case4.xlsx";

        InputStream in = this.getClass().getResourceAsStream(fileIn);
        Excel2Pdf excel2Pdf = new Excel2Pdf(Arrays.asList(
                new ExcelObject(in)
        ), new FileOutputStream(fileOut(fileIn)));
        excel2Pdf.convert();
    }

    private File fileOut(String fileIn) {
        String uri = this.getClass().getResource(fileIn).getPath();
        String fileOut = uri.replaceAll(".xls$|.xlsx$",".pdf");
        File file = new File(fileOut);
        return file;
    }
}
