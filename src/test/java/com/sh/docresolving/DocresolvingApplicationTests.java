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

    @Test
    public void contextLoads() {
    }

    @Test
    public void HSSFWORKBOOKTEST() throws Exception{
        String fileIn = "sample1/tab.xlsx";
        String uri = this.getClass().getResource(fileIn).getPath();
        String fileOut = uri.replaceAll(".xls$|.xlsx$",".pdf");
        ExcelToPdf.convert(uri,fileOut);
    }

    private File fileOut(String fileIn) {
        String uri = this.getClass().getResource(fileIn).getPath();
        String fileOut = uri.replaceAll(".xls$|.xlsx$",".pdf");
        File file = new File(fileOut);
        return file;
    }
}
