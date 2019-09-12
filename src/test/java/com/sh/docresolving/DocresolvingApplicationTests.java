package com.sh.docresolving;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.sh.docresolving.dto.PrintSetup;
import com.sh.docresolving.utils.Excel2Pdf;
import com.sh.docresolving.utils.ExcelToPdf;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;
import org.springframework.util.StringUtils;

import java.io.File;

@RunWith(SpringRunner.class)
@SpringBootTest
public class DocresolvingApplicationTests {

    @Test
    public void test(){

        System.out.println("$1:$2".matches("\\$\\d:\\$\\d"));
        System.out.println("$1:$B".matches("\\$[A-Z]:\\$[A-Z]"));
    }
}
