package com.sh.docresolving.service;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

/**
 * @Author:Dawn
 * @Date: 2019/10/22 9:44
 **/
@Service
public class PdfResolvingService {

    @Autowired
    private FastDFSService fastDFSService;

    public String pdfMerge(List<String> localPaths) throws Exception{
        String outFileName = System.currentTimeMillis()+"merged"+".pdf";
        String fileOut = "C:\\pdf"+ File.separator+outFileName;
        File file = new File(fileOut);
        PdfReader pdfReader = new PdfReader(localPaths.get(0));
        Document document = new Document(pdfReader.getPageSize(1));
        PdfCopy copy = new PdfCopy(document, new FileOutputStream(fileOut));
        document.open();
        for (int i = 0; i < localPaths.size(); i++) {
            PdfReader reader = new PdfReader(localPaths.get(i));
            int n = reader.getNumberOfPages();
            for (int j = 1; j <= n; j++) {
                document.newPage();
                PdfImportedPage page = copy.getImportedPage(reader, j);
                copy.addPage(page);
            }
            reader.close();
        }
        pdfReader.close();
        document.close();
        String fastOutUrl = fastDFSService.uploadFile(file);
        return fastOutUrl;
    }

}
