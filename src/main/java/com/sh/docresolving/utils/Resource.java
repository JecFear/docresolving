package com.sh.docresolving.utils;

import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.pdf.BaseFont;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.io.File;

public class Resource {

    protected static BaseFont BASE_FONT_CHINESE;
    static {
        try {
            BASE_FONT_CHINESE = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
            FontFactory.registerDirectories();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static Font getFont(XSSFFont font) {
        try {
            Font iTextFont = FontFactory.getFont(font.getFontName(),
                    BaseFont.IDENTITY_H, BaseFont.EMBEDDED,
                    font.getFontHeightInPoints());
            return iTextFont;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

     public static Font getFontByName(XSSFFont font){
        String fontName = font.getFontName();
        String ttfDir =  Thread.currentThread().getContextClassLoader().getResource("ttf").getPath();
        File file = new File(ttfDir+fontName+".ttf");
        boolean a = file.exists();
        return null;
     }
}