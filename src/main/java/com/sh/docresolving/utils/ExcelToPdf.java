package com.sh.docresolving.utils;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.sh.docresolving.entity.Merge;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;

public class ExcelToPdf{


    public static XSSFWorkbook readExcel(String excelPath) throws Exception{
        InputStream is = new FileInputStream(excelPath);
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        return workbook;
    }

    public static void convert(String excelPath,String pdfPath) throws Exception {
        XSSFWorkbook workbook = readExcel(excelPath);
        Rectangle rectangle = new Rectangle(PageSize.A4);
        Document document = new Document(rectangle);
        OutputStream os = new FileOutputStream(pdfPath);
        //PdfWriter pdfWriter = PdfWriter.getInstance(document,os);
        //document.open();
        getPdfTables(workbook,document,os);
        //document.close();
    }

    public static List<PdfPTable> getPdfTables(XSSFWorkbook workbook,Document document,OutputStream os) throws Exception{
        int sheetCount = workbook.getNumberOfSheets();
        List<PdfPTable> tables = new ArrayList<>();
        for(int i = 0 ; i< sheetCount;i++){
            //获取单个sheet
            XSSFSheet sheet = workbook.getSheetAt(i);
            List<PdfPCell> cells = getPdfCells(sheet);
        }
        return tables;
    }

    public static List<PdfPCell> getPdfCells(XSSFSheet sheet) throws Exception{
        int rowCount = sheet.getPhysicalNumberOfRows();
        List<PdfPCell> cells = new ArrayList<>();
        for(int i = 0 ; i < rowCount ; i++){
            XSSFRow row = sheet.getRow(i);
            int cellCount = row.getLastCellNum();
            float[] widths = new float[cellCount];
            for(int j = 0 ; j < cellCount; j++){
                XSSFCell cell = row.getCell(j);
                if(cell == null) cell = row.createCell(j);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                float cellWidth = getPOICellWidth(sheet,cell);
                widths[j] = cellWidth;
                String cellValue = cell.getStringCellValue();
                XSSFFont font = cell.getCellStyle().getFont();
                short height = font.getFontHeightInPoints();
                Merge merge = getColspanRowspanByExcel(sheet,row.getRowNum(),cell.getColumnIndex());
                PdfPCell pdfpCell = new PdfPCell();
                pdfpCell.setBackgroundColor(new BaseColor(POIUtil.getRGB(
                        cell.getCellStyle().getFillForegroundColorColor())));
                pdfpCell.setColspan(merge.getColspan());
                pdfpCell.setRowspan(merge.getRowpan());
                pdfpCell.setVerticalAlignment(getVAlignByExcel(cell.getCellStyle().getVerticalAlignment()));
                pdfpCell.setHorizontalAlignment(getHAlignByExcel(cell.getCellStyle().getAlignment()));
                pdfpCell.setPhrase(getPhrase(sheet.getWorkbook(),cell));
                setPdfCellHeight(row,pdfpCell);
                addBorderByExcel(sheet.getWorkbook(),pdfpCell, cell.getCellStyle());
                addImageByPOICell(pdfpCell , cell , cellWidth);
                cells.add(pdfpCell);
                j += merge.getColspan() - 1;
            }
        }
        return null;
    }

    public static int getPOICellWidth(Sheet sheet,Cell cell) {
        int poiCWidth = sheet.getColumnWidth(cell.getColumnIndex());
        int cellWidthpoi = poiCWidth;
        int widthPixel = 0;
        if (cellWidthpoi >= 416) {
            widthPixel = (int) (((cellWidthpoi - 416.0) / 256.0) * 8.0 + 13.0 + 0.5);
        } else {
            widthPixel = (int) (cellWidthpoi / 416.0 * 13.0 + 0.5);
        }
        return widthPixel;
    }

    public static Merge getColspanRowspanByExcel(XSSFSheet sheet,int rowIndex, int colIndex) {
        CellRangeAddress result = null;
        int num = sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstColumn() == colIndex && range.getFirstRow() == rowIndex) {
                result = range;
            }
        }
        int rowspan = 1;
        int colspan = 1;
        if (result != null) {
            rowspan = result.getLastRow() - result.getFirstRow() + 1;
            colspan = result.getLastColumn() - result.getFirstColumn() + 1;
        }
        Merge merge = new Merge(rowspan,colspan);
        return merge;
    }

    public static int getVAlignByExcel(short align) {
        int result = 0;
        if (align == CellStyle.VERTICAL_BOTTOM) {
            result = Element.ALIGN_BOTTOM;
        }
        if (align == CellStyle.VERTICAL_CENTER) {
            result = Element.ALIGN_MIDDLE;
        }
        if (align == CellStyle.VERTICAL_JUSTIFY) {
            result = Element.ALIGN_JUSTIFIED;
        }
        if (align == CellStyle.VERTICAL_TOP) {
            result = Element.ALIGN_TOP;
        }
        return result;
    }

    public static int getHAlignByExcel(short align) {
        int result = 0;
        if (align == CellStyle.ALIGN_LEFT) {
            result = Element.ALIGN_LEFT;
        }
        if (align == CellStyle.ALIGN_RIGHT) {
            result = Element.ALIGN_RIGHT;
        }
        if (align == CellStyle.ALIGN_JUSTIFY) {
            result = Element.ALIGN_JUSTIFIED;
        }
        if (align == CellStyle.ALIGN_CENTER) {
            result = Element.ALIGN_CENTER;
        }
        return result;
    }

    public static Phrase getPhrase(XSSFWorkbook workbook,Cell cell) {
        return new Phrase(cell.getStringCellValue(), getFontByExcel(workbook,cell.getCellStyle()));
    }

    public static Font getFontByExcel(XSSFWorkbook workbook,CellStyle style) {
        Font result = new Font(Resource.BASE_FONT_CHINESE , 8 , Font.NORMAL);
        short index = style.getFontIndex();
        org.apache.poi.ss.usermodel.Font font = workbook.getFontAt(index);

        if(font.getBoldweight() == org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD){
            result.setStyle(Font.BOLD);
        }

        HSSFColor color = HSSFColor.getIndexHash().get(font.getColor());

        if(color != null){
            int rbg = POIUtil.getRGB(color);
            result.setColor(new BaseColor(rbg));
        }

        FontUnderline underline = FontUnderline.valueOf(font.getUnderline());
        if(underline == FontUnderline.SINGLE){
            String ulString = Font.FontStyle.UNDERLINE.getValue();
            result.setStyle(ulString);
        }
        return result;
    }

    public static void setPdfCellHeight(XSSFRow row,PdfPCell pdfpCell){
        XSSFSheet sheet = row.getSheet();
        if (sheet.getDefaultRowHeightInPoints() != row.getHeightInPoints()) {
            pdfpCell.setFixedHeight(getPixelHeight(row.getHeightInPoints()));
        }
    }

    public static float getPixelHeight(float poiHeight){
        float pixel = poiHeight / 28.6f * 26f;
        return pixel;
    }

    public static void addBorderByExcel(XSSFWorkbook workbook,PdfPCell cell , CellStyle style) {
        cell.setBorderColorLeft(new BaseColor(POIUtil.getBorderRBG(workbook,style.getLeftBorderColor())));
        cell.setBorderColorRight(new BaseColor(POIUtil.getBorderRBG(workbook,style.getRightBorderColor())));
        cell.setBorderColorTop(new BaseColor(POIUtil.getBorderRBG(workbook,style.getTopBorderColor())));
        cell.setBorderColorBottom(new BaseColor(POIUtil.getBorderRBG(workbook,style.getBottomBorderColor())));
    }

    public static void addImageByPOICell(PdfPCell pdfpCell , Cell cell , float cellWidth) throws BadElementException, MalformedURLException, IOException {
        POIImage poiImage = new POIImage().getCellImage(cell);
        byte[] bytes = poiImage.getBytes();
        if(bytes != null){
//           double cw = cellWidth;
//           double ch = pdfpCell.getFixedHeight();
//
//           double iw = poiImage.getDimension().getWidth();
//           double ih = poiImage.getDimension().getHeight();
//
//           double scale = cw / ch;
//
//           double nw = iw * scale;
//           double nh = ih - (iw - nw);
//
//           POIUtil.scale(bytes , nw  , nh);
            pdfpCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            pdfpCell.setHorizontalAlignment(Element.ALIGN_CENTER);
            Image image = Image.getInstance(bytes);
            pdfpCell.setImage(image);
        }
    }
}


