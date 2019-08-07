package com.sh.docresolving.utils;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.sh.docresolving.entity.BorderParam;
import com.sh.docresolving.entity.Merge;
import com.sh.docresolving.entity.MergeBack;
import com.sh.docresolving.entity.PdfPTableEx;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.springframework.security.core.parameters.P;
import org.springframework.util.StringUtils;

import java.io.*;
import java.net.MalformedURLException;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelToPdf{


    public static XSSFWorkbook readExcel(String excelPath) throws Exception{
        InputStream is = new FileInputStream(excelPath);
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        return workbook;
    }

    public static void convert(String excelPath,String pdfPath) throws Exception {
        XSSFWorkbook workbook = readExcel(excelPath);
        Rectangle rectangle = new Rectangle(PageSize.A4);
        float A4Height = PageSize.A4.getHeight();
        Document document = new Document(rectangle);
        OutputStream os = new FileOutputStream(pdfPath);
        PdfWriter pdfWriter = PdfWriter.getInstance(document,os);
        document.open();
        List<PdfPTable> tables = getPdfTables(workbook,document,os,A4Height);
        for(int i = 0;i<tables.size();i++){
            if(i!=0) document.newPage();
            document.add(tables.get(i));
        }
        document.close();
    }

    public static List<PdfPTable> getPdfTables(XSSFWorkbook workbook,Document document,OutputStream os,float A4Height) throws Exception{
        int sheetCount = workbook.getNumberOfSheets();
        List<PdfPTable> tables = new ArrayList<>();
        for(int i = 0 ; i< sheetCount;i++){
            //获取单个sheet
            XSSFSheet sheet = workbook.getSheetAt(i);
            XSSFPrintSetup xssfPrintSetup = sheet.getPrintSetup();
            boolean landScape = xssfPrintSetup.getLandscape();
            PdfPTableEx pdfPTableEx = getPdfCells(sheet,A4Height);
            PdfPTable table = new PdfPTable(pdfPTableEx.getWidths());
            table.setWidthPercentage(100);
            for (PdfPCell pdfpCell : pdfPTableEx.getCells()) {
                table.addCell(pdfpCell);
            }
            tables.add(table);
        }
        return tables;
    }

    public static PdfPTableEx getPdfCells(XSSFSheet sheet,float A4Height) throws Exception{
        List<PdfPCell> cells = new ArrayList<>();
        PdfPTableEx pdfPTableEx = new PdfPTableEx();
        float[] widths = null;
        float mw = 0;
        BorderParam borderParam = getBorderParam(sheet);
        int maxColNum = borderParam.getMaxCol();
        int maxRowNum = borderParam.getMaxRow();
        Map<Integer,MergeBack> mergeBacks = new HashMap<>();
        for(int i = 0 ; i < maxRowNum ; i++){
            XSSFRow row = sheet.getRow(i);
            float[] cws = new float[maxColNum];
            MergeBack mergeBack = mergeBacks.get(i);
            for(int j = 0 ; j < maxColNum; j++){
                XSSFCell cell = row.getCell(j);
                if(mergeBack!=null&&j<=mergeBack.getEnd()&&j>=mergeBack.getStart()) continue;
                if(cell == null) cell = row.createCell(j);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                float cellWidth = sheet.getColumnWidth(cell.getColumnIndex());
                cws[cell.getColumnIndex()] = cellWidth;
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
                pdfpCell.setPhrase(getPhrase(sheet.getWorkbook(),cell,getZoomInFontHeight(height)));
                setPdfCellHeight(row,pdfpCell);
                addBorderByExcel(sheet.getWorkbook(),pdfpCell, cell,cell.getCellStyle());
                addImageByPOICell(pdfpCell , cell , cellWidth);
                cells.add(pdfpCell);
                if(merge.getRowpan()>1){
                    MergeBack newMergeBack = new MergeBack(j,j+merge.getColspan()-1);
                    for(int k=1;k<merge.getRowpan();k++){
                        mergeBacks.put(i+k,newMergeBack);
                    }
                }
                j += merge.getColspan() - 1;
            }
            float rw = 0;
            for (int j = 0; j < cws.length; j++) {
                rw += cws[j];
            }
            if (rw > mw ||  mw == 0) {
                widths = cws;
                mw = rw;
            }
        }
        pdfPTableEx.setWidths(widths);
        pdfPTableEx.setCells(cells);
        return pdfPTableEx;
    }

    public static float getRowPixelHeight(float pageHeight,float maxHegiht,float rowHegiht){
        float rowPixelHeight = pageHeight * (rowHegiht/maxHegiht);
        Integer intValue = (int)rowPixelHeight;
        return intValue;
    }

    public static int getMaxHeight(XSSFSheet sheet,int maxRowNum){
        int maxHegiht = 0;
        for(int i = 0;i<maxRowNum;i++){
            XSSFRow row = sheet.getRow(i);
            maxHegiht+=row.getHeightInPoints();
        }
        return maxHegiht;
    }

    public static BorderParam getBorderParam(XSSFSheet sheet){
        int rowCount = sheet.getLastRowNum();
        int maxCellNum = 0;
        int lastTextRowNum = 0;
        for(int i = 0 ; i < rowCount ; i++){
            XSSFRow row = sheet.getRow(i);
            int cellCount = row.getLastCellNum();
            if(cellCount>maxCellNum) maxCellNum = cellCount;
            for(int j = 0 ; j< cellCount ; j++){
                XSSFCell cell = row.getCell(j);
                if(cell == null) cell = row.createCell(j);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                if(StringUtils.hasText(cell.getStringCellValue())&&i>=lastTextRowNum){
                    lastTextRowNum = i+1;
                }
            }
        }
        return(new BorderParam(maxCellNum,lastTextRowNum,0));
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

    public static Phrase getPhrase(XSSFWorkbook workbook,Cell cell,float fontSize) {
        return new Phrase(cell.getStringCellValue(), getFontByExcel(workbook,cell.getCellStyle(),fontSize));
    }

    public static Font getFontByExcel(XSSFWorkbook workbook,CellStyle style,float fontSize) {
        Font result = new Font(Resource.BASE_FONT_CHINESE , Font.NORMAL);
        result.setSize(fontSize);
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

    public static float setPdfCellHeight(XSSFRow row,PdfPCell pdfpCell){
        XSSFSheet sheet = row.getSheet();
        float rowHegiht = 0;
        if (sheet.getDefaultRowHeightInPoints() != row.getHeightInPoints()) {
            rowHegiht = row.getHeightInPoints();
            pdfpCell.setFixedHeight(rowHegiht);
        }
        return rowHegiht;
    }

    public static float getZoomInFontHeight(float originalHeight){
        float zoomHeight = originalHeight * 998 / 1000;
        return zoomHeight;
    }

    public static float getPixelHeight(float poiHeight){
        float pixel = poiHeight / 72 * 96;
        return pixel;
    }

    public static void addBorderByExcel(XSSFWorkbook workbook,PdfPCell pdfCell ,XSSFCell cell, CellStyle style) {
        short borderTop = style.getBorderTop();
        short borderLeft = style.getBorderLeft();
        short borderBottom = style.getBorderBottom();
        short borderRight = style.getBorderRight();
        if(borderTop>0) {
            pdfCell.setBorderColorTop(new BaseColor(POIUtil.getBorderRBG(workbook, style.getTopBorderColor())));
        }else {
            pdfCell.disableBorderSide(1);
        }
        if(borderLeft>0) {
            pdfCell.setBorderColorLeft(new BaseColor(POIUtil.getBorderRBG(workbook,style.getLeftBorderColor())));
        }else{
            pdfCell.disableBorderSide(4);
        }
        if(borderBottom>0) {
            pdfCell.setBorderColorBottom(new BaseColor(POIUtil.getBorderRBG(workbook,style.getBottomBorderColor())));
        }else {
            pdfCell.disableBorderSide(2);
        }
        if(borderRight>0) {
            pdfCell.setBorderColorRight(new BaseColor(POIUtil.getBorderRBG(workbook,style.getRightBorderColor())));
        }else{
            pdfCell.disableBorderSide(8);
        }
    }

    public static void addImageByPOICell(PdfPCell pdfpCell , Cell cell , float cellWidth) throws BadElementException, MalformedURLException, IOException {
        POIImage poiImage = new POIImage().getCellImage(cell);
        byte[] bytes = poiImage.getBytes();
        if(bytes != null){
            pdfpCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            pdfpCell.setHorizontalAlignment(Element.ALIGN_CENTER);
            Image image = Image.getInstance(bytes);
            pdfpCell.setImage(image);
        }
    }

    private static final Integer WORD_TO_PDF_OPERAND = 17;
    private static final Integer PPT_TO_PDF_OPERAND = 32;
    private static final Integer EXCEL_TO_PDF_OPERAND = 0;

    public static void excel2Pdf(String inFilePath, String outFilePath) throws Exception {
        ActiveXComponent ax = null;
        Dispatch excel = null;
        try {
            ComThread.InitSTA();
            ax = new ActiveXComponent("Excel.Application");
            ax.setProperty("Visible", new Variant(false));
            ax.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
            Dispatch excels = ax.getProperty("Workbooks").toDispatch();

            Object[] obj = new Object[]{
                    inFilePath,
                    new Variant(false),
                    new Variant(false)
            };
            excel = Dispatch.invoke(excels, "Open", Dispatch.Method, obj, new int[9]).toDispatch();

            // 转换格式
            Object[] obj2 = new Object[]{
                    new Variant(EXCEL_TO_PDF_OPERAND), // PDF格式=0
                    outFilePath,
                    new Variant(0)  //0=标准 (生成的PDF图片不会变模糊) ; 1=最小文件
            };
            Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method,obj2, new int[1]);

        } catch (Exception es) {
            es.printStackTrace();
            throw es;
        } finally {
            if (excel != null) {
                Dispatch.call(excel, "Close", new Variant(false));
            }
            if (ax != null) {
                ax.invoke("Quit", new Variant[] {});
                ax = null;
            }
            ComThread.Release();
        }

    }
}


