package com.sh.docresolving.utils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.sh.docresolving.dto.PrintSetup;
import org.springframework.util.StringUtils;

public class Excel2Pdf {

    private static final Integer WORD_TO_PDF_OPERAND = 17;
    private static final Integer PPT_TO_PDF_OPERAND = 32;
    private static final Integer EXCEL_TO_PDF_OPERAND = 0;

    public static String excel2Pdf(String inFilePath, String outFilePath,PrintSetup printSetup) throws Exception {
        ActiveXComponent ax = null;
        Dispatch excels = null;
        Dispatch excel = null;
        Dispatch sheets = null;
        Object[] obj2 = null;
        Integer headerStart = printSetup.headerStart();
        try {
            ComThread.InitSTA();
            ax = new ActiveXComponent("Excel.Application");
            ax.setProperty("Visible", new Variant(false));//可视
            ax.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
            excels = ax.getProperty("Workbooks").toDispatch();

            Object[] obj = new Object[]{
                    inFilePath,
                    new Variant(false),
                    new Variant(false)
            };
            excel = Dispatch.invoke(excels, "Open", Dispatch.Method, obj, new int[9]).toDispatch();
            sheets = Dispatch.get((Dispatch) excel, "Sheets")
                    .toDispatch();
            int count = Dispatch.get(sheets, "Count").getInt();
            for (int j = 1; j <=count; j++) {
                Dispatch currentSheet = Dispatch.invoke(sheets, "Item",
                        Dispatch.Get, new Object[] { new Integer(j) },
                        new int[1]).toDispatch();
                String sheetName = Dispatch.get(currentSheet,"Name").toString();
                Dispatch pageSetup = Dispatch.get(currentSheet, "PageSetup")
                        .toDispatch();
                if(j>=headerStart){
                    Object leftHeaderObj = printSetup.get("leftHeader");
                    Object rightHeaderObj = printSetup.get("rightHeader");
                    Object centerHeaderObj = printSetup.get("centerHeaderObj");
                    if(leftHeaderObj!=null&& StringUtils.hasText(leftHeaderObj.toString())){
                        Dispatch.put(pageSetup,"LeftHeader",leftHeaderObj);
                    }
                    if(rightHeaderObj!=null&& StringUtils.hasText(rightHeaderObj.toString())){
                        Dispatch.put(pageSetup,"RightHeader",rightHeaderObj);
                    }
                    if(centerHeaderObj!=null&& StringUtils.hasText(centerHeaderObj.toString())){
                        Dispatch.put(pageSetup,"CenterHeader",centerHeaderObj);
                    }
                }
                //Dispatch.put(pageSetup,"CenterHorizontally",true);
                /*Dispatch.put(pageSetup, "LeftMargin", new Variant(60));
                Dispatch.put(pageSetup, "RightMargin", new Variant(47));
                Dispatch.put(pageSetup, "TopMargin", new Variant(33));
                Dispatch.put(pageSetup, "BottomMargin", new Variant(40));
                Dispatch.put(pageSetup,"PaperSize",new Variant(10));*/
                Variant defalutVariant = printSetup.getJacobVariantByOrientation(sheetName);
                Dispatch.put(pageSetup, "Orientation", defalutVariant);
            }
            // 转换格式
            obj2 = new Object[]{
                    new Variant(EXCEL_TO_PDF_OPERAND), // PDF格式=0
                    outFilePath,
                    new Variant(0)  //0=标准 (生成的PDF图片不会变模糊) ; 1=最小文件
            };
            System.out.println("ExportAsFixedFormat前..................");
            Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method,obj2, new int[1]);
            System.out.println("导出成功..................");
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
            return outFilePath;
        }
    }

    public static String checkFileOutPathAndOut(String originalFileOut){
        originalFileOut=originalFileOut.replaceAll("!","");
        char[] ochars = originalFileOut.toCharArray();
        String firstChar = String.valueOf(ochars[0]);
        if(!firstChar.equals("/")){
            return originalFileOut;
        }else{
            return originalFileOut.substring(1);
        }
    }
}
