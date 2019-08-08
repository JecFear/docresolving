package com.sh.docresolving.utils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.sh.docresolving.entity.PrintSetup;

public class Excel2Pdf {

    private static final Integer WORD_TO_PDF_OPERAND = 17;
    private static final Integer PPT_TO_PDF_OPERAND = 32;
    private static final Integer EXCEL_TO_PDF_OPERAND = 0;

    public static void excel2Pdf(String inFilePath, String outFilePath,PrintSetup printSetup) throws Exception {
        ActiveXComponent ax = null;
        Dispatch excel = null;
        Dispatch sheets = null;
        Object[] obj2 = null;
        try {
            ComThread.InitSTA();
            ax = new ActiveXComponent("Excel.Application");
            ax.setProperty("Visible", new Variant(false));//可视
            ax.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
            Dispatch excels = ax.getProperty("Workbooks").toDispatch();

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
                System.out.println(sheetName);
                Dispatch pageSetup = Dispatch.get(currentSheet, "PageSetup")
                        .toDispatch();
                Dispatch.put(pageSetup, "Orientation", printSetup.getJacobVariantByOrientation(sheetName));
            }
            // 转换格式
            obj2 = new Object[]{
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
