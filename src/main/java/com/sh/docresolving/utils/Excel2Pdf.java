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

    /**
     * excel转pdf
     * @param inFilePath 要转换的excel,可以是fastdfs路径，可以是本地文件路径
     * @param outFilePath 本地输出文件路径
     * @param printSetup  打印设置，详见PrintSetup对象
     * @return
     * @throws Exception
     */
    public static String excel2Pdf(String inFilePath, String outFilePath, PrintSetup printSetup) throws Exception {
        ActiveXComponent ax = null;
        Dispatch excels = null;
        Dispatch excel = null;
        Dispatch sheets = null;
        Object[] obj2 = null;
        Integer headerStart = printSetup.headerStart();
        try {

            ComThread.InitSTA();//初始化COM资源
            ax = new ActiveXComponent("Excel.Application"); //初始化EXCEL程序
            ax.setProperty("Visible", new Variant(false));//可视
            ax.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
            excels = ax.getProperty("Workbooks").toDispatch();//获取Workbooks属性以使用

            Object[] obj = new Object[]{
                    inFilePath,
                    new Variant(false),
                    new Variant(false)
            };//构建invoke方法参数
            excel = Dispatch.invoke(excels, "Open", Dispatch.Method, obj, new int[9]).toDispatch();//打开这个excel
            sheets = Dispatch.get(excel, "Sheets")
                    .toDispatch();//获取excel的sheets
            int pageCount = beforeSheetPageNum(sheets,printSetup.pageNumStart());//获取客户设定的起始sheet前会生成多少页pdf
            int count = Dispatch.get(sheets, "Count").getInt();//获取sheets的数量
            for (int j = 1; j <=count; j++) {
                Dispatch currentSheet = Dispatch.invoke(sheets, "Item",
                        Dispatch.Get, new Object[] { new Integer(j) },
                        new int[1]).toDispatch();//获取每个sheet
                String sheetName = Dispatch.get(currentSheet,"Name").toString();
                Dispatch pageSetup = Dispatch.get(currentSheet, "PageSetup")
                        .toDispatch();//获取pageSetup对象，以设置打印参数
                Dispatch.put(pageSetup,"AlignMarginsHeaderFooter",new Variant(true));
                Dispatch.put(pageSetup,"ScaleWithDocHeaderFooter",new Variant(true));
                if(j>=headerStart){
                    Object leftHeaderObj = printSetup.get("leftHeader");
                    Object rightHeaderObj = printSetup.get("rightHeader");
                    Object centerHeaderObj = printSetup.get("centerHeader");
                    if(leftHeaderObj!=null&& StringUtils.hasText(leftHeaderObj.toString())) Dispatch.put(pageSetup,"LeftHeader",leftHeaderObj);
                    if(rightHeaderObj!=null&& StringUtils.hasText(rightHeaderObj.toString())) Dispatch.put(pageSetup,"RightHeader",rightHeaderObj);
                    if(centerHeaderObj!=null&& StringUtils.hasText(centerHeaderObj.toString())) Dispatch.put(pageSetup,"CenterHeader",centerHeaderObj);
                }
                if(printSetup.needPageNum()&&j>=printSetup.pageNumStart()){
                    Dispatch.put(pageSetup,"CenterFooter","第 &P-"+pageCount+ " 页,共 &N-"+pageCount+" 页");
                }
                /*Dispatch.put(pageSetup,"LeftFooter","&G");
                Dispatch leftFooterPicture = Dispatch.get(pageSetup,"LeftFooterPicture").toDispatch();
                Dispatch.put(leftFooterPicture,"FileName","http://www.shouhouzn.net/group1/M00/00/1E/rBGmcV1nhEaACsA5AADJuei4SZI817.png");
                Dispatch.put(leftFooterPicture,"Width",100);*/

                Dispatch.put(pageSetup,"CenterHorizontally",printSetup.getCenterHorizontally());
                Dispatch.put(pageSetup,"CenterVertically",printSetup.getCenterVertically());
                Dispatch.put(pageSetup,"PaperSize",new Variant(9));
                Variant defalutVariant = printSetup.getJacobVariantByOrientation(sheetName);
                Dispatch.put(pageSetup, "Orientation", defalutVariant);
                if(printSetup.getSheetOrientation(sheetName)){
                    if(printSetup.getDouble("leftMargin")!=null) Dispatch.put(pageSetup, "LeftMargin", new Variant(printSetup.getDouble("leftMargin")*28.35));
                    if(printSetup.getDouble("rightMargin")!=null) Dispatch.put(pageSetup, "RightMargin", new Variant(printSetup.getDouble("rightMargin")*28.35));
                    if(printSetup.getDouble("topMargin")!=null) Dispatch.put(pageSetup, "TopMargin", new Variant(printSetup.getDouble("topMargin")*28.35));
                    if(printSetup.getDouble("bottomMargin")!=null) Dispatch.put(pageSetup, "BottomMargin", new Variant(printSetup.getDouble("bottomMargin")*28.35));
                }else{
                    if(printSetup.getDouble("leftMargin")!=null) Dispatch.put(pageSetup, "BottomMargin", new Variant(printSetup.getDouble("leftMargin")*28.35));
                    if(printSetup.getDouble("rightMargin")!=null) Dispatch.put(pageSetup, "TopMargin", new Variant(printSetup.getDouble("rightMargin")*28.35));
                    if(printSetup.getDouble("topMargin")!=null) Dispatch.put(pageSetup, "LeftMargin", new Variant(printSetup.getDouble("topMargin")*28.35));
                    if(printSetup.getDouble("bottomMargin")!=null) Dispatch.put(pageSetup, "RightMargin", new Variant(printSetup.getDouble("bottomMargin")*28.35));
                }
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

    /**
     * 获取页码起始sheet前的所有sheet转换pdf时的总页数，以方便&P减去，获取真实页码值
     * @param sheets
     * @param pageStartSheet
     * @return
     */
    public static int beforeSheetPageNum(Dispatch sheets,int pageStartSheet){
        int pageCount = 0;
        for (int j = 1; j <pageStartSheet; j++) {
            Dispatch currentSheet = Dispatch.invoke(sheets, "Item",
                    Dispatch.Get, new Object[]{new Integer(j)},
                    new int[1]).toDispatch();
            Dispatch pageSetup = Dispatch.get(currentSheet, "PageSetup")
                    .toDispatch();
            Dispatch pages = Dispatch.get(pageSetup,"Pages").toDispatch();
            int sheetPageCount = Dispatch.get(pages,"Count").getInt();
            pageCount+=sheetPageCount;
        }
        return pageCount;
    }
}
