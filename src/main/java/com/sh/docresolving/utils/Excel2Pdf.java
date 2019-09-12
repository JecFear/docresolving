package com.sh.docresolving.utils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.sh.docresolving.dto.PrintSetup;
import org.aspectj.weaver.ast.Var;
import org.springframework.util.StringUtils;

import java.util.Map;

public class Excel2Pdf {

    private static final Integer WORD_TO_PDF_OPERAND = 17;
    private static final Integer PPT_TO_PDF_OPERAND = 32;
    private static final Integer EXCEL_TO_PDF_OPERAND = 0;

    private static final double A4_Height = 841.95;
    private static final double A4_Width = 595.35;

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
            int pageHeaderCount = 0;
            int pageFooterCount = 0;
            if(printSetup.headerOrFooterStart(printSetup.getHeader())!=0){
                pageHeaderCount = beforeSheetPageNum(sheets,printSetup.headerOrFooterStart(printSetup.getHeader()));
            }
            if(printSetup.headerOrFooterStart(printSetup.getFooter())!=0){
                pageFooterCount = beforeSheetPageNum(sheets,printSetup.headerOrFooterStart(printSetup.getFooter()));
            }
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

                /*Map<String,String> sheetValue = printSetup.getSheetsValues(sheetName);
                if(sheetValue.get("topTitleRows")!=null&&sheetValue.get("topTitleRows").matches("\\$\\d:\\$\\d")){
                    Dispatch.put(pageSetup,"PrintTitleRows",sheetValue.get("topTitleRows"));
                }
                if(sheetValue.get("bottomTitleRows")!=null&&sheetValue.get("bottomTitleRows").matches("\\$\\d:\\$\\d")){

                }
                if(sheetValue.get("leftTitleRows")!=null&&sheetValue.get("leftTitleRows").matches("\\$[A-Z]:\\$[A-Z]")){
                    Dispatch.put(pageSetup,"PrintTitleColumns",sheetValue.get("leftTitleRows"));
                }
                if(printSetup.getSheetOrientation(sheetName)){
                    Map<String,Double> landScapeSet = printSetup.getLandscapeSet();
                    if(landScapeSet.get("leftMargin")!=null&&landScapeSet.get("leftMargin")!=0) Dispatch.put(pageSetup, "LeftMargin", new Variant(landScapeSet.get("leftMargin")*28.35));
                    if(landScapeSet.get("rightMargin")!=null&&landScapeSet.get("rightMargin")!=0) Dispatch.put(pageSetup, "RightMargin", new Variant(landScapeSet.get("rightMargin")*28.35));
                    if(landScapeSet.get("topMargin")!=null&&landScapeSet.get("topMargin")!=0) Dispatch.put(pageSetup, "TopMargin", new Variant(landScapeSet.get("topMargin")*28.35));
                    if(landScapeSet.get("bottomMargin")!=null&&landScapeSet.get("bottomMargin")!=0) Dispatch.put(pageSetup, "BottomMargin", new Variant(landScapeSet.get("bottomMargin")*28.35));
                    if(landScapeSet.get("headerMargin")!=null&&landScapeSet.get("headerMargin")!=0) Dispatch.put(pageSetup,"HeaderMargin",new Variant(landScapeSet.get("headerMargin")*28.35));
                    if(landScapeSet.get("footerMargin")!=null&&landScapeSet.get("footerMargin")!=0) Dispatch.put(pageSetup,"FooterMargin",new Variant(landScapeSet.get("footerMargin")*28.35));
                }else{
                    Map<String,Double> portraitSet = printSetup.getPortraitSet();
                    if(portraitSet.get("leftMargin")!=null&&portraitSet.get("leftMargin")!=0) Dispatch.put(pageSetup, "LeftMargin", new Variant(portraitSet.get("leftMargin")*28.35));
                    if(portraitSet.get("rightMargin")!=null&&portraitSet.get("rightMargin")!=0) Dispatch.put(pageSetup, "RightMargin", new Variant(portraitSet.get("rightMargin")*28.35));
                    if(portraitSet.get("topMargin")!=null&&portraitSet.get("topMargin")!=0) Dispatch.put(pageSetup, "TopMargin", new Variant(portraitSet.get("topMargin")*28.35));
                    if(portraitSet.get("bottomMargin")!=null&&portraitSet.get("bottomMargin")!=0) Dispatch.put(pageSetup, "BottomMargin", new Variant(portraitSet.get("bottomMargin")*28.35));
                    if(portraitSet.get("headerMargin")!=null&&portraitSet.get("headerMargin")!=0) Dispatch.put(pageSetup,"HeaderMargin",new Variant(portraitSet.get("headerMargin")*28.35));
                    if(portraitSet.get("footerMargin")!=null&&portraitSet.get("footerMargin")!=0) Dispatch.put(pageSetup,"FooterMargin",new Variant(portraitSet.get("footerMargin")*28.35));
                }

                if(pageHeaderCount!=0&&j>=printSetup.headerOrFooterStart(printSetup.getHeader())){
                    String leftHeaderStr = printSetup.getExactHeaderOrFooter(printSetup.getHeader(),"leftHeader",pageHeaderCount);
                    String centerHeaderStr = printSetup.getExactHeaderOrFooter(printSetup.getHeader(),"centerHeader",pageHeaderCount);
                    String rightHeaderStr = printSetup.getExactHeaderOrFooter(printSetup.getHeader(),"rightHeader",pageHeaderCount);
                    Map<String,Object> leftHeaderPictureObject =  printSetup.getHeaderOrFooterPictureObject(printSetup.getHeader(),"leftHeaderPicture");
                    Map<String,Object> centerHeaderPictureObject = printSetup.getHeaderOrFooterPictureObject(printSetup.getHeader(),"centerHeaderPicture");
                    Map<String,Object> rightHeaderPictureObject = printSetup.getHeaderOrFooterPictureObject(printSetup.getHeader(),"rightHeaderPicture");
                    if(StringUtils.hasText(leftHeaderStr)) Dispatch.put(pageSetup,"LeftHeader",leftHeaderStr);
                    if(StringUtils.hasText(centerHeaderStr)) Dispatch.put(pageSetup,"CenterHeader",centerHeaderStr);
                    if(StringUtils.hasText(rightHeaderStr)) Dispatch.put(pageSetup,"RightHeader",rightHeaderStr);
                    if(printSetup.doesContainsGraph(leftHeaderStr)&&StringUtils.hasText(printSetup.getHeaderOrFooterPicture(leftHeaderPictureObject))){
                        Dispatch leftHeaderPicture = Dispatch.get(pageSetup,"LeftHeaderPicture").toDispatch();
                        Dispatch.put(leftHeaderPicture,"FileName",printSetup.getHeaderOrFooterPicture(leftHeaderPictureObject));
                        Dispatch.put(leftHeaderPicture,"Width",printSetup.getPictureParam(leftHeaderPictureObject,"width")*28.35*0.809);
                        Dispatch.put(leftHeaderPicture,"Height",printSetup.getPictureParam(leftHeaderPictureObject,"height")*28.35*0.809);
                    }
                    if(printSetup.doesContainsGraph(centerHeaderStr)&&StringUtils.hasText(printSetup.getHeaderOrFooterPicture(centerHeaderPictureObject))){
                        Dispatch centerHeaderPicture = Dispatch.get(pageSetup,"CenterHeaderPicture").toDispatch();
                        Dispatch.put(centerHeaderPicture,"FileName",printSetup.getHeaderOrFooterPicture(centerHeaderPictureObject));
                        Dispatch.put(centerHeaderPicture,"Width",printSetup.getPictureParam(centerHeaderPictureObject,"width")*28.35*0.809);
                        Dispatch.put(centerHeaderPicture,"Height",printSetup.getPictureParam(centerHeaderPictureObject,"height")*28.35*0.809);
                    }
                    if(printSetup.doesContainsGraph(rightHeaderStr)&&StringUtils.hasText(printSetup.getHeaderOrFooterPicture(rightHeaderPictureObject))){
                        Dispatch rightHeaderPicture = Dispatch.get(pageSetup,"RightHeaderPicture").toDispatch();
                        Dispatch.put(rightHeaderPicture,"FileName",printSetup.getHeaderOrFooterPicture(rightHeaderPictureObject));
                        Dispatch.put(rightHeaderPicture,"Width",printSetup.getPictureParam(rightHeaderPictureObject,"width")*28.35*0.809);
                        Dispatch.put(rightHeaderPicture,"Height",printSetup.getPictureParam(rightHeaderPictureObject,"height")*28.35*0.809);
                    }
                }
                if(pageFooterCount!=0&&j>=printSetup.headerOrFooterStart(printSetup.getFooter())){
                    String leftFooterStr = printSetup.getExactHeaderOrFooter(printSetup.getFooter(),"leftFooter",pageFooterCount);
                    String centerFooterStr = printSetup.getExactHeaderOrFooter(printSetup.getFooter(),"centerFooter",pageFooterCount);
                    String righFooterStr = printSetup.getExactHeaderOrFooter(printSetup.getFooter(),"rightFooter",pageFooterCount);
                    Map<String,Object> leftFooterPictureObject = printSetup.getHeaderOrFooterPictureObject(printSetup.getFooter(),"leftFooterPicture");
                    Map<String,Object> centerFooterPictureObject = printSetup.getHeaderOrFooterPictureObject(printSetup.getFooter(),"centerFooterPicture");
                    Map<String,Object> rightFooterPictureObject = printSetup.getHeaderOrFooterPictureObject(printSetup.getFooter(),"rightFooterPicture");
                    if(StringUtils.hasText(leftFooterStr)) Dispatch.put(pageSetup,"LeftFooter",leftFooterStr);
                    if(StringUtils.hasText(centerFooterStr)) Dispatch.put(pageSetup,"CenterFooter",centerFooterStr);
                    if(StringUtils.hasText(righFooterStr)) Dispatch.put(pageSetup,"RightFooter",righFooterStr);
                    if(printSetup.doesContainsGraph(leftFooterStr)&&StringUtils.hasText(printSetup.getHeaderOrFooterPicture(leftFooterPictureObject))){
                        Dispatch leftFooterPicture = Dispatch.get(pageSetup,"LeftFooterPicture").toDispatch();
                        Dispatch.put(leftFooterPicture,"FileName",printSetup.getHeaderOrFooterPicture(leftFooterPictureObject));
                        Dispatch.put(leftFooterPicture,"Width",printSetup.getPictureParam(leftFooterPictureObject,"width")*28.35*0.809);
                        Dispatch.put(leftFooterPicture,"Height",printSetup.getPictureParam(leftFooterPictureObject,"height")*28.35*0.809);
                    }
                    if(printSetup.doesContainsGraph(centerFooterStr)&&StringUtils.hasText(printSetup.getHeaderOrFooterPicture(centerFooterPictureObject))){
                        Dispatch centerFooterPicture = Dispatch.get(pageSetup,"CenterFooterPicture").toDispatch();
                        Dispatch.put(centerFooterPicture,"FileName",printSetup.getHeaderOrFooterPicture(centerFooterPictureObject));
                        Dispatch.put(centerFooterPicture,"Width",printSetup.getPictureParam(centerFooterPictureObject,"width")*28.35*0.809);
                        Dispatch.put(centerFooterPicture,"Height",printSetup.getPictureParam(centerFooterPictureObject,"height")*28.35*0.809);
                    }
                    if(printSetup.doesContainsGraph(righFooterStr)&&StringUtils.hasText(printSetup.getHeaderOrFooterPicture(rightFooterPictureObject))){
                        Dispatch rightFooterPicture = Dispatch.get(pageSetup,"RightFooterPicture").toDispatch();
                        Dispatch.put(rightFooterPicture,"FileName",printSetup.getHeaderOrFooterPicture(rightFooterPictureObject));
                        Dispatch.put(rightFooterPicture,"Width",printSetup.getPictureParam(rightFooterPictureObject,"width")*28.35*0.809);
                        Dispatch.put(rightFooterPicture,"Height",printSetup.getPictureParam(rightFooterPictureObject,"height")*28.35*0.809);
                    }
                }*/
                if(j==6){
                    String topTitleRowsStr = "$1:$1";
                    String bottomTitleRowsStr = "$44:$44";
                    double topMargin = Dispatch.get(pageSetup,"TopMargin").getDouble();
                    double bottomMargin = +Dispatch.get(pageSetup,"BottomMargin").getDouble();

                    Dispatch.invoke(sheets,"Add",Dispatch.Method,new Object[]{
                            null,Dispatch.invoke(sheets, "Item",
                            Dispatch.Get, new Object[] { new Integer(count) },
                            new int[1]).toDispatch(),new Variant(1),new Variant(-4167)
                    },new int[1]);

                    Dispatch topTitleRange = Dispatch.invoke(currentSheet, "Range",
                            Dispatch.Get, new Object[] { topTitleRowsStr },
                            new int[1]).toDispatch();
                    double topTitleHeight = Dispatch.get(topTitleRange,"Height").getDouble();
                    Dispatch topTitleRows = Dispatch.get(topTitleRange,"Rows").toDispatch();
                    int topTitleRowsCount = Dispatch.get(topTitleRows,"Count").getInt();

                    Dispatch bottomTitleRange = Dispatch.invoke(currentSheet, "Range",
                            Dispatch.Get, new Object[] { bottomTitleRowsStr },
                            new int[1]).toDispatch();
                    double bottomTitleHeight = Dispatch.get(bottomTitleRange,"Height").getDouble();
                    Dispatch bottomTitleRows = Dispatch.get(bottomTitleRange,"Rows").toDispatch();
                    int bottomTitleRowsCount = Dispatch.get(bottomTitleRows,"Count").getInt();

                    Dispatch.invoke(bottomTitleRange,"Copy",Dispatch.Method,new Object[]{
                    },new int[1]);

                    Dispatch newlySheet =Dispatch.invoke(sheets, "Item",
                            Dispatch.Get, new Object[] { new Integer(count+1) },
                            new int[1]).toDispatch();
                    Dispatch.invoke(Dispatch.invoke(newlySheet, "Range",Dispatch.Get, new Object[] { "A1" },new int[1]).toDispatch(),"PasteSpecial",Dispatch.Method,new Object[]{
                            new Variant(8)
                    },new int[1]);
                    Dispatch.invoke(Dispatch.invoke(newlySheet, "Range",Dispatch.Get, new Object[] { "A1" },new int[1]).toDispatch(),"PasteSpecial",Dispatch.Method,new Object[]{
                            new Variant(-4104)
                    },new int[1]);
                    Dispatch.call(bottomTitleRange,"Delete");

                    Dispatch usedRange = Dispatch.get(currentSheet,"UsedRange").toDispatch();
                    Dispatch rows = Dispatch.get(usedRange,"Rows").toDispatch();
                    int rowCount = Dispatch.get(rows,"Count").getInt();
                    double marginHeight = topTitleHeight+bottomTitleHeight;
                    double recordHeight = 0;


                    for(int r=1;r<=rowCount;r++){
                        if(r==1||r==44) continue;
                        Dispatch row =  Dispatch.invoke(rows, "Item",
                                Dispatch.Get, new Object[] { new Integer(r) },
                                new int[1]).toDispatch();
                        Dispatch rowNext =  Dispatch.invoke(rows, "Item",
                                Dispatch.Get, new Object[] { new Integer(r+1) },
                                new int[1]).toDispatch();
                        double rowHeight = Dispatch.get(row,"RowHeight").getDouble();
                        double rowNextHeight = Dispatch.get(rowNext,"RowHeight").getDouble();
                        recordHeight+=rowHeight;

                        if(recordHeight+marginHeight+rowNextHeight-A4_Width>=0){
                            Dispatch currentRow = Dispatch.invoke(currentSheet, "Range",Dispatch.Get, new Object[] { "$"+r+":$"+r },new int[1]).toDispatch();
                            Dispatch.invoke(currentRow,"Insert",Dispatch.Method,new Object[]{
                                    new Variant(-4121)
                            },new int[1]);

                            Dispatch.invoke(Dispatch.invoke(newlySheet, "Range",Dispatch.Get, new Object[] { "$"+1+":$"+bottomTitleRowsCount },new int[1]).toDispatch(),"CopyPicture",Dispatch.Method,new Object[]{
                                    new Variant(2),new Variant(1)
                            },new int[1]);

                            Dispatch.invoke(currentSheet,"Paste",Dispatch.Method,new Object[]{
                                    Dispatch.invoke(currentSheet, "Range",Dispatch.Get, new Object[] { "$"+r+":$"+r },new int[1]).toDispatch()
                            },new int[1]);
                            recordHeight=0;
                        }
                    }
                }


                Dispatch.put(pageSetup,"CenterHorizontally",printSetup.getCenterHorizontally());
                Dispatch.put(pageSetup,"CenterVertically",printSetup.getCenterVertically());
                Dispatch.put(pageSetup,"PaperSize",new Variant(9));
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
            Dispatch.invoke(excel, "Save",Dispatch.Method, new Object[]{}, new int[1]);
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
