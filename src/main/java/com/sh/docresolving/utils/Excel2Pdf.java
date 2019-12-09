package com.sh.docresolving.utils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.DispatchProxy;
import com.jacob.com.Variant;
import com.sh.docresolving.dto.PrintSetup;
import org.aspectj.weaver.ast.Var;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Excel2Pdf {

    private static final Integer WORD_TO_PDF_OPERAND = 17;
    private static final Integer PPT_TO_PDF_OPERAND = 32;
    private static final Integer EXCEL_TO_PDF_OPERAND = 0;

    private static final double A4_Height = 841.995;
    private static final double A4_Width = 595.35;
    private static List<String> letterArr = new ArrayList();

    static {
        for(char a = 97;a<=122;a++){
            letterArr.add(String.valueOf(a));
        }
    }

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
            ax.setProperty("DisplayAlerts",new Variant(false));
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
            Assert.notNull(sheets,"未能获取到工作簿，请联系管理员或检查sheets");
            int count = Dispatch.get(sheets, "Count").getInt();//获取sheets的数量

            List<String> remainSheets = printSetup.getRemainSheets();

            for (int j = 1; j <=count; j++) {
                Dispatch currentSheet = Dispatch.invoke(sheets, "Item",
                        Dispatch.Get, new Object[] { new Integer(j) },
                        new int[1]).toDispatch();//获取每个sheet
                Dispatch.put(currentSheet,"DisplayPageBreaks",new Variant(true));
                String sheetName = Dispatch.get(currentSheet,"Name").toString();

                /*if(remainSheets.size()<=0||!remainSheets.contains(sheetName)){
                    Assert.isTrue(count>1,"当前工作薄保留的sheet少于1，请重新设置");
                    Dispatch.call(currentSheet,"Delete");
                    count--;j--;
                    continue;
                }
*/
                Dispatch pageSetup = Dispatch.get(currentSheet, "PageSetup")
                        .toDispatch();//获取pageSetup对象，以设置打印参数
                Dispatch.put(pageSetup,"AlignMarginsHeaderFooter",new Variant(true));
                Dispatch.put(pageSetup,"ScaleWithDocHeaderFooter",new Variant(true));
                Dispatch.put(pageSetup,"PrintGridlines",new Variant(false));
                Dispatch.put(pageSetup,"CenterHorizontally",printSetup.getCenterHorizontally());
                Dispatch.put(pageSetup,"CenterVertically",printSetup.getCenterVertically());
                Dispatch.put(pageSetup,"PaperSize",new Variant(9));
                Variant defalutVariant = printSetup.getJacobVariantByOrientation(sheetName);
                Dispatch.put(pageSetup, "Orientation", defalutVariant);
                //System.out.println(printSetup.getSheetOrientation(sheetName));
                if(printSetup.getSheetOrientation(sheetName)){
                    Map<String,Double> portraitSet = printSetup.getPortraitSet();
                    setMarginsByParam(pageSetup,portraitSet);
                }else{
                    Map<String,Double> landScapeSet = printSetup.getLandscapeSet();
                    setMarginsByParam(pageSetup,landScapeSet);
                }
                Map<String,String> sheetValue = printSetup.getSheetsValues(sheetName);
                if(sheetValue!=null&&sheetValue.get("topTitleRows")!=null&&StringUtils.hasText(sheetValue.get("topTitleRows"))&&sheetValue.get("topTitleRows").matches("\\$[1-9]\\d*:\\$[1-9]\\d*")){
                    Dispatch.put(pageSetup,"PrintTitleRows",sheetValue.get("topTitleRows"));
                }
                Dispatch pageBreaks = Dispatch.get(currentSheet,"HPageBreaks").toDispatch();
                int pageBreaksCount = Dispatch.get(pageBreaks,"Count").getInt();
                System.out.println(pageBreaksCount);
                if(sheetValue!=null&&sheetValue.get("bottomTitleRows")!=null&&StringUtils.hasText(sheetValue.get("bottomTitleRows"))&&sheetValue.get("bottomTitleRows").matches("\\$[1-9]\\d*:\\$[1-9]\\d*")&&pageBreaksCount!=0){
                    boolean keepBottom = true;

                    String topTitleRowsStr = sheetValue.get("topTitleRows");
                    String bottomTitleRowsStr = sheetValue.get("bottomTitleRows");
                    List<Integer> topTitleRowIndexs = new ArrayList<>();
                    if(sheetValue.get("topTitleRows")!=null&&sheetValue.get("topTitleRows").matches("\\$[1-9]\\d*:\\$[1-9]\\d*")){
                        String[] topTitleRowSplit = topTitleRowsStr.split(":");
                        for(int splitArrayIndex = 0;splitArrayIndex<topTitleRowSplit.length;splitArrayIndex++){
                            topTitleRowSplit[splitArrayIndex]=topTitleRowSplit[splitArrayIndex].replace("$","");
                            if (!topTitleRowIndexs.contains(topTitleRowSplit[splitArrayIndex])) topTitleRowIndexs.add(Integer.parseInt(topTitleRowSplit[splitArrayIndex]));
                        }
                    }
                    if (topTitleRowIndexs.size()==2){
                        for(int i=Integer.min(topTitleRowIndexs.get(0),topTitleRowIndexs.get(1))+1;i<Integer.max(topTitleRowIndexs.get(0),topTitleRowIndexs.get(1));i++){
                            topTitleRowIndexs.add(i);
                        }
                    }

                    Dispatch.invoke(sheets,"Add", Dispatch.Method,new Object[]{
                            null, Dispatch.invoke(sheets, "Item",
                            Dispatch.Get, new Object[] { new Integer(count) },
                            new int[1]).toDispatch(),new Variant(1),new Variant(-4167)
                    },new int[1]);

                    Dispatch bottomTitleRange = Dispatch.invoke(currentSheet, "Range",
                            Dispatch.Get, new Object[] { bottomTitleRowsStr },
                            new int[1]).toDispatch();
                    double bottomTitleHeight = Dispatch.get(bottomTitleRange,"Height").getDouble();
                    Dispatch bottomTitleRows = Dispatch.get(bottomTitleRange,"Rows").toDispatch();
                    int bottomTitleRowsCount = Dispatch.get(bottomTitleRows,"Count").getInt();

                    Dispatch newlySheet = Dispatch.invoke(sheets, "Item",
                            Dispatch.Get, new Object[] { new Integer(count+1) },
                            new int[1]).toDispatch();

                    /**
                     * 图形复制
                     */
                    /*Dispatch.invoke(bottomTitleRange,"CopyPicture",Dispatch.Method,new Object[]{
                            new Variant(2),new Variant(1)
                    },new int[1]);

                    Dispatch.invoke(newlySheet,"Paste",Dispatch.Method,new Object[]{
                            Dispatch.invoke(newlySheet, "Range",Dispatch.Get, new Object[] { "A1" },new int[1]).toDispatch()
                    },new int[1]);*/

                    /**
                     *  普通复制
                     */
                    Dispatch.invoke(bottomTitleRange,"Copy", Dispatch.Method,new Object[]{
                    },new int[1]);

                    Dispatch.invoke(Dispatch.invoke(newlySheet, "Range", Dispatch.Get, new Object[] { "A1" },new int[1]).toDispatch(),"PasteSpecial", Dispatch.Method,new Object[]{
                            new Variant(8)
                    },new int[1]);
                    Dispatch.invoke(Dispatch.invoke(newlySheet, "Range", Dispatch.Get, new Object[] { "A1" },new int[1]).toDispatch(),"PasteSpecial", Dispatch.Method,new Object[]{
                            new Variant(14)
                    },new int[1]);

                    Dispatch.call(bottomTitleRange,"Delete");

                    boolean block = false;

                    Dispatch currentRow = Dispatch.invoke(currentSheet, "Range", Dispatch.Get, new Object[] { "$"+(32)+":$"+(32) },new int[1]).toDispatch();
                    Dispatch.put(currentRow,"RowHeight",22.5);
                    Dispatch cell = Dispatch.invoke(currentRow, "Item",
                            Dispatch.Get, new Object[] { new Integer(1) },
                            new int[1]).toDispatch();
                    Dispatch.put(cell,"Value","111");


                    for(int r=1;r<=pageBreaksCount;r++){
                        Dispatch pageBreak = Dispatch.invoke(pageBreaks, "Item",
                                Dispatch.Get, new Object[] { new Integer(r) },
                                new int[1]).toDispatch();
                        Dispatch location = Dispatch.get(pageBreak,"Location").toDispatch();
                        String address = Dispatch.get(location,"Address").toString();
                        String lineNoStr = address.substring(address.lastIndexOf("$")+1);
                        Integer lineNo = Integer.parseInt(lineNoStr);
                        //System.out.println(lineNoStr);

                        Dispatch currentLineRanges = Dispatch.invoke(currentSheet, "Range",
                                Dispatch.Get, new Object[] { "$"+(lineNo-bottomTitleRowsCount)+":$"+(lineNo-1) },
                                new int[1]).toDispatch();
                        double currentLineRangeHeight = Dispatch.get(currentLineRanges,"Height").getDouble();

                        if(bottomTitleHeight>currentLineRangeHeight) {
                            int endIndex = 0;
                            for(int x = 1;x<=10;x++){
                                Dispatch xxx = Dispatch.invoke(currentSheet, "Range",
                                        Dispatch.Get, new Object[] { "$"+(lineNo-bottomTitleRowsCount-x)+":$"+(lineNo-1) },
                                        new int[1]).toDispatch();
                                double xxxHeight = Dispatch.get(xxx,"Height").getDouble();
                                if(xxxHeight>=bottomTitleHeight) {
                                    endIndex=x;
                                    break;
                                }
                            }
                            lineNo = lineNo -endIndex;
                        }
                        copyAndPasteBottomTitleRows(currentSheet,bottomTitleRowsCount,newlySheet,lineNo-bottomTitleRowsCount,r==1?true:false);
                        Dispatch newlyPageBreaks = Dispatch.get(currentSheet,"HPageBreaks").toDispatch();
                        pageBreaksCount = Dispatch.get(newlyPageBreaks,"Count").getInt();

                        if (!block&&r==pageBreaksCount){
                            Dispatch usedRange = Dispatch.get(currentSheet,"UsedRange").toDispatch();
                            Dispatch rows = Dispatch.get(usedRange,"Rows").toDispatch();
                            int rowCount = Dispatch.get(rows,"Count").getInt();
                            if(keepBottom){
                                if(rowCount>lineNo){
                                    
                                }
                            }else{
                                if(rowCount>lineNo){
                                    copyAndPasteBottomTitleRows(currentSheet,bottomTitleRowsCount,newlySheet,rowCount+1,r==1?true:false);
                                }
                            }
                        }
                    }

                    //Dispatch.call(newlySheet,"Delete");
                }
                if(sheetValue!=null&&sheetValue.get("leftTitleRows")!=null&&StringUtils.hasText(sheetValue.get("leftTitleRows"))&&sheetValue.get("leftTitleRows").matches("\\$[A-Z]:\\$[A-Z]")){
                    Dispatch.put(pageSetup,"PrintTitleColumns",sheetValue.get("leftTitleRows"));
                }
                if(printSetup.headerOrFooterStart(printSetup.getHeader())!=0&&printSetup.headerOrFooterEnd(printSetup.getHeader())!=0&&j>=printSetup.headerOrFooterStart(printSetup.getHeader())&&j<=printSetup.headerOrFooterEnd(printSetup.getHeader())){
                    setPageHeader(printSetup,pageHeaderCount,pageSetup);
                }else if(printSetup.headerOrFooterStart(printSetup.getHeader())!=0&&printSetup.headerOrFooterEnd(printSetup.getHeader())==0&&j>=printSetup.headerOrFooterStart(printSetup.getHeader())){
                    setPageHeader(printSetup,pageHeaderCount,pageSetup);
                }else if(printSetup.headerOrFooterStart(printSetup.getHeader())==0&&printSetup.headerOrFooterEnd(printSetup.getHeader())!=0&&j<=printSetup.headerOrFooterEnd(printSetup.getHeader())){
                    setPageHeader(printSetup,pageHeaderCount,pageSetup);
                }
                if(printSetup.headerOrFooterStart(printSetup.getFooter())!=0&&printSetup.headerOrFooterEnd(printSetup.getFooter())!=0&&j>=printSetup.headerOrFooterStart(printSetup.getFooter())&&j<=printSetup.headerOrFooterEnd(printSetup.getFooter())){
                    setPageFooter(printSetup,pageFooterCount,pageSetup);
                }else if(printSetup.headerOrFooterStart(printSetup.getFooter())!=0&&printSetup.headerOrFooterEnd(printSetup.getFooter())==0&&j>=printSetup.headerOrFooterStart(printSetup.getFooter())){
                    setPageFooter(printSetup,pageFooterCount,pageSetup);
                }else if(printSetup.headerOrFooterStart(printSetup.getFooter())==0&&printSetup.headerOrFooterEnd(printSetup.getFooter())!=0&&j<=printSetup.headerOrFooterEnd(printSetup.getFooter())){
                    setPageFooter(printSetup,pageFooterCount,pageSetup);
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
    public static int beforeSheetPageNum(Dispatch sheets, int pageStartSheet){
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

    public static void copyAndPasteBottomTitleRows(Dispatch currentSheet, int bottomTitleRowsCount, Dispatch newlySheet, int currentRowIndex, boolean first){
        Dispatch currentRow = Dispatch.invoke(currentSheet, "Range", Dispatch.Get, new Object[] { "$"+currentRowIndex+":$"+currentRowIndex },new int[1]).toDispatch();
        for(int insertI = 0;insertI<bottomTitleRowsCount;insertI++){
            Dispatch.invoke(currentRow,"Insert", Dispatch.Method,new Object[]{
                    new Variant(-4121)
            },new int[1]);
        }

        /**
         * 图形复制
         */
        /*Dispatch shapes = Dispatch.get(newlySheet,"Shapes").toDispatch();
        Dispatch shape = Dispatch.invoke(shapes, "Item",
                Dispatch.Get, new Object[]{new Integer(1)},
                new int[1]).toDispatch();
        Dispatch.call(shape,"Copy");
        Dispatch.invoke(currentSheet,"Paste",Dispatch.Method,new Object[]{
                Dispatch.invoke(currentSheet, "Range",Dispatch.Get, new Object[] { "A"+currentRowIndex },new int[1]).toDispatch()
        },new int[1]);*/


       /* Dispatch.invoke(Dispatch.invoke(newlySheet, "Range",Dispatch.Get, new Object[] { "$"+1+":$"+bottomTitleRowsCount },new int[1]).toDispatch(),"Copy",Dispatch.Method,new Object[]{
        },new int[1]);

        Dispatch.invoke(Dispatch.invoke(currentSheet, "Range",Dispatch.Get, new Object[] {"$"+currentRowIndex+":$"+(currentRowIndex+bottomTitleRowsCount-1) },new int[1]).toDispatch(),"PasteSpecial",Dispatch.Method,new Object[]{
                new Variant(8)
        },new int[1]);

        Dispatch.invoke(Dispatch.invoke(currentSheet, "Range",Dispatch.Get, new Object[] { "$"+currentRowIndex+":$"+(currentRowIndex+bottomTitleRowsCount-1) },new int[1]).toDispatch(),"PasteSpecial",Dispatch.Method,new Object[]{
                new Variant(14)
        },new int[1]);*/

        /**
         * copyPicture复制
         */
        if(first) {
            Dispatch.invoke(Dispatch.invoke(newlySheet, "Range", Dispatch.Get, new Object[] { "$"+1+":$"+bottomTitleRowsCount },new int[1]).toDispatch(),"CopyPicture", Dispatch.Method,new Object[]{
                    new Variant(2),new Variant(1)
            },new int[1]);
        }
        copyPicture(newlySheet,bottomTitleRowsCount,currentSheet,currentRowIndex);


    }

    public static void copyPicture(Dispatch newlySheet, int bottomTitleRowsCount, Dispatch currentSheet, int currentRowIndex){
        Dispatch.invoke(Dispatch.invoke(newlySheet, "Range", Dispatch.Get, new Object[] { "$"+1+":$"+bottomTitleRowsCount },new int[1]).toDispatch(),"CopyPicture", Dispatch.Method,new Object[]{
                new Variant(2),new Variant(1)
        },new int[1]);

        Dispatch.invoke(currentSheet,"Paste", Dispatch.Method,new Object[]{
                Dispatch.invoke(currentSheet, "Range", Dispatch.Get, new Object[] { "$"+currentRowIndex+":$"+(currentRowIndex+bottomTitleRowsCount-1) },new int[1]).toDispatch()
        },new int[1]);
    }

    public static void setMarginsByParam(Dispatch pageSetup, Map<String,Double> param){
        if(param!=null&&getDouble(param,"leftMargin")!=null) Dispatch.put(pageSetup, "LeftMargin", new Variant(getDouble(param,"leftMargin")*28.35));
        if(param!=null&&getDouble(param,"rightMargin")!=null) Dispatch.put(pageSetup, "RightMargin", new Variant(getDouble(param,"rightMargin")*28.35));
        if(param!=null&&getDouble(param,"topMargin")!=null) Dispatch.put(pageSetup, "TopMargin", new Variant(getDouble(param,"topMargin")*28.35));
        if(param!=null&&getDouble(param,"bottomMargin")!=null) Dispatch.put(pageSetup, "BottomMargin", new Variant(getDouble(param,"bottomMargin")*28.35));
        if(param!=null&&getDouble(param,"headerMargin")!=null) Dispatch.put(pageSetup,"HeaderMargin",new Variant(getDouble(param,"headerMargin")*28.35));
        if(param!=null&&getDouble(param,"footerMargin")!=null) Dispatch.put(pageSetup,"FooterMargin",new Variant(getDouble(param,"footerMargin")*28.35));
    }

    public static Double getDouble(Map setup,String key){
        Object object = setup.get(key);
        if(object==null) return null;
        try {
            String doubleStr = object.toString();
            Double thisDouble = Double.parseDouble(doubleStr);
            return thisDouble;
        }catch (Exception e){
            return null;
        }
    }

    protected static void setPageHeader(PrintSetup printSetup,int pageHeaderCount,Dispatch pageSetup){
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

    protected static void setPageFooter(PrintSetup printSetup,int pageFooterCount,Dispatch pageSetup){
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
    }
}
