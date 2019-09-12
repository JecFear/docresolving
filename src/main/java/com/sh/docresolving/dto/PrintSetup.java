package com.sh.docresolving.dto;

import com.jacob.com.Variant;
import lombok.Data;
import org.omg.CORBA.OBJ_ADAPTER;

import java.util.HashMap;
import java.util.Map;

@Data
public class PrintSetup extends HashMap<String,Object> {

    //竖向 true
    private static boolean vertical = true;
    //横向 false
    private static boolean horizontal =false;


    @Override
    public Object put(String key, Object value) {
        return super.put(key, value);
    }

    @Override
    public Object get(Object key) {
        return super.get(key);
    }

    public Boolean getBoolean(String key){
        Object object = super.get(key);
        if(object == null) return null;
        try {
            Boolean thisBoolean = (Boolean) object;
            return thisBoolean;
        }catch (Exception e){
            return null;
        }
    }

    public Integer getInt(String key){
        Object object = super.get(key);
        if(object == null) return null;
        try {
            Integer thisInteger = (Integer) object;
            return thisInteger;
        }catch (Exception e) {
            return null;
        }
    }

    public Double getDouble(String key){
        Object object = super.get(key);
        if(object==null) return null;
        try {
            String doubleStr = object.toString();
            Double thisDouble = Double.parseDouble(doubleStr);
            return thisDouble;
        }catch (Exception e){
            return null;
        }
    }

    public Variant getJacobVariantByOrientation(String key){
        Boolean orientation = this.getSheetOrientation(key);
        if(orientation == null) return new Variant(1);
        if(orientation){
            return new Variant(1);
        }else{
            return new Variant(2);
        }
    }

    public Boolean getSheetOrientation(String key){
        Boolean orientation = getBoolean(key);
        if(orientation == null) return true;
        return orientation;
    }

    public Map<String, Object> getHeader(){
        Object headerObj = this.get("header");
        try {
            Map header = (HashMap)headerObj;
            return header;
        }catch (Exception e) {
            return null;
        }
    }

    public Map<String, Object> getFooter(){
        Object footerObj = this.get("footer");
        try {
            Map footer = (HashMap)footerObj;
            return footer;
        }catch (Exception e) {
            return null;
        }
    }

    public Map<String,Object> getHeaderOrFooterPictureObject(Map<String,Object> headerOrFooter,String objStr){
        return ((HashMap)headerOrFooter.get(objStr));
    }

    public String getHeaderOrFooterPicture(Map<String,Object> headerOrFooterObject){
       return headerOrFooterObject.get("imgUrl").toString();
    }

    public Double getPictureParam(Map<String,Object> pictureObj,String key){
        Object object = pictureObj.get(key);
        if(object==null) return null;
        try {
            String doubleStr = object.toString();
            Double thisDouble = Double.parseDouble(doubleStr);
            return thisDouble;
        }catch (Exception e){
            return 0.0;
        }
    }

    public Integer headerOrFooterStart(Map<String,Object> headerOrFooter){
        if(headerOrFooter==null) return 0;
        Object pageNumStartObj = headerOrFooter.get("pageNumStart");
        if (pageNumStartObj==null) return 0;
        try {
            Integer pageNumStart = Integer.parseInt(pageNumStartObj.toString());
            return pageNumStart;
        }catch (Exception e){
            return 0;
        }
    }



    public Boolean getCenterHorizontally(){
        Boolean centerHorizontally = getBoolean("centerHorizontally");
        if(centerHorizontally==null) return false;
        return centerHorizontally;
    }

    public Boolean getCenterVertically(){
        Boolean centerVertically = getBoolean("centerVertically");
        if(centerVertically==null) return false;
        return centerVertically;
    }

    public String currentPageReplace(int pageOffset){
        return"&P-"+pageOffset;
    }

    public String totalPageReplace(int pageOffset){
        return "&N-"+pageOffset;
    }

    public String graphReplace(){
        return "&G";
    }

    public String dateReplace(){
        return "&D";
    }

    public String timeReplace(){
        return "&T";
    }

    public boolean doesContainsGraph(String headerOrFooter){
        if(headerOrFooter.contains("&G")) return true;
        return false;
    }

    public String getExactHeaderOrFooter(Map<String,Object> headerOrFooter,String objStr,int pageOffset){
        String obj = headerOrFooter.get(objStr).toString();
        obj=obj.replace("&[页码]",currentPageReplace(pageOffset));
        obj=obj.replace("&[总页数]",totalPageReplace(pageOffset));
        obj=obj.replace("&[日期]",dateReplace());
        obj=obj.replace("&[时间]",timeReplace());
        obj=obj.replace("&[图片]",graphReplace());
        return obj;
    }

    public Map<String,Double> getLandscapeSet(){
        return (HashMap)this.get("landscapeSet");
    }

    public Map<String,Double> getPortraitSet(){
        return (HashMap)this.get("portraitSet");
    }

    public Map<String,String> getSheetsValues(String sheetName){
        Map<String,Map<String,String>> sheetsValues = (HashMap)this.get("sheetsValues");
        return sheetsValues.get(sheetName);
    }

    /** commitExample
     * {
     * 	"landscape": "1-3",
     * 	"portrait": "",
     * 	"centerHorizontally": false,
     * 	"centerVertically": false,
     * 	"sheetsValues": {
     * 		"sheet1": {
     * 			"topTitleRows": "23123123",
     * 			"bottomTitleRows": "14",
     * 			"leftTitleRows": "1134"
     *                },
     * 		"sheet2": {
     * 			"topTitleRows": "2341",
     * 			"bottomTitleRows": "12341",
     * 			"leftTitleRows": "21"
     *        },
     * 		"sheet3": {
     * 			"topTitleRows": "113",
     * 			"bottomTitleRows": "1234",
     * 			"leftTitleRows": "1231"
     *        },
     * 		"Sheet4": {
     * 			"topTitleRows": "",
     * 			"bottomTitleRows": "",
     * 			"leftTitleRows": ""
     *        },
     * 		"Sheet5": {
     * 			"topTitleRows": "",
     * 			"bottomTitleRows": "",
     * 			"leftTitleRows": ""
     *        },
     * 		"Sheet6": {
     * 			"topTitleRows": "",
     * 			"bottomTitleRows": "",
     * 			"leftTitleRows": ""
     *        },
     * 		"Sheet7": {
     * 			"topTitleRows": "",
     * 			"bottomTitleRows": "",
     * 			"leftTitleRows": ""
     *        }* 	},
     * 	"landscapeSet": {
     * 		"leftMargin": 0.4,
     * 		"rightMargin": 0.3,
     * 		"topMargin": 0.4,
     * 		"bottomMargin": 0.2,
     * 	    "headerMargin":0.2,
     * 	    "footerMargin":0.2
     * 	}    ,
     * 	"portraitSet": {
     * 		"leftMargin": 0.3,
     * 		"rightMargin": 0.3,
     * 		"topMargin": 0.3,
     * 		"bottomMargin": 0.2,
     *      "headerMargin":0.2,
     *      "footerMargin":0.2
     *    },
     * 	"header": {
     * 		"pageNumStart": 2,
     * 		"leftHeader": "&[页码]&[总页数]&[日期]&[时间]",
     * 		"centerHeader": "&[页码]&[总页数]&[日期]&[时间]",
     * 		"rightHeader": "&[页码]&[总页数]&[日期]&[时间]",
     * 		"leftHeaderPicture": {
     * 			"imgUrl": "",
     * 			"width": "",
     * 			"height": ""
     *        },
     * 		"centerHeaderPicture": {
     * 			"imgUrl": "",
     * 			"width": "",
     * 			"height": ""
     *        },
     * 		"rightHeaderPicture": {
     * 			"imgUrl": "",
     * 			"width": "",
     * 			"height": ""
     *        }
     *    },
     * 	"footer": {
     * 		"pageNumStart": 1,
     * 		"leftFooter": "&[页码]&[总页数]&[日期]&[时间]",
     * 		"centerFooter": "&[页码]&[总页数]&[日期]&[时间]",
     * 		"rightFooter": "&[页码]&[总页数]&[日期]&[时间]&[图片]",
     * 		"leftFooterPicture": {
     * 			"imgUrl": "",
     * 			"width": "",
     * 			"height": ""
     *        },
     * 		"centerFooterPicture": {
     * 			"imgUrl": "",
     * 			"width": "",
     * 			"height": ""
     *        },
     * 		"rightFooterPicture": {
     * 			"imgUrl": "http://www.shouhouzn.net/group1/M00/00/1F/rBGmcV1yIMCAF2-5AAI43gaTEbw438.png",
     * 			"width": "",
     * 			"height": ""
     *        }
     *    },
     * 	"sheet1": false,
     * 	"sheet2": false,
     * 	"sheet3": false,
     * 	"Sheet4": true,
     * 	"Sheet5": true,
     * 	"Sheet6": true,
     * 	"Sheet7": true
     * }
     */
}
