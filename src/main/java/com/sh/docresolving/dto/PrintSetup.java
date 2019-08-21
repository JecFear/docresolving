package com.sh.docresolving.dto;

import com.jacob.com.Variant;
import lombok.Data;
import org.springframework.util.StringUtils;

import java.util.HashMap;

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
        return (Boolean) object;
    }

    public Integer getInt(String key){
        Object object = super.get(key);
        if(object == null) return null;
        return (Integer) object;
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
        return getBoolean(key);
    }

    public Boolean needPageNum(){
        Boolean needPageNum = getBoolean("needPageNum");
        if(needPageNum == null) return true;
        return needPageNum;
    }

    public Integer pageNumStart(){
        Integer pageNumStart = getInt("pageNumStart");
        if(pageNumStart == null) return 2;
        return pageNumStart;
    }

    public Integer headerStart(){
        Integer headerStart = getInt("headerStart");
        if(headerStart == null) return 2;
        return headerStart;
    }
}
