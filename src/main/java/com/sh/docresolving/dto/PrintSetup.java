package com.sh.docresolving.dto;

import com.jacob.com.Variant;
import lombok.Data;

import java.util.HashMap;

@Data
public class PrintSetup extends HashMap<String,Boolean> {

    //竖向 true
    private static boolean vertical = true;
    //横向 false
    private static boolean horizontal =false;

    @Override
    public Boolean put(String key, Boolean value) {
        return super.put(key, value);
    }

    @Override
    public Boolean get(Object key) {
        return super.get(key);
    }

    public Variant getJacobVariantByOrientation(Object key){
        Boolean orientation = super.get(key);
        if(orientation == null) return null;
        if(orientation){
            return new Variant(1);
        }else{
            return new Variant(2);
        }
    }
}
