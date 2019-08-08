package com.sh.docresolving.entity;

import com.jacob.com.Variant;

import java.util.HashMap;

public class PrintSetup extends HashMap<String,Boolean> {

    private static boolean vertical = true;
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
        if(orientation){
            return new Variant(1);
        }else{
            return new Variant(2);
        }
    }
}
