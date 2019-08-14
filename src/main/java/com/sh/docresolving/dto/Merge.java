package com.sh.docresolving.dto;

import lombok.Data;

@Data
public class Merge  {

    private int rowpan;

    private int colspan;

    public Merge() {
    }

    public Merge(int rowpan, int colspan) {
        this.rowpan = rowpan;
        this.colspan = colspan;
    }
}
