package com.sh.docresolving.dto;

import lombok.Data;

@Data
public class MergeBack {

    private int rowNum;

    private int start;

    private int end;

    public MergeBack() {
    }

    public MergeBack(int start, int end) {
        this.start = start;
        this.end = end;
    }

    public MergeBack(int rowNum, int start, int end) {
        this.rowNum = rowNum;
        this.start = start;
        this.end = end;
    }
}
