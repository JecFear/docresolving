package com.sh.docresolving.entity;

public class MergeBack {

    private int rowNum;

    private int start;

    private int end;

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public int getStart() {
        return start;
    }

    public void setStart(int start) {
        this.start = start;
    }

    public int getEnd() {
        return end;
    }

    public void setEnd(int end) {
        this.end = end;
    }

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
