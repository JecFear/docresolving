package com.sh.docresolving.dto;

public class Merge  {

    private int rowpan;

    private int colspan;

    public Merge() {
    }

    public Merge(int rowpan, int colspan) {
        this.rowpan = rowpan;
        this.colspan = colspan;
    }

    public int getRowpan() {
        return rowpan;
    }

    public void setRowpan(int rowpan) {
        this.rowpan = rowpan;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }
}
