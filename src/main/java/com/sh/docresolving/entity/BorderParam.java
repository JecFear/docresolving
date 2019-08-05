package com.sh.docresolving.entity;

public class BorderParam {

    private int maxCol;

    private int maxRow;

    private int maxHegiht;

    public int getMaxCol() {
        return maxCol;
    }

    public void setMaxCol(int maxCol) {
        this.maxCol = maxCol;
    }

    public int getMaxRow() {
        return maxRow;
    }

    public void setMaxRow(int maxRow) {
        this.maxRow = maxRow;
    }

    public int getMaxHegiht() {
        return maxHegiht;
    }

    public void setMaxHegiht(int maxHegiht) {
        this.maxHegiht = maxHegiht;
    }

    public BorderParam(int maxCol, int maxRow, int maxHegiht) {
        this.maxCol = maxCol;
        this.maxRow = maxRow;
        this.maxHegiht = maxHegiht;
    }
}
