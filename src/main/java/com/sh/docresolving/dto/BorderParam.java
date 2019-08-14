package com.sh.docresolving.dto;

import lombok.Data;

@Data
public class BorderParam {

    private int maxCol;

    private int maxRow;

    private int maxHegiht;

    public BorderParam(int maxCol, int maxRow, int maxHegiht) {
        this.maxCol = maxCol;
        this.maxRow = maxRow;
        this.maxHegiht = maxHegiht;
    }
}
