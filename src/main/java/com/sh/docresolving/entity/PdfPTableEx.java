package com.sh.docresolving.entity;

import com.itextpdf.text.pdf.PdfPCell;

import java.util.List;

public class PdfPTableEx {

    private List<PdfPCell> cells;

    private float[] widths;

    public List<PdfPCell> getCells() {
        return cells;
    }

    public void setCells(List<PdfPCell> cells) {
        this.cells = cells;
    }

    public float[] getWidths() {
        return widths;
    }

    public void setWidths(float[] widths) {
        this.widths = widths;
    }
}
