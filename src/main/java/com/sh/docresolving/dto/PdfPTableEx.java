package com.sh.docresolving.dto;

import com.itextpdf.text.pdf.PdfPCell;
import lombok.Data;

import java.util.List;

@Data
public class PdfPTableEx {

    private List<PdfPCell> cells;

    private float[] widths;
}
