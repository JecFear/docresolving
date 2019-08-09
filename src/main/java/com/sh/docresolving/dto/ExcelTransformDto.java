package com.sh.docresolving.dto;

public class ExcelTransformDto {

    private String fileIn;

    private String fileout;

    private PrintSetup printSetup;

    public String getFileIn() {
        return fileIn;
    }

    public void setFileIn(String fileIn) {
        this.fileIn = fileIn;
    }

    public String getFileout() {
        return fileout;
    }

    public void setFileout(String fileout) {
        this.fileout = fileout;
    }

    public PrintSetup getPrintSetup() {
        return printSetup;
    }

    public void setPrintSetup(PrintSetup printSetup) {
        this.printSetup = printSetup;
    }
}
