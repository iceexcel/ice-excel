package com.ice.excel;

public abstract class AbstractExcelParser{
    private String excelFileName;
    private String sheet;
    protected void setExcelFileName(String excelFileName) {
        this.excelFileName = excelFileName;
    }
    protected void setSheet(String sheet) {
        this.sheet = sheet;
    }
    public String getExcelFileName() {
        return excelFileName;
    }
    public String getSheet() {
        return sheet;
    }
    
}
