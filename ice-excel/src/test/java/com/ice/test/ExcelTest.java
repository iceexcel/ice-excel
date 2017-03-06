package com.ice.test;

import org.junit.Test;

import com.ice.excel.IceExcel;
import com.ice.excel.IceExcelConfig;
import com.ice.excel.ParserType;
// clean package
public class ExcelTest {
    @Test
    public void export() {
        String[][] data = { { "aaa", "167" }, { "278", "bbb2" }, { "aaa3", "bbb3" }, { "aaa4", "bbb4" } };
        IceExcel iceExcel = new IceExcel("D:/test/test.xls");
        // IceExcel iceExcel=new
        // IceExcel("C:/Users/Administrator/Desktop/tomcat/xls.xls","jnm");
        // IceExcelConfig.setSheet(iceExcel,"你好");
        IceExcelConfig.setParserType(iceExcel, ParserType.POI);
        iceExcel.setData(data);
        // iceExcel.setSheet("测试文件");
    }

    @Test
    public void importx() {
        IceExcel iceExcel = new IceExcel("D:/test/test.xls");
        IceExcelConfig.setParserType(iceExcel, ParserType.POI);
        String[][] data = iceExcel.getData();
        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < data[i].length; j++) {
                System.out.print(data[i][j] + "\t  ");
            }
            System.out.println();
        }
    }
}
