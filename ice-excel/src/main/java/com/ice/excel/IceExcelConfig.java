package com.ice.excel;
/**
 * 对Excel操作对象的一些默认参数的修改
 * 
 * @author iceWater
 * @date 2017-01-20
 * @version 1.0
 */
public class IceExcelConfig {
    /**
     * 修改Excel导出对象的Sheet名称
     * @param iceExcel Excel操作对象
     * @param sheetName 需要修改成的新的Sheet的名称
     */
   public static void setSheet(IceExcel iceExcel,String sheetName){
       iceExcel.setSheetName(sheetName);
   }
   /**
    * 修改Excel导出对象的操作API类型
    * @param iceExcel Excel操作对象
    * @param parserType 需要使用的新的API类型名称 @see com.ice.excel.ParserType
    */
   public static void setParserType(IceExcel iceExcel,ParserType parserType) {
       iceExcel.setParserType(parserType);
   }
}
