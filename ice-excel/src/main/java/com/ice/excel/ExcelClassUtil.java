package com.ice.excel;

public class ExcelClassUtil {
    public static boolean hasClass(String className) {
        try {
            Class<?> clazz = Class.forName(className);
            return true;
        } catch (ClassNotFoundException e) {
            return false;
        }
    }
}
