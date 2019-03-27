package com.demo.utils;

public class ConstantUtil {
    //    判断是否为全数字
    public static boolean isDigit(String strNum) {
        return strNum.matches("[0-9]{1,}");
    }

    public static boolean firstDigit(String strNum) {
        strNum = strNum.substring(0, 1);
        return isDigit(strNum);
    }


}
