package com.demo.utils;

import java.text.SimpleDateFormat;
import java.util.Locale;

public class ConstantUtil {
    /**
     * 判断是否为全数字
     *
     * @param strNum
     * @return
     */
    public static boolean isDigit(String strNum) {
        return strNum.matches("[0-9]{1,}");
    }

    /**
     * 判断第一个是否是数字
     *
     * @param strNum
     * @return
     */
    public static boolean firstDigit(String strNum) {
        strNum = strNum.substring(0, 1);
        return isDigit(strNum);
    }

    /**
     * 将长整形转换成时间
     *
     * @param time
     * @param sdf
     * @return
     */
    public static String longToDate(Long time, String sdf) {
        Locale localeCN = Locale.SIMPLIFIED_CHINESE;
        SimpleDateFormat format = new SimpleDateFormat(sdf, localeCN);
        String d = format.format(time);
        return d;
    }
}
