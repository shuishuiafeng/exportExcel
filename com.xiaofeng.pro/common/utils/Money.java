package com.xiaofeng.pro.common.utils;

import java.text.DecimalFormat;

/**
 * @Author: Xiaofeng
 * @Date: 2019/4/12 9:56
 * @Description: 金额工具类
 */
public class Money {
    /**
     * 分转元（保留两位有效数字）字符串
     *
     * @param amount
     * @return
     */
    public static String centToYuanStr(Long amount) {
        DecimalFormat df = new DecimalFormat("######0.00");
        return df.format(centToYuanDouble(amount));
    }

    /**
     * 分转元（保留两位有效数字）Double
     *
     * @param amount
     * @return
     */
    public static Double centToYuanDouble(Long amount) {
        DecimalFormat df = new DecimalFormat("######0.00");
        Double d = Double.parseDouble(amount.toString()) / 100;
        return Double.valueOf(df.format(d));
    }

    /**
     * 元转分String
     *
     * @param amount
     * @return
     */
    public static String yuanToCentStr(String amount) {
        return yuanToCentLong(amount).toString();
    }

    /**
     * 元转分Long
     *
     * @param amount
     * @return
     */
    public static Long yuanToCentLong(String amount) {
        DecimalFormat df = new DecimalFormat("######0");
        Double d = Double.parseDouble(amount) * 100;
        return Long.valueOf(df.format(d));
    }

}
