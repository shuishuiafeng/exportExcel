package com.xiaofeng.pro.common.utils;

import org.apache.commons.lang3.EnumUtils;

/**
 * @Author: Xiaofeng
 * @Date: 2019/4/11 20:41
 * @Description:
 */
public class ExportEnumUtils  extends EnumUtils {
    public static String getEnumName(String code, Class<?> enumClass, String defaultValue) {
        for (Object str : enumClass.getEnumConstants()) {
            Object codeValue = ReflectionUtil.invokeGetter(str, "code");
            if (code.equals(String.valueOf(codeValue))) {
                return (String) ReflectionUtil.invokeGetter(str, "name");
            }
        }
        return defaultValue;
    }

    public static String getEnumCode(String name, Class<?> enumClass, String defaultValue) {
        for (Object str : enumClass.getEnumConstants()) {
            Object nameValue = ReflectionUtil.invokeGetter(str, "name");
            if (name.equals(String.valueOf(nameValue))) {
                return String.valueOf(ReflectionUtil.invokeGetter(str, "code"));
            }
        }
        return defaultValue;
    }
}
