package com.fe.mysbpoidemo.util;

import java.util.Random;

/**
 * @author rk
 */
public class FileUtil {

    /**
     * 获取文件后缀
     * @param filePath
     * @return
     */
    public static String getFileSuffix(String filePath) {
        return filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
    }

    /**
     * 随机生成n位0~9数字组成的文件名
     *
     * @param fileSuffix
     * @param myLength
     * @return
     */
    public static String randomFileName(String fileSuffix, int myLength) {
        Random random = new Random();
        StringBuffer stringBuffer = new StringBuffer();
        for (int i = 0; i < myLength; i++) {
            stringBuffer.append(random.nextInt(10));
        }
        stringBuffer.append(".").append(fileSuffix);
        return stringBuffer.toString();
    }

    /**
     * 将https转为http
     * @param fileUrl
     * @return
     */
    public static String solveHttps(String fileUrl){
        String str = fileUrl.replaceFirst("https","http");
        return str;
    }
}
