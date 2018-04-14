package com.github.chenmingq.excel.utils;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author chenmq
 * @version V1.0
 * ProjectName: chenmq-excel
 * Package com.github.chenmingq.excel.utils
 * Description: io流操作
 * date 2018-04-10 上午12:58
 */

public class IoOptionUtils {
    
    /**
     * 写文件
     * @param workbook Workbook
     * @param outputStream OutputStream
     */
    public static void writeWorkbook (Workbook workbook, OutputStream outputStream){
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            closeOutputStream(outputStream);
            closeWorkbook(workbook);
        }
    }

    /**
     * 关闭Workbook
     * @param workbook Workbook
     */
    public static void closeWorkbook (Workbook workbook){
        if (null != workbook) {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭OutputStream
     * @param outputStream OutputStream
     */
    public static void closeOutputStream (OutputStream outputStream) {
        if (null != outputStream) {
            try {
                outputStream.flush();
                outputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭InputStream
     * @param inputStream InputStream
     */
    public static void closeInputStream (InputStream inputStream) {
        if (null != inputStream) {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭ByteArrayOutputStream
     * @param byteArrayOutputStream ByteArrayOutputStream
     */
    public static void closeByteArrayOutputStream (ByteArrayOutputStream byteArrayOutputStream){
        if (null != byteArrayOutputStream) {
            try {
                byteArrayOutputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
