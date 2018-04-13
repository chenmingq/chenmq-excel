package com.chenmq.excel.utils;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author chenmq
 * @version V1.0
 * @ProjectName: chenmq-excel
 * @Package com.chenmq.excel.utils
 * @Description: TODO
 * @date 2018-04-10 上午12:58
 */

public class IoOptionUtils {
    
    /**
     * 写文件
     * @param workbook
     * @param outputStream
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
     * @param workbook
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
     * @param outputStream
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
     * @param inputStream
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
     * @param byteArrayOutputStream
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
