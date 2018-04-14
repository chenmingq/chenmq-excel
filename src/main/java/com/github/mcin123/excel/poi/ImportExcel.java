package com.github.mcin123.excel.poi;

import com.github.mcin123.excel.annotation.ExcelSheet;
import com.github.mcin123.excel.utils.DateUtil;
import com.github.mcin123.excel.utils.IoOptionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

/**
 * @author chenmq
 * @version V1.0
 * ProjectName: chenmq-excel
 * Package com.github.mcin123.excel.poi
 * Description: 导入excel
 * date 2018-04-10 上午1:00
 */

public class ImportExcel {


    /**
     * file写入
     * @param excelFilePath File
     * @return 表格数据位置和数据
     */
    private static List<Map<Integer,String>> readInputFile (File excelFilePath){
        List<Map<Integer,String>> list = new ArrayList<Map<Integer,String>>();
        try {
            InputStream inp = new FileInputStream(excelFilePath);
            list = inStream(inp);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return list;
    }

    /**
     * 文件路径写入
     * @param excelNamePath String
     * @return 表格数据位置和数据
     */
    private static  List<Map<Integer,String>> readInputName (String excelNamePath){
        List<Map<Integer,String>> list = new ArrayList<Map<Integer,String>>();
        try {
            InputStream inp = new FileInputStream(excelNamePath);
            list = inStream(inp);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return list;
    }


    /**
     * InputStream 初始化写入表格
     * @param inp InputStream
     * @return 表格数据位置和数据
     */
    private static List<Map<Integer,String>> inStream (InputStream inp){

        List<Map<Integer,String>> list = new ArrayList<Map<Integer,String>>();

        HSSFWorkbook wb = null;
        try {
            wb = new HSSFWorkbook(new POIFSFileSystem(inp));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            IoOptionUtils.closeInputStream(inp);
        }

        DataFormatter formatter = new DataFormatter();

        for (Sheet sheet : wb) {
            int i = 0;
            for (Row row : sheet) {
                Map<Integer,String> cellMap = new HashMap<Integer, String>();
                i++;
                for (Cell cell : row) {
                    if (i==1) {
                        continue;
                    }
                    String text = formatter.formatCellValue(cell);
                    cellMap.put(cell.getColumnIndex(),text);
                }
                if (!cellMap.isEmpty()) {
                    list.add(cellMap);
                }
            }
        }

        IoOptionUtils.closeWorkbook(wb);

        return list;
    }

    /**
     * 文件路径名称 excel 转 Object 对象
     * @param clazz Class
     * @param filePathName String
     * @return 返回Java对象
     */
    public static List<?> importExcelToObject (Class<?> clazz,String filePathName){
        List<Map<Integer,String>> list = readInputName(filePathName);
        return initBean(clazz,list);
    }

    /**
     * file excel 转 Object 对象
     * @param clazz Class
     * @param file File
     * @return 返回List数组对象
     */
    public static List<?> importExcelToObject (Class<?> clazz,File file){
        List<Map<Integer,String>> list = readInputFile(file);
        return initBean(clazz,list);
    }

    /**
     * 初始化转换的对象
     * @param clazz Class
     * @param list List<Map<Integer,String>>
     * @return 返回List数组对象
     */
    private static List<?> initBean (Class<?> clazz,List<Map<Integer,String>> list){

        Field[] fields = clazz.getDeclaredFields();

        List<Object> lists = new ArrayList<Object>();

        for (Map<Integer,String> map:list) {

            Object obj = null;

            try {
                obj = clazz.newInstance();
            } catch (InstantiationException e) {
                e.printStackTrace();
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }

            for (Field field: fields) {
                field.setAccessible(true);
                ExcelSheet excelSheet = field.getAnnotation(ExcelSheet.class);
                try {
                    Object objValue = map.get(excelSheet.sheetPosition());

                    String simpleName = field.getType().getSimpleName();

                    if ("Date".equals(simpleName)) {
                        field.set(obj,objValue !=null?DateUtil.fomatDate(objValue.toString(), excelSheet.dateFormat()):null);
                    } else if ("Integer".equalsIgnoreCase(simpleName)) {
                        field.set(obj,null!=objValue?Integer.parseInt(objValue.toString()):null);
                    } else if ("boolean".equalsIgnoreCase(simpleName)) {
                        field.set(obj,null!= objValue?Boolean.valueOf(objValue.toString()):null);
                    } else if ("BigDecimal".equalsIgnoreCase(simpleName)) {
                        field.set(obj,null!= objValue?new BigDecimal(objValue.toString()):null);
                    } else if ("Double".equalsIgnoreCase(simpleName)) {
                        field.set(obj,null!= objValue?Double.parseDouble(objValue.toString()):null);
                    } else {
                        field.set(obj,objValue);
                    }

                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }

            lists.add(obj);
        }

        return lists;
    }

}
