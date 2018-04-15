package com.github.chenmingq.excel.annotation;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

/**
 * @author chenmq
 * @version V1.0
 * ProjectName: chenmq-excel
 * Package com.github.chenmingq.excel.annotation
 * Description: 表格类型等信息注解
 * date 2018-04-13 上午1:22
 */
@Documented
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelTable {

    /**
     * excel类型 (xls || xlsx)
     */
    @Deprecated
    String excelType () default "xls";

    /**
     * excel sheet名称
     */
    String sheetName () default "";

    /**
     * 标题字体颜色
     */
    IndexedColors titleFontColor () default IndexedColors.BLACK;

    /**
     * 标题字体大小
     */
    short titleFontSize() default 12;


    /**
     * 字体名称
     */
    String titleFontName () default "";

    /**
     * 标题背景颜色
     */
    IndexedColors titleBackgroundColor () default IndexedColors.WHITE ;

}
