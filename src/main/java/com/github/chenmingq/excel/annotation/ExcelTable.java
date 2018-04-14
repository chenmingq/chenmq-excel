package com.github.chenmingq.excel.annotation;

/**
 * @author chenmq
 * @version V1.0
 * ProjectName: chenmq-excel
 * Package com.github.chenmingq.excel.annotation
 * Description: 表格类型等信息注解
 * date 2018-04-13 上午1:22
 */

public @interface ExcelTable {

    /**
     * excel类型 (xls || xlsx)
     */
    String excelType () default "xls";

    /**
     * excel sheet名称
     */
    String sheetName () default "";

    /**
     * 标题字体颜色
     */
    String titeFontColor () default "";

    /**
     * 标题字体大小
     */
    int titleFontSize() default 12;

    /**
     * 标题背景颜色
     */
    String titleBackgroudColor () default "";

}
