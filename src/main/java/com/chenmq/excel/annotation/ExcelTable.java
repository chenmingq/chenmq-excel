package com.chenmq.excel.annotation;

/**
 * @author chenmq
 * @version V1.0
 * @ProjectName: chenmq-excel
 * @Package com.chenmq.excel.annotation
 * @Description: TODO
 * @date 2018-04-13 上午1:22
 */

public @interface ExcelTable {

    /**
     * excel类型 (2003[xls] || 2007[xlsx])
     * @return
     */
    String excelType () default "xls";

    /**
     * excel sheet名称
     * @return
     */
    String sheetName () default "";

    /**
     * 标题字体颜色
     * @return
     */
    String titFontColor () default "";

    /**
     * 标题字体大小
     * @return
     */
    int titleFontSize() default 12;

    /**
     * 标题背景颜色
     * @return
     */
    String titleBackgroudColor () default "";

}
