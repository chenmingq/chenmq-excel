package com.github.mcin123.excel.annotation;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;


/**
 * @author chenmq
 * @version V1.0
 * ProjectName: chenmq-excel
 * Package com.github.mcin123.excel.annotation
 * Description: 单元格信息
 * date 2018-04-08 下午7:23
 */


@Inherited
@Documented
@Target({FIELD,ElementType.TYPE, ElementType.METHOD})
@Retention(RUNTIME)
public @interface ExcelShell {

    /**
     * 表格字段名称
     */
    String cellName() default "";

    /**
     * 表格字段位置
     */
    int sheetPosition() default 0;

    /**
     * 时间格式
     */
    String dateFormat() default "m/d/yy h:mm";

    /**
     * 字体名称
     */
    String fontName() default "Courier New";

    /**
     * 字体大小
     */
    short fontSize() default 12;

    /**
     * 是否斜体
     */
    boolean isItalic() default false;

    /**
     * 是否中划线
     */
    boolean isStrikeout() default false;

    /**
     * 水平对齐方式
     */
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.CENTER;

    /**
     * 垂直排列方式
     */
    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;
}
