package com.github.chenmingq.excel.utils;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author chenmq
 */

public class DateUtil {

    /**
     * 格式化日期
     *
     * @return Date时间
     */
    public static Date fomatDate(String date,String simpleDateFormat) {
        DateFormat fmt = new SimpleDateFormat(simpleDateFormat);
        try {
            return fmt.parse(date);
        } catch (ParseException e) {
            e.printStackTrace();
            return null;
        }
    }

}
