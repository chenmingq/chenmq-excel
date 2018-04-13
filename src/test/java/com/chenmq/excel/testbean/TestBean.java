package com.chenmq.excel.testbean;

import com.chenmq.excel.annotation.ExcelShell;
import lombok.Data;

import java.util.Date;

/**
 * @author chenmq
 * @version V1.0
 * @ProjectName: chenmq-excel
 * @Package com.chenmq.excel.testbean
 * @Description: TODO
 * @date 2018-04-08 上午11:57
 */


@Data
public class TestBean {

    @ExcelShell(cellName = "时间",sheetPosition = 0,dateFormat = "MM月dd日 HH:mm:ss")
    private Date dateTime;

    @ExcelShell(cellName = "创建时间",sheetPosition = 1,dateFormat = "MM月dd日 HH:mm:ss")
    private Date createDate;

    @ExcelShell(cellName = "字符串",sheetPosition = 2,isItalic = true,isStrikeout = true)
    private String str;

    @ExcelShell(cellName = "boolean类型",sheetPosition = 3)
    private Boolean flage;

    @ExcelShell(cellName = "文本类型",sheetPosition = 4,fontName = "Airal")
    private String textType;

}