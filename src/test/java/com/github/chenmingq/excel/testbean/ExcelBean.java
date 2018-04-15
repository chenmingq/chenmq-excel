package com.github.chenmingq.excel.testbean;

import com.github.chenmingq.excel.annotation.ExcelSheet;
import com.github.chenmingq.excel.annotation.ExcelTable;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author chenmq
 * @version V1.0
 * ProjectName: chenmq-excel
 * Package com.github.chenmingq.excel.testbean
 * Description: excel待转化对象
 * @date 2018-04-13 上午12:36
 */

@ExcelTable(sheetName="excelName",titleFontColor = IndexedColors.BLUE,
        titleBackgroundColor = IndexedColors.YELLOW,titleFontName = "Tlwg typo",titleFontSize = 20)
public class ExcelBean {

    @ExcelSheet(cellName = "创建时间",sheetPosition = 0,dateFormat = "yyyy-MM-dd HH:mm:ss",fontBackgroundColor = IndexedColors.BLUE)
    private Date createDate;

    @ExcelSheet(cellName = "字符串",sheetPosition = 1,isStrikeout = true,fontSize = 24 )
    private String str;

    @ExcelSheet(cellName = "字符串url链接",sheetPosition = 2,isStrikeout = true)
    private String link;

    @ExcelSheet(cellName = "boolean类型",sheetPosition = 3,horizontalAlignment= HorizontalAlignment.LEFT)
    private boolean flage;

    @ExcelSheet(cellName = "BigDecimal类型",sheetPosition = 4,verticalAlignment = VerticalAlignment.BOTTOM)
    private BigDecimal money;

    @ExcelSheet(cellName = "Integer类型",sheetPosition = 5,isItalic = true)
    private Integer ints;

    @ExcelSheet(cellName = "Double类型",sheetPosition = 6,fontName = "Airal",fontColor = IndexedColors.YELLOW)
    private double sources;

    public Date getCreateDate() {
        return createDate;
    }

    public void setCreateDate(Date createDate) {
        this.createDate = createDate;
    }

    public String getStr() {
        return str;
    }

    public void setStr(String str) {
        this.str = str;
    }

    public String getLink() {
        return link;
    }

    public void setLink(String link) {
        this.link = link;
    }

    public boolean isFlage() {
        return flage;
    }

    public void setFlage(boolean flage) {
        this.flage = flage;
    }

    public BigDecimal getMoney() {
        return money;
    }

    public void setMoney(BigDecimal money) {
        this.money = money;
    }

    public Integer getInts() {
        return ints;
    }

    public void setInts(Integer ints) {
        this.ints = ints;
    }

    public double getSources() {
        return sources;
    }

    public void setSources(double sources) {
        this.sources = sources;
    }

    @Override
    public String toString() {
        return "ExcelBean{" +
                "createDate=" + createDate +
                ", str='" + str + '\'' +
                ", link='" + link + '\'' +
                ", flage=" + flage +
                ", money=" + money +
                ", ints=" + ints +
                ", sources=" + sources +
                '}';
    }
}
