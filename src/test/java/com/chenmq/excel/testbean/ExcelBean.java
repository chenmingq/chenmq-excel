package com.chenmq.excel.testbean;

import com.chenmq.excel.annotation.ExcelShell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author chenmq
 * @version V1.0
 * @ProjectName: chenmq-excel
 * @Package com.chenmq.excel.testbean
 * @Description: TODO
 * @date 2018-04-13 上午12:36
 */

public class ExcelBean {

    @ExcelShell(cellName = "创建时间",sheetPosition = 0,dateFormat = "yyyy-MM-dd HH:mm:ss")
    private Date createDate;

    @ExcelShell(cellName = "字符串",sheetPosition = 1,isStrikeout = true,fontSize = 24 )
    private String str;

    @ExcelShell(cellName = "字符串url链接",sheetPosition = 2,isStrikeout = true)
    private String link;

    @ExcelShell(cellName = "boolean类型",sheetPosition = 3,horizontalAlignment= HorizontalAlignment.LEFT)
    private boolean flage;

    @ExcelShell(cellName = "BigDecimal类型",sheetPosition = 4,verticalAlignment = VerticalAlignment.BOTTOM)
    private BigDecimal money;

    @ExcelShell(cellName = "Integer类型",sheetPosition = 5,isItalic = true)
    private Integer ints;

    @ExcelShell(cellName = "Double类型",sheetPosition = 6,fontName = "Airal")
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
