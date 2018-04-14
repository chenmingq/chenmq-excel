package com.github.chenmingq.excel.run;


import com.github.chenmingq.excel.filter.ExcelException;
import com.github.chenmingq.excel.poi.ExportExcel;
import com.github.chenmingq.excel.poi.ImportExcel;
import com.github.chenmingq.excel.testbean.ExcelBean;
import org.junit.Test;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author chenmq
 * @version V1.0
 * @ProjectName: chenmq-excel
 * @Package com.github.chenmingq.excel.run
 * @Description: 测试
 * @date 2018-04-10 上午1:01
 */

public class Runs {

    @Test
    public void exportExcelTest() {
        try {

            List<ExcelBean> excelBeanList = new ArrayList<ExcelBean>();
            ExcelBean excelBean = new ExcelBean();
            excelBean.setCreateDate(new Date());
            excelBean.setFlage(false);
            excelBean.setMoney(new BigDecimal(100));
            excelBean.setSources(1.1);
            excelBean.setInts(100);
            excelBean.setStr("hello chenmq-excel");
            excelBean.setLink("http://doc.excel.chenmq.com");
            excelBeanList.add(excelBean);
            ExportExcel.exportToFile("chenmq-excel.xls",excelBeanList);
        } catch (ExcelException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void importExcelTest (){
        List<?> list = ImportExcel.importExcelToObject(ExcelBean.class,"chenmq-excel.xls");
        System.out.println(list);
        for (int i = 0; i < list.size(); i++) {
            System.out.println(list.get(i));
        }
    }
}
