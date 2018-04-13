package com.chenmq.excel.testbean;

import com.chenmq.excel.annotation.ExcelShell;
import lombok.Data;

/**
 * @author chenmq
 * @version V1.0
 * @ProjectName: chenmq-excel
 * @Package com.chenmq.excel.testbean
 * @Description: TODO
 * @date 2018-04-09 下午11:08
 */

@Data
public class ShellName {

    @ExcelShell(cellName = "姓名",sheetPosition = 2)
    private String  name;
}
