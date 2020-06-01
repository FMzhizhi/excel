package com.hash.excel.dojo;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;
import lombok.ToString;

@Data
@ToString
public class ReadData {
    //设置列对应的属性
    @ExcelProperty(index = 0)
    private int sid;

    //设置列对应的属性
    @ExcelProperty(index = 1)
    private String sname;
}
