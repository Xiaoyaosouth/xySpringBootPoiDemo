package com.fe.mysbpoidemo.model;

import lombok.Data;

import java.util.List;

/**
 * 对象模型
 */
@Data
public class ExcelCommon {
    // 第一行的属性
    private List<String> property;
    // 属性列下的每个单元格内容
    private List<List<String>> cellData;
}
