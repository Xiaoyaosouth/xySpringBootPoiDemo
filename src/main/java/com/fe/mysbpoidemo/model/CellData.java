package com.fe.mysbpoidemo.model;

import java.util.List;

import lombok.Data;

@Data
public class CellData {
    // 属性
    private String property;
    // 单元格值
    private List<String> cellValue;
}
