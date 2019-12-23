package com.fe.mysbpoidemo.model;

import java.util.List;
import lombok.Data;

@Data
public class ExcelCommonV2 {
    // 第一行的属性
    private List<String> property;
    // 属性列下的每个单元格内容
    private List<CellData> cellData;
}
