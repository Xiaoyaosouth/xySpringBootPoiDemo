package com.fe.mysbpoidemo.model;

import lombok.Data;

import java.util.List;

@Data
public class ExcelCommon {
    private List<String> property;
    private List<List<String>> cellData;
}
