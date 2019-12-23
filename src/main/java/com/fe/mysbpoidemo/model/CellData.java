package com.fe.mysbpoidemo.model;

import java.util.List;

import lombok.Data;

@Data
public class CellData {
    private String property;
    private List<String> cellValue;
}
