package com.fe.mysbpoidemo.service;

import com.fe.mysbpoidemo.model.ExcelCommon;
import com.fe.mysbpoidemo.util.ExcelReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

/**
 * 业务逻辑，处理Excel数据，封装成对象
 * @author rk
 */
public class CommonExcelReaderService {
    /**
     * 返回组装的excel数据对象（正式使用）
     * @param filePath
     * @return
     */
    public ExcelCommon getCommonExcelData(String filePath) {
        // 待组装的对象
        ExcelCommon data = new ExcelCommon();
        try {
            // 获取表格
            Sheet sheet = ExcelReader.getFirstSheet(filePath);
            // 获取第一行的属性
            List<String> headNameList = ExcelReader.getHeadNameList(sheet);
            // 放入第一行的属性
            data.setProperty(headNameList);
            // 处理表格内容
            if (sheet != null) {
                // 待组装的单元格数据
                List<List<String>> cellData = new ArrayList<>();
                // 获取总行数
                int rowCount = sheet.getPhysicalNumberOfRows();

                for (int i = 0; i < headNameList.size(); i++) {
                    // 待组装列数据
                    List<String> cellValues = new ArrayList<>();

                    // 从第2行开始
                    for (int j = 1; j < rowCount; j++) {
                        // 第j行
                        Row row = sheet.getRow(j);
                        // 第i列
                        Cell cell = row.getCell(i);
                        // 取第j行第i列的单元格值
                        String value = ExcelReader.getCellValue(cell);
                        System.out.println("第" + (j + 1) + "行，第" + (i + 1) + "列，值：" + value);
                        // 处理空值
                        if (value == null){
                            value = "null";
                        }
                        cellValues.add(value);
                    } // 遍历行结束
                    cellData.add(cellValues);
                } // 遍历列结束

                data.setCellData(cellData);
            } // 处理表格内容完成

        } catch (Exception e) {
            e.printStackTrace();
        }
        return data;
    }
}
