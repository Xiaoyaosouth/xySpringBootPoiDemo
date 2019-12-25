package com.fe.mysbpoidemo.service;

import com.fe.mysbpoidemo.model.*;
import com.fe.mysbpoidemo.util.*;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.*;

/**
 * 业务逻辑，用于对解析到的excel数据进行封装
 * @author rk
 */
public class ExcelDataService {

    /**
     * 获取动态对象数组
     * @param filePath
     * @return
     */
    public List<Object> getDynamicModelList(String filePath) {
        List<Object> dataList = new ArrayList<>();
        try {
            Sheet sheet = ExcelReader.getFirstSheet(filePath);
            List<String> headNameList = ExcelReader.getHeadNameList(sheet);

            // 设置类属性
            HashMap propertyMap = new HashMap();
            for (int i = 0; i < headNameList.size(); i++) {
                propertyMap.put(headNameList.get(i), Class.forName("java.lang.String"));
            }
            // 声明动态 Bean
            DynamicBean dynamicBean;
            // 解析excel
            if (sheet != null) {
                // 从第2行开始遍历每一行，组装对象
                for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                    Row row = sheet.getRow(i); // 第i行
                    // 当前行不为空
                    if (row != null) {
                        // 生成动态 Bean
                        dynamicBean = new DynamicBean(propertyMap);

                        // 记录是否整行为空（按属性数判断，存在非空属性单元格时不丢弃）
                        // 每存在一个空单元格就-1，为0时说明无价值，丢弃
                        int tempSum = headNameList.size();

                        // 遍历第i行每一列
                        for (int j = 0; j < headNameList.size(); j++) {
                            Cell cell = row.getCell(j); // 当前单元格
                            String value = ExcelReader.getCellValue(cell);
                            // 处理空值
                            if (value == null){
                                value = "null";
                                tempSum --;
                            }else {
                                System.out.println("第" + (i+1) +"行，第" + (j + 1) + "列，值：" + value);
                            }
                            // 这里处理每个单元格的值
                            dynamicBean.setValue(headNameList.get(j), value);
                        } // 遍历每一列结束
                        if (tempSum > 0){
                            Object object = dynamicBean.getObject();
                            dataList.add(object);
                        }else {
                            // System.out.println("丢弃第" + (i+1) + "行数据");
                        }

                    }
                } // 遍历每一行结束
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return dataList;
    }

}
