package com.fe.mysbpoidemo.service;

import com.fe.mysbpoidemo.model.UserXy;
import com.fe.mysbpoidemo.util.ExcelReader;
import com.fe.mysbpoidemo.util.ObjectUtil;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * 业务逻辑
 *
 * @author rk
 */
public class UserDataService {
    /**
     * 获取excel文件第一个表格中的内容，组装成对象数组
     *
     * @param filePath
     * @return
     */
    public List<UserXy> getModelList(String filePath) {
        Field[] fields = UserXy.class.getDeclaredFields();
        int fieldsLength = fields.length;
        System.out.println("属性数量：" + fieldsLength);
        List<UserXy> dataList = new ArrayList<>();
        try {
            Sheet sheet = ExcelReader.getFirstSheet(filePath);
            String[] headNames = ExcelReader.getHeadNames(sheet);
            if (sheet != null) {
                // 从第2行开始遍历每一行，组装对象
                for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                    Row row = sheet.getRow(i); // 第i行
                    // 当前行不为空
                    if (row != null) {
                        UserXy userXy = new UserXy();
                        // 遍历第i行每一列
                        for (int j = 0; j < headNames.length; j++) {
                            Cell cell = row.getCell(j); // 当前单元格
                            String value = ExcelReader.getCellValue(cell);
                            //System.out.println("第" + (i+1) +"行，第" + (j + 1) + "列，值：" + value);
                            if (value == null){
                                value = "null";
                            }
                            // 第j个属性名
                            String fieldName = fields[j].getName();
                            // 调用属性名对应的set方法
                            ObjectUtil.setValue(userXy, UserXy.class, fieldName, UserXy.class.getDeclaredField(fieldName).getType(), value);
                            /*
                            switch (j) {
                                case 0:
                                    userXy.setName(value);
                                    break;
                                case 1:
                                    userXy.setAge(value);
                                    break;
                                case 2:
                                    userXy.setPhone(value);
                                    break;
                                default:
                                    break;
                            }
                            */
                        } // 遍历每一列结束
                        dataList.add(userXy);
                    }
                } // 遍历每一行结束
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return dataList;
    }

}
