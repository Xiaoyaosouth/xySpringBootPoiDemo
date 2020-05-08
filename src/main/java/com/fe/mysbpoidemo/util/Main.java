package com.fe.mysbpoidemo.util;

import com.fe.mysbpoidemo.service.*;

import javax.swing.filechooser.FileSystemView;
import java.net.URL;
import java.net.URLEncoder;
import java.util.*;

public class Main {

    public static void main(String[] args) {
        System.out.println("Hello World!");
        try {
            // 获取桌面路径
            FileSystemView fsv = FileSystemView.getFileSystemView();
            String desktop = fsv.getHomeDirectory().getPath();

            /*
            String filePath1 = desktop + "\\temp\\testExcelData.xlsx";
            // 读取Excel文件内容
            ExcelReader.readExcel(filePath1);

            String filePath2 =desktop + "\\temp\\template.xls";
            ExcelReader.readExcel(filePath2);

            // 转成JSON写入文件
            ArrayList<String> strAryList1 = ExcelReader.convertExcelToJson(filePath1);
            FileUtil.writeTotxtFile(strAryList1, desktop + "\\temp\\testExcelData.json");
            ArrayList<String> strAryList2 = ExcelReader.convertExcelToJson(filePath2);
            FileUtil.writeTotxtFile(strAryList2, desktop + "\\temp\\template.json");
            */

            // String fileUrl = "http://xybucket.obs.cn-south-1.myhuaweicloud.com/enterpriseData.xlsx";
            // String fileUrl = "https://xybucket.obs.cn-south-1.myhwclouds.com/uistudio/upload/"
            //        + URLEncoder.encode("能源表基础信息导入模板 (1).xlsx", "utf-8");
            String fileUrl = "https://xybucket.obs.cn-south-1.myhuaweicloud.com/单元表装表导入模板（2）.xlsx";
            System.out.println("文件地址：" + fileUrl);
            DynamicExcelDataService excelDataService = new DynamicExcelDataService();
            List<Object> data = excelDataService.getDynamicModelListV2(fileUrl);
//            for (int i = 0; i < data.size(); i++) {
//                System.out.println(data.get(i).toString()); // 打印出的是地址
//            }


        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
