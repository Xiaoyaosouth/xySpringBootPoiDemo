package com.fe.mysbpoidemo.util;

import com.fe.mysbpoidemo.model.DynamicBean;
import com.fe.mysbpoidemo.model.ExcelCommon;
import com.fe.mysbpoidemo.model.ExcelModel;
import com.fe.mysbpoidemo.model.UserXy;
import com.fe.mysbpoidemo.service.CommonExcelReaderService;
import com.fe.mysbpoidemo.service.ExcelDataService;
import com.fe.mysbpoidemo.service.UserDataService;

import javax.swing.filechooser.FileSystemView;
import java.net.URL;
import java.util.List;

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

            /*
            String fileUrl = "http://xybucket.obs.cn-south-1.myhuaweicloud.com/userData.xlsx";
            UserDataService userDataService = new UserDataService();
            List<UserXy> data = userDataService.getModelList(fileUrl);
            for (int i = 0; i < data.size(); i++){
                System.out.println(data.get(i));
            }
            */

            /*
            String fileUrl = "http://xybucket.obs.cn-south-1.myhuaweicloud.com/userData.xlsx";
            List<String> data = ExcelReader.readExcelToJson(fileUrl);
            for (int i = 0; i < data.size(); i++){
                System.out.println(data.get(i));
            }
            */
            /*
            String fileUrl = "http://xybucket.obs.cn-south-1.myhuaweicloud.com/userData.xlsx";
            ExcelDataService excelDataService = new ExcelDataService();
            List<ExcelModel> data = excelDataService.getModelList(fileUrl);
            for (int i = 0; i < data.size(); i++){
                System.out.println(data.get(i));
            }
            */
            /*
            String fileUrl = "http://xybucket.obs.cn-south-1.myhuaweicloud.com/enterpriseData.xlsx";
            ExcelDataService excelDataService = new ExcelDataService();
            List<Object> data = excelDataService.getDynamicModelList(fileUrl);
            for (int i = 0; i < data.size(); i++){
                System.out.println(data.get(i));
            }
            */
            String fileUrl = "http://xybucket.obs.cn-south-1.myhuaweicloud.com/userData.xlsx";
            CommonExcelReaderService service = new CommonExcelReaderService();
            ExcelCommon data = service.getCommonExcelData(fileUrl);
            System.out.println(data.getProperty());
            System.out.println(data.getCellData());

        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
